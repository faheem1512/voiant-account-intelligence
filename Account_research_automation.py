"""
Voiant Sales Planning Signal & Account Intelligence
Automated Account Research Script -- BATCH VERSION

WHAT CHANGED IN THIS VERSION
----------------------------
1. BATCHES API instead of one-at-a-time calls. Submits all accounts that
   need fresh research as a single Anthropic Message Batch. This cuts
   token + web search cost by 50% versus the standard API, at the
   tradeoff that results are NOT instant -- batches process
   asynchronously, this script polls until done.

2. CROSS-WEEK SMART CACHE. If you run this every Monday on the same
   60-account list, accounts already researched within the last
   CACHE_MAX_AGE_DAYS are reused automatically -- NO new API call, NO
   new tokens, NO new web searches spent on them. Only accounts that are
   new, previously errored, or past the freshness window get sent to the
   API.

   DESIGN CHOICE: platform/department findings are usually stable for
   weeks, but TRIGGER EVENTS go stale fast. Rather than re-checking only
   triggers, this caches the WHOLE result for CACHE_MAX_AGE_DAYS, then
   forces a full refresh after that. Simpler and auditable, at the cost
   of not catching a brand new trigger mid-window.

3. INTERRUPTION-SAFE BATCH SUBMISSION. If you Ctrl-C this script (or it
   crashes) while a batch is still processing, re-running it will NOT
   submit a duplicate batch and double-charge you -- it saves the
   in-flight batch ID to a small state file and resumes polling that
   instead.

CARRIED OVER FROM THE PRIOR VERSION
-------------------------------------
- LinkedIn URL lookup for an already-known name (not bulk people-search)
- Fixed CFO/CRO trigger rubric
- Section 7 priority scoring, computed in Python, not by the model
- Contact-finding stays on Option B; no LinkedIn scraping of any kind

SETUP
-----
1. pip install anthropic pandas openpyxl python-dotenv --break-system-packages
2. .env file with ANTHROPIC_API_KEY=sk-ant-... in this same folder
3. Point INPUT_XLSX at your account list
4. Run: python account_research_automation.py

NOTE ON SDK SURFACE
--------------------
The Batches API has moved in and out of "beta" naming across SDK
versions. This script tries client.messages.batches first and falls
back to client.beta.messages.batches if that doesn't exist. If BOTH
fail: pip install --upgrade anthropic
"""

import os
import sys
import json
import time
import logging
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Optional

import pandas as pd
from anthropic import Anthropic

import contacts_tracker

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

INPUT_XLSX = "priority_requeue.xlsx"
OUTPUT_XLSX = "Account_Research_Results.xlsx"
BATCH_STATE_FILE = "current_batch_id.txt"
MODEL = "claude-sonnet-5"
MAX_SEARCHES_PER_ACCOUNT = 10
CACHE_MAX_AGE_DAYS = 30
TRIGGER_FLAGS_XLSX = "trigger_flags.xlsx"  # written by the free scanner.py -- companies
                                            # listed here get refreshed regardless of
                                            # the CACHE_MAX_AGE_DAYS window
BATCH_POLL_INTERVAL_SECONDS = 30
BATCH_MAX_WAIT_SECONDS = 24 * 60 * 60

# Rough, deliberately conservative per-account cost estimate for the app's
# pre-confirm screen (Batches API discount + ~10 web searches + system
# prompt/output tokens on Sonnet). Not a billing guarantee -- actual varies
# with how much a given account's research turns up. Round up, not down.
ESTIMATED_COST_PER_ACCOUNT_USD = 0.35

_log_handlers = [logging.StreamHandler(sys.stdout)]
try:
    _log_handlers.insert(0, logging.FileHandler("account_research.log"))
except OSError:
    pass  # read-only filesystem (e.g. Vercel) -- stdout only is fine

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=_log_handlers,
)
log = logging.getLogger(__name__)

ICP_INDUSTRIES = {"tech & saas", "healthcare & medtech", "finserv & insurance"}

SYSTEM_PROMPT = """You are running account intelligence research for a B2B \
sales planning software consultancy (Voiant). Your job is to research ONE \
named company and determine whether it is a qualified sales target, \
following this exact method. Do not skip steps and do not guess -- if a \
search comes back empty, say so explicitly rather than inferring.

TARGET PLATFORMS (only these three count): Anaplan, Pigment, Varicent.
Ignore Excel-only shops, Xactly, and any other platform for platform-match \
purposes (you may mention them if directly relevant to a trigger, e.g. a \
migration away from Xactly toward one of the three).

STEP 1 -- PEOPLE SEARCH
Search for public mentions of the company alongside "Anaplan", "Pigment", \
or "Varicent" plus terms like "Sales Ops", "Model Builder", "Sales \
Compensation". Do NOT attempt to access LinkedIn profile pages or scrape \
LinkedIn search results, and do NOT run any bulk/systematic LinkedIn \
people-search -- only use what surfaces through general web search of \
publicly indexed content (case studies, press, job boards).

If a named individual with a title appears in a public case study, \
webinar, or press mention, capture their name and title. Then, and ONLY \
then (i.e. only once you already have a specific real name from a \
public source), run ONE additional targeted search of the form: \
"[Full Name]" "[Company]" LinkedIn -- to try to surface their public, \
already-indexed LinkedIn profile URL. This is a lookup of a single \
already-known individual, equivalent to a person googling a name they \
already have -- it is NOT a bulk people-search and must never be used to \
go hunting for names you don't already have from the public-source \
requirement above.

If that LinkedIn URL search succeeds, include the URL. If it does NOT \
surface a confident match, leave the URL field blank -- but ALWAYS still \
return the name and title you found from the original public source. A \
name without a confirmed LinkedIn URL is still useful for manual \
follow-up; never discard a genuinely-found name just because the URL \
lookup didn't pan out. Do not invent a name, title, or URL under any \
circumstances -- if this step found nothing, leave all three fields \
blank.

STEP 2 -- JOB POSTINGS
Search for direct roles (e.g. "[Company] Anaplan admin", "[Company] \
Pigment builder", "[Company] Varicent analyst") and adjacent roles (Sales \
Comp Analyst, RevOps Manager, Sales Planning Lead) that name these tools. \
Note explicitly which DEPARTMENT the posting sits under (Finance/FP&A vs \
Sales Ops/RevOps/Sales Comp) -- this distinction is the whole point of \
the exercise. A tool confirmed under a Finance/FP&A posting does NOT \
count as sales-side.

STEP 3 -- PARTNER / CASE STUDY / CONFERENCE MENTIONS
Search for vendor customer-story pages, conference speaker lists (Anaplan \
CPX, Pigment events, Varicent Accelerate), and testimonial webinars \
naming the company.

STEP 4 -- TRIGGER EVENTS
Only search this step if Steps 1-3 produced ANY platform signal (strong \
or weak). Check for:
  STRONG (publicly detectable): new CRO/CFO hire wanting to restructure \
    planning; IPO prep or scaling threshold forcing a move off Excel; \
    failed/incomplete SPM implementation (repeated job reposting, \
    negative G2/Gartner/TrustRadius reviews, departed platform admin).
  WEAK (inference only, never qualifies alone): inaccurate forecast / \
    chaotic planning language in earnings calls or Glassdoor; sales miss \
    attributed to territory/quota misalignment; rep attrition spike tied \
    to comp or quota confidence.

IMPORTANT -- a new CFO/CRO hire is NOT automatically a Strong trigger. \
It only qualifies if there is actual evidence (a quote, an announcement, \
an analyst note) that this person specifically intends to change/rebuild \
sales planning, forecasting, quota-setting, or comp structures. A CFO \
hired for general cost containment, financial recalibration, M&A, or \
"operational efficiency" WITHOUT any stated planning/forecasting/quota/ \
comp angle does NOT qualify as Strong -- downgrade it to a Weak trigger \
at most (only if there's separate evidence fitting the Weak category), \
or None if there isn't. Be skeptical of your own urge to count a \
leadership change as a trigger just because it's dated and convenient -- \
the bar is a stated connection to planning, not just a new name in a \
C-suite seat.

A trigger found must be dated and sourced. If a CFO/CRO change is more \
than ~12 months old, treat it as stale, not a current strong trigger, \
regardless of mandate.

Also count how many DISTINCT weak triggers you found (0, 1, or 2 -- cap \
at 2 even if more appear) and, separately, how recent the single \
strongest trigger is (whether Strong or Weak): "<3mo", "3-12mo", ">12mo", \
or "N/A" if no trigger was found at all. This feeds the Priority Score \
formula downstream -- be precise, this is not just narrative color.

SCORING RUBRIC (apply exactly, do not soften):
Platform Confidence:
  - Confirmed: tool is named in a source AND that source (or a separate \
    source) confirms the department is Sales-Ops/RevOps/Sales Comp, not \
    Finance-only.
  - Inferred: tool is named, but department is unclear, unconfirmed, or \
    appears Finance-led.
  - Unconfirmed: no credible mention of any of the three platforms tied \
    to this company at all.

Trigger Strength: "Strong" only for the three strong-category triggers \
above with a clear date AND (for CFO/CRO hires specifically) explicit \
evidence of a planning-related mandate, per the rule above. "Weak" for \
the three weak-category triggers. "None" if nothing found.

Outreach Ready = TRUE only if: (one Strong trigger) OR (Platform \
Confidence = Confirmed AND two or more Weak triggers). Confirmed \
platform alone, with no triggers, is NOT Outreach Ready.

IMPORTANT HONESTY RULES
- If you find nothing, say "None found" -- do not pad the answer.
- If a piece of evidence is ambiguous (e.g. an aggregated job board \
  listing that might not actually belong to this company), flag it as \
  low confidence rather than treating it as confirmed. Note this in \
  low_confidence_flag with a brief reason.
- Never fabricate a named contact, title, or LinkedIn URL. Leave the \
  relevant field blank if it wasn't genuinely found in public content.

TRIGGER_EVENT FORMAT -- Faheem (the reviewer) said long paragraphs here \
were hard to scan quickly. Write trigger_event as 2-3 short bullet-point \
clauses, each its own line starting with "• ", not one long paragraph. \
Each clause should be a single plain fact (what happened, who, roughly \
when) -- save the reasoning/caveats for low_confidence_flag or \
research_summary instead of packing them into this field. If there's only \
one fact, one bullet is fine -- never pad to hit 2-3 lines artificially.
Example:
"• New CFO (Jane Smith) hired March 2026, replacing prior CFO.\n• Smith's \
appointment announcement specifically cites 'rebuilding forecasting \
discipline' as a mandate.\n• No public statement yet on tooling changes."

OUTPUT FORMAT
Respond with ONLY a single JSON object (no markdown fences, no preamble, \
no commentary before or after), matching exactly this schema:

{
  "company": "string",
  "platform": "Anaplan | Pigment | Varicent | None found",
  "platform_confidence": "Confirmed | Inferred | Unconfirmed",
  "department": "string describing what was found, or 'N/A'",
  "named_contact": "string or empty string if none found",
  "contact_title": "string or empty string",
  "contact_linkedin_url": "string or empty string if not found -- see Step 1 rules",
  "trigger_event": "2-3 short bullet lines (each starting with '• ', joined with \\n), or 'None identified'",
  "trigger_strength": "Strong | Weak | None",
  "trigger_source_and_date": "string or empty string",
  "weak_trigger_count": 0 or 1 or 2,
  "trigger_recency_bucket": "<3mo | 3-12mo | >12mo | N/A",
  "outreach_ready": true or false,
  "low_confidence_flag": "string reason if any finding is shaky, else empty string",
  "source_links": "semicolon-separated list of URLs used",
  "research_summary": "2-3 sentence plain-language summary of what was found"
}
"""

USER_PROMPT_TEMPLATE = """Research this company following the method and \
rubric exactly as specified in your instructions:

Company: {company}
ICP Industry (if known): {icp_industry}
Headquarters (if known): {hq}

Return only the JSON object described in your instructions."""


@dataclass
class AccountResult:
    company: str
    platform: str = ""
    platform_confidence: str = ""
    department: str = ""
    named_contact: str = ""
    contact_title: str = ""
    contact_linkedin_url: str = ""
    trigger_event: str = ""
    trigger_strength: str = ""
    trigger_source_and_date: str = ""
    weak_trigger_count: int = 0
    trigger_recency_bucket: str = ""
    outreach_ready: bool = False
    low_confidence_flag: str = ""
    source_links: str = ""
    research_summary: str = ""
    needs_manual_review: bool = False
    error: str = ""
    priority_score: Optional[float] = None
    priority_band: str = ""
    score_notes: str = ""
    last_researched: str = ""
    from_cache: bool = False


def _trigger_points(result):
    if result.trigger_strength == "Strong":
        return 9.0
    if result.trigger_strength == "Weak":
        return 6.0 if result.weak_trigger_count >= 2 else 3.0
    return 0.0


def _platform_points(result):
    return {"Confirmed": 10.0, "Inferred": 6.0, "Unconfirmed": 2.0}.get(result.platform_confidence, 2.0)


def _size_points(employees):
    if employees is None:
        return 4.0, True
    if employees >= 10000:
        return 10.0, False
    if employees >= 5000:
        return 8.0, False
    if employees >= 1000:
        return 6.0, False
    return 4.0, False


def _industry_points(icp_industry):
    return 9.0 if icp_industry.strip().lower() in ICP_INDUSTRIES else 3.0


def _recency_points(bucket):
    return {"<3mo": 10.0, "3-12mo": 6.0, ">12mo": 3.0, "N/A": 0.0}.get(bucket, 0.0)


def _band(score):
    if score >= 8:
        return "A"
    if score >= 6:
        return "B"
    if score >= 4:
        return "C"
    return "D"


def compute_priority_score(result, employees, icp_industry):
    size_pts, size_unknown = _size_points(employees)
    weighted = (
        _trigger_points(result) * 0.30
        + _platform_points(result) * 0.25
        + size_pts * 0.20
        + _industry_points(icp_industry) * 0.15
        + _recency_points(result.trigger_recency_bucket) * 0.10
    )
    result.priority_score = round(weighted, 2)
    result.priority_band = _band(weighted)
    if size_unknown:
        note = "Employees unknown -- scored at lowest size bucket (4pts); add real headcount and re-score."
        result.score_notes = note
        result.needs_manual_review = True
        result.low_confidence_flag = f"{result.low_confidence_flag} | {note}" if result.low_confidence_flag else note


class MissingAPIKeyError(RuntimeError):
    pass


def build_client():
    """Raises MissingAPIKeyError rather than calling sys.exit() -- this is
    called both from the CLI path (main(), which is fine exiting the
    process) and from app.py's background threads / request handlers,
    where sys.exit() only kills that one thread silently and leaves the
    UI's status stuck on 'running' forever with no error ever shown. Raise
    a normal exception and let each caller decide how to handle it."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise MissingAPIKeyError(
            "ANTHROPIC_API_KEY is not set. Create a '.env' file in this "
            "folder containing 'ANTHROPIC_API_KEY=sk-ant-...', or export "
            "it in the same terminal session you run this from."
        )
    return Anthropic(api_key=api_key)


def parse_json_response(raw_text, company):
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.strip("`")
        if cleaned.lower().startswith("json"):
            cleaned = cleaned[4:]
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        log.error(f"[{company}] Failed to parse JSON response: {e}")
        log.error(f"[{company}] Raw response was:\n{raw_text[:2000]}")
        return None


def apply_parsed_result(result, parsed):
    result.platform = parsed.get("platform", "")
    result.platform_confidence = parsed.get("platform_confidence", "")
    result.department = parsed.get("department", "")
    result.named_contact = parsed.get("named_contact", "")
    result.contact_title = parsed.get("contact_title", "")
    result.contact_linkedin_url = parsed.get("contact_linkedin_url", "")
    result.trigger_event = parsed.get("trigger_event", "")
    result.trigger_strength = parsed.get("trigger_strength", "")
    result.trigger_source_and_date = parsed.get("trigger_source_and_date", "")
    result.weak_trigger_count = int(parsed.get("weak_trigger_count", 0) or 0)
    result.trigger_recency_bucket = parsed.get("trigger_recency_bucket", "N/A")
    result.outreach_ready = bool(parsed.get("outreach_ready", False))
    result.low_confidence_flag = parsed.get("low_confidence_flag", "")
    result.source_links = parsed.get("source_links", "")
    result.research_summary = parsed.get("research_summary", "")
    result.needs_manual_review = bool(result.low_confidence_flag)

    if result.named_contact and not result.contact_linkedin_url:
        note = "Name found, no LinkedIn URL confirmed -- search manually."
        result.low_confidence_flag = f"{result.low_confidence_flag} | {note}" if result.low_confidence_flag else note
        result.needs_manual_review = True


def _batches_client(client):
    if hasattr(client, "messages") and hasattr(client.messages, "batches"):
        return client.messages.batches
    if hasattr(client, "beta") and hasattr(client.beta, "messages") and hasattr(client.beta.messages, "batches"):
        log.info("Using client.beta.messages.batches (older SDK path).")
        return client.beta.messages.batches
    raise RuntimeError(
        "Could not find a Batches API on this anthropic SDK install. "
        "Run: pip install --upgrade anthropic"
    )


def build_batch_requests(accounts_to_research):
    from anthropic.types.message_create_params import MessageCreateParamsNonStreaming
    from anthropic.types.messages.batch_create_params import Request

    requests = []
    for custom_id, company, icp_industry, hq in accounts_to_research:
        user_prompt = USER_PROMPT_TEMPLATE.format(
            company=company, icp_industry=icp_industry or "unknown", hq=hq or "unknown"
        )
        requests.append(
            Request(
                custom_id=custom_id,
                params=MessageCreateParamsNonStreaming(
                    model=MODEL,
                    max_tokens=4000,
                    system=SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_prompt}],
                    tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": MAX_SEARCHES_PER_ACCOUNT}],
                ),
            )
        )
    return requests


def submit_or_resume_batch(client, accounts_to_research):
    batches = _batches_client(client)

    if Path(BATCH_STATE_FILE).exists():
        old_batch_id = Path(BATCH_STATE_FILE).read_text().strip()
        if old_batch_id:
            try:
                existing = batches.retrieve(old_batch_id)
                if getattr(existing, "processing_status", None) in ("in_progress", "ended", "canceling"):
                    log.info(
                        f"Found an in-flight/completed batch from a previous run "
                        f"({old_batch_id}, status={existing.processing_status}). "
                        f"Resuming that instead of submitting a new one."
                    )
                    return existing
            except Exception as e:
                log.warning(f"Could not resume saved batch {old_batch_id}: {e}. Submitting a new batch.")

    log.info(f"Submitting a new batch of {len(accounts_to_research)} account(s)...")
    requests = build_batch_requests(accounts_to_research)
    batch = batches.create(requests=requests)
    Path(BATCH_STATE_FILE).write_text(batch.id)
    log.info(f"Batch submitted: {batch.id}")
    return batch


def poll_batch_until_done(client, batch):
    batches = _batches_client(client)
    waited = 0
    while True:
        batch = batches.retrieve(batch.id)
        status = getattr(batch, "processing_status", "unknown")
        log.info(f"Batch {batch.id} status: {status} (waited {waited}s)")
        if status == "ended":
            break
        if waited >= BATCH_MAX_WAIT_SECONDS:
            # Raise, don't sys.exit() -- this runs inside app.py's background
            # research thread too, where SystemExit would silently kill the
            # thread and leave the UI stuck on "running" forever with the
            # batch ID already saved to BATCH_STATE_FILE, unable to start a
            # new run. Raising lets _run_research_job's except report a
            # clean error, and the batch ID file is untouched either way so
            # rerunning still resumes polling instead of resubmitting.
            raise TimeoutError(f"Batch did not finish within {BATCH_MAX_WAIT_SECONDS}s safety ceiling. Rerun to keep polling this same batch.")
        time.sleep(BATCH_POLL_INTERVAL_SECONDS)
        waited += BATCH_POLL_INTERVAL_SECONDS
    if Path(BATCH_STATE_FILE).exists():
        Path(BATCH_STATE_FILE).unlink()
    return batch


def fetch_batch_results(client, batch, id_to_company):
    batches = _batches_client(client)
    results = {}
    today = datetime.now(timezone.utc).date().isoformat()

    for entry in batches.results(batch.id):
        custom_id = entry.custom_id
        company, icp_industry, hq = id_to_company.get(custom_id, (custom_id, "", ""))
        result = AccountResult(company=company, last_researched=today, from_cache=False)

        entry_type = getattr(entry.result, "type", None)
        if entry_type == "succeeded":
            message = entry.result.message
            text_parts = [b.text for b in message.content if getattr(b, "type", None) == "text"]
            raw_text = "\n".join(text_parts)
            if not raw_text.strip():
                result.error = "Empty response from model (no text content) in batch result"
                result.needs_manual_review = True
            else:
                parsed = parse_json_response(raw_text, company)
                if parsed is None:
                    result.error = "JSON parse failure in batch result"
                    result.needs_manual_review = True
                else:
                    apply_parsed_result(result, parsed)
        elif entry_type == "errored":
            error_info = getattr(entry.result, "error", None)
            result.error = f"Batch request errored: {error_info}"
            result.needs_manual_review = True
        elif entry_type == "expired":
            result.error = "Batch request expired before completion"
            result.needs_manual_review = True
        else:
            result.error = f"Unrecognized batch result type: {entry_type}"
            result.needs_manual_review = True

        results[company] = result

    return results


def load_accounts(path):
    """Raises instead of sys.exit() -- this is called from plan_research(),
    which app.py's /api/research/preview route calls SYNCHRONOUSLY inside a
    request handler (not a background thread). A sys.exit() here raises
    SystemExit, which is not an Exception subclass, so it blew straight
    through the route's `except Exception` and broke the connection
    outright instead of returning a clean error -- confirmed by testing
    before this fix. Raising a normal exception fixes both this route and
    the background-thread callers in one place."""
    if not Path(path).exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    # priority_requeue.xlsx now ships as a 3-sheet workbook (Prospect List /
    # Read Me / Reliability Key) so Faheem sees the same structure whether
    # an account's been researched yet or not -- target that sheet by name
    # explicitly rather than assuming sheet order. Single-sheet inputs
    # (the master list, old-style files) are unaffected.
    xls = pd.ExcelFile(path)
    sheet_name = "Prospect List" if "Prospect List" in xls.sheet_names else 0
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    if "Company" not in df.columns:
        raise ValueError(f"Input file must have a 'Company' column: {path}")
    return df


def _clean(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value)


def load_flagged_companies(path):
    """Reads the free scanner's output file (if it exists) and returns the
    set of company names it flagged as having new activity. These
    override the normal cache freshness window -- flagged means
    re-research regardless of how recently it was last checked."""
    if not Path(path).exists():
        return set()
    try:
        df = pd.read_excel(path)
    except Exception as e:
        log.warning(f"Could not read {path} (scanner flags): {e}")
        return set()
    if "Company" not in df.columns or df.empty:
        return set()
    flagged = set(str(c).strip() for c in df["Company"].dropna())
    return flagged


def load_cached_results(path):
    if not Path(path).exists():
        return {}, {}

    try:
        df = pd.read_excel(path)
    except Exception as e:
        log.warning(f"Could not read existing {path} for cache check: {e}")
        return {}, {}

    fresh, stale = {}, {}
    cutoff = datetime.now(timezone.utc).date() - timedelta(days=CACHE_MAX_AGE_DAYS)

    for _, row in df.iterrows():
        company = str(row.get("Company", "")).strip()
        if not company:
            continue

        error_val = row.get("Error", "")
        had_error = pd.notna(error_val) and str(error_val).strip()

        last_researched_raw = row.get("Last Researched", "")
        last_researched_date = None
        if pd.notna(last_researched_raw) and str(last_researched_raw).strip():
            try:
                last_researched_date = datetime.fromisoformat(str(last_researched_raw)[:10]).date()
            except ValueError:
                last_researched_date = None

        result = AccountResult(company=company)
        result.platform = _clean(row.get("Platform (Anaplan/Pigment/Varicent)"))
        result.platform_confidence = _clean(row.get("Platform Confidence (Confirmed/Inferred/Unconfirmed)"))
        result.department = _clean(row.get("Department Confirmed? (Sales-Ops/RevOps/Finance-only/Unknown)"))
        result.named_contact = _clean(row.get("Named Contact"))
        result.contact_title = _clean(row.get("Contact Title"))
        result.contact_linkedin_url = _clean(row.get("Contact LinkedIn URL"))
        result.trigger_event = _clean(row.get("Trigger Event"))
        result.trigger_strength = _clean(row.get("Trigger Strength (Strong/Weak)"))
        result.trigger_source_and_date = _clean(row.get("Trigger Source/Date"))
        result.weak_trigger_count = int(row.get("Weak Trigger Count") or 0)
        result.trigger_recency_bucket = _clean(row.get("Trigger Recency")) or "N/A"
        result.outreach_ready = bool(row.get("Outreach Ready", False))
        result.source_links = _clean(row.get("Source Links"))
        result.research_summary = _clean(row.get("Research Summary"))
        result.low_confidence_flag = _clean(row.get("Low Confidence Reason"))
        result.needs_manual_review = bool(row.get("NEEDS_MANUAL_REVIEW", False))
        result.last_researched = _clean(last_researched_raw)

        is_fresh = (not had_error) and (last_researched_date is not None) and (last_researched_date >= cutoff)
        if is_fresh:
            result.from_cache = True
            fresh[company] = result
        else:
            stale[company] = result

    return fresh, stale


def results_to_dataframe(results):
    rows = []
    for r in results:
        rows.append({
            "Company": r.company,
            "Platform (Anaplan/Pigment/Varicent)": r.platform,
            "Platform Confidence (Confirmed/Inferred/Unconfirmed)": r.platform_confidence,
            "Department Confirmed? (Sales-Ops/RevOps/Finance-only/Unknown)": r.department,
            "Named Contact": r.named_contact,
            "Contact Title": r.contact_title,
            "Contact LinkedIn URL": r.contact_linkedin_url,
            "Trigger Event": r.trigger_event,
            "Trigger Strength (Strong/Weak)": r.trigger_strength,
            "Trigger Source/Date": r.trigger_source_and_date,
            "Weak Trigger Count": r.weak_trigger_count,
            "Trigger Recency": r.trigger_recency_bucket,
            "Outreach Ready": r.outreach_ready,
            "Priority Score": r.priority_score,
            "Priority Band": r.priority_band,
            "Source Links": r.source_links,
            "Research Summary": r.research_summary,
            "NEEDS_MANUAL_REVIEW": r.needs_manual_review,
            "Low Confidence Reason": r.low_confidence_flag,
            "Score Notes": r.score_notes,
            "Error": r.error,
            "Last Researched": r.last_researched,
            "From Cache (No New API Call)": r.from_cache,
            "Status": "Automated - Needs Review" if r.needs_manual_review else "Automated - Researched",
        })
    return pd.DataFrame(rows)


def _load_account_context(input_xlsx):
    """Shared by plan_research() and run() so the preview and the real run
    can never disagree about who's in scope."""
    accounts_df = load_accounts(input_xlsx)
    account_context = {}
    for _, row in accounts_df.iterrows():
        company = str(row["Company"]).strip()
        icp_industry_raw = row.get("ICP_Bucket", row.get("ICP Industry", row.get("Industry", None)))
        icp_industry = str(icp_industry_raw) if pd.notna(icp_industry_raw) else ""
        hq = str(row.get("Headquarters", row.get("HQ", ""))) if pd.notna(row.get("Headquarters", row.get("HQ", None))) else ""
        employees_raw = row.get("Employees", row.get("Est. Employees", None))
        employees = float(employees_raw) if pd.notna(employees_raw) else None
        if company in account_context:
            # Dict-keyed-by-company already prevented a duplicate ROW from
            # reaching the output (last one wins), but that was happening
            # silently -- log it so a genuinely duplicated/misspelled input
            # row gets noticed and fixed at the source instead of quietly
            # masked.
            log.warning(f"'{company}' appears more than once in {input_xlsx} -- using the last occurrence, ignoring the earlier one(s).")
        account_context[company] = (icp_industry, hq, employees)
    return account_context


def plan_research(input_xlsx=None):
    """Preview-only: figures out who WOULD be sent to a fresh batch without
    submitting anything or spending a cent. This is what the app's
    pre-confirm screen calls -- 'N companies need research, ~$X estimated'
    -- before the user commits to a real (billed) run."""
    input_xlsx = input_xlsx or INPUT_XLSX
    account_context = _load_account_context(input_xlsx)

    fresh_cache, stale_cache = load_cached_results(OUTPUT_XLSX)
    flagged_companies = load_flagged_companies(TRIGGER_FLAGS_XLSX)

    fresh_in_scope = {c for c in account_context if c in fresh_cache}
    bumped_by_scanner = flagged_companies & fresh_in_scope

    needs_research = sorted(
        c for c in account_context
        if c not in fresh_cache or c in bumped_by_scanner
    )

    return {
        "total_in_scope": len(account_context),
        "needs_research": needs_research,
        "needs_research_count": len(needs_research),
        "reused_from_cache_count": len(account_context) - len(needs_research),
        "estimated_cost_usd": round(len(needs_research) * ESTIMATED_COST_PER_ACCOUNT_USD, 2),
        "cost_per_account_usd": ESTIMATED_COST_PER_ACCOUNT_USD,
    }


def run(input_xlsx=None):
    """The real, billed run. Same logic main() always had -- just callable
    with a return value instead of only usable as a CLI script, so the app
    can invoke it directly and report a structured summary back to the UI."""
    input_xlsx = input_xlsx or INPUT_XLSX
    client = build_client()
    account_context = _load_account_context(input_xlsx)

    fresh_cache, stale_cache = load_cached_results(OUTPUT_XLSX)
    flagged_companies = load_flagged_companies(TRIGGER_FLAGS_XLSX)

    if flagged_companies:
        bumped = flagged_companies & set(fresh_cache.keys())
        if bumped:
            log.info(
                f"Scanner flagged {len(bumped)} account(s) with new activity "
                f"-- overriding cache, will re-research even though still fresh: {sorted(bumped)}"
            )
            for company in bumped:
                del fresh_cache[company]

    if fresh_cache:
        log.info(f"{len(fresh_cache)} account(s) are within the {CACHE_MAX_AGE_DAYS}-day freshness window -- reusing, no API call, no cost.")
    if stale_cache:
        log.info(f"{len(stale_cache)} account(s) have prior results but are stale/errored -- will be re-researched.")

    needs_research = [c for c in account_context if c not in fresh_cache]

    id_to_company = {}
    batch_requests_input = []
    for i, company in enumerate(needs_research):
        custom_id = f"acct_{i}"
        icp_industry, hq, _ = account_context[company]
        id_to_company[custom_id] = (company, icp_industry, hq)
        batch_requests_input.append((custom_id, company, icp_industry, hq))

    fresh_results = {}
    if batch_requests_input:
        batch = submit_or_resume_batch(client, batch_requests_input)
        batch = poll_batch_until_done(client, batch)
        fresh_results = fetch_batch_results(client, batch, id_to_company)
    else:
        log.info("Nothing needs fresh research this run -- every account was cached.")

    all_results = []
    for company in account_context:
        icp_industry, hq, employees = account_context[company]
        if company in fresh_cache:
            result = fresh_cache[company]
        elif company in fresh_results:
            result = fresh_results[company]
        else:
            result = AccountResult(company=company, error="No result produced (missing from batch output)", needs_manual_review=True)

        compute_priority_score(result, employees, icp_industry)
        all_results.append(result)

    final_df = results_to_dataframe(all_results)
    final_df.to_excel(OUTPUT_XLSX, index=False)

    new_contacts_added = contacts_tracker.update_tracker(all_results)

    needs_review = int(final_df["NEEDS_MANUAL_REVIEW"].sum())
    outreach_ready = int(final_df["Outreach Ready"].sum())
    from_cache_count = int(final_df["From Cache (No New API Call)"].sum())
    log.info(
        f"DONE. {len(all_results)} accounts total. {from_cache_count} reused from cache "
        f"(zero new cost). {len(needs_research)} freshly researched via batch. "
        f"{outreach_ready} Outreach Ready. {needs_review} flagged NEEDS_MANUAL_REVIEW. "
        f"{new_contacts_added} new named contact(s) added to {contacts_tracker.CONTACTS_TRACKER_XLSX}."
    )
    log.info(f"Results written to {OUTPUT_XLSX}")

    return {
        "total_accounts": len(all_results),
        "from_cache_count": from_cache_count,
        "freshly_researched_count": len(needs_research),
        "new_contacts_added": new_contacts_added,
        "outreach_ready_count": outreach_ready,
        "needs_review_count": needs_review,
        "output_path": OUTPUT_XLSX,
    }


def current_batch_status():
    """Non-blocking peek at whatever batch is currently in flight (if any),
    for the app's status endpoint to poll without re-running the whole
    poll_batch_until_done() loop itself. Returns None if there's no
    in-flight batch recorded."""
    if not Path(BATCH_STATE_FILE).exists():
        return None
    batch_id = Path(BATCH_STATE_FILE).read_text().strip()
    if not batch_id:
        return None
    try:
        client = build_client()
        batches = _batches_client(client)
        batch = batches.retrieve(batch_id)
        return {"batch_id": batch_id, "processing_status": getattr(batch, "processing_status", "unknown")}
    except Exception as e:
        return {"batch_id": batch_id, "processing_status": "unknown", "error": str(e)}


def main():
    try:
        run()
    except (MissingAPIKeyError, FileNotFoundError, ValueError, TimeoutError) as e:
        log.error(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()