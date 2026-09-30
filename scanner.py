"""
Voiant Sales Planning Signal & Account Intelligence
FREE Trigger Scanner (Script A)

WHAT THIS DOES
--------------
Checks two completely free, public sources per company for recent
activity that might signal a Section 5 trigger event -- WITHOUT using
the Anthropic API, without spending any tokens, without any cost at all:

  1. SEC EDGAR submissions API (by CIK) -- checks for recent 8-K filings
     (material events: exec departures/hires, restructuring) AND S-1 /
     S-1/A filings (IPO registration). Both are LEGALLY REQUIRED
     disclosures, so both are high-confidence structural signals, not
     inference.

  2. Google News RSS -- free, no API key, no scraping -- catches things
     that may not have (or need) an SEC filing: forecasting-miss
     language, territory/quota-miss commentary, rep attrition, layoffs,
     IPO news ahead of the S-1, and CFO/CRO change coverage.

WHY THIS VERSION IS DIFFERENT
-------------------------------
The previous version treated every keyword hit identically: 1 news hit
= "Moderate", 2+ = "Strong", regardless of WHAT was found. That doesn't
match Section 5's actual taxonomy, where some triggers are inherently
strong on their own (an IPO filing) and others only add up to something
meaningful when two DIFFERENT weak signals appear together (not two
articles about the same thing).

This version tags every hit with a specific trigger category (matching
Section 5's own categories) and classifies strength based on WHICH
categories were found, not just how many hits there were:
  - EDGAR 8-K or S-1/S-1-A -> Strong on its own (mandatory disclosure)
  - A single well-corroborated IPO/scaling news hit -> Strong on its own
  - A single CFO/CRO-change news hit (no EDGAR corroboration) -> Moderate
    only -- matches the same caution the paid script already applies
    (a leadership change alone isn't proof of a restructuring mandate)
  - 2+ DISTINCT weak categories (forecasting chaos, territory/quota miss,
    rep attrition, workforce reduction) -> Strong, mirroring the brief's
    "confirmed platform + two weak triggers" bar
  - Exactly 1 weak category -> Moderate



WHAT THIS STILL CANNOT CATCH (free-source limits, being upfront about it)
----------------------------------------------------------------------------
- Failed/incomplete SPM implementation signals: a repeatedly-reposted
  job listing (no free historical job-posting data source), negative
  G2/Gartner/TrustRadius reviews (no free API for these), a departed
  platform admin on LinkedIn (would require LinkedIn people-search,
  which this project deliberately avoids). These stay the paid script's
  job, via its case-study/job-posting research.
- Glassdoor-sourced attrition/comp complaints specifically -- no free
  API; the news-based "rep_attrition" category is a partial proxy only.

HOW THIS FITS THE WORKFLOW
----------------------------
Does NOT replace account_research_automation.py -- it cannot produce
Platform, Platform Confidence, Named Contact, Priority Score, etc. It
only decides which companies are worth a paid re-check sooner than the
normal cache window would trigger on its own.

OUTPUT:
  - TRIGGER_FLAGS_XLSX: every hit found, full detail, with its category
  - PRIORITY_REQUEUE_XLSX: only companies meeting REQUEUE_MIN_STRENGTH,
    ready to use directly as account_research_automation.py's input

SETUP
-----
1. pip install requests feedparser pandas openpyxl --break-system-packages
2. Edit SEC_USER_AGENT below to a real identifying contact (SEC's
   fair-access policy requires this).
3. Point INPUT_XLSX at your account list.
4. Run: python scanner.py
"""

import sys
import time
import logging
from pathlib import Path
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import requests
import feedparser
import pandas as pd

import xlsx_style
import report_export

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------

INPUT_XLSX = "test_40_accounts_READY.xlsx"
TRIGGER_FLAGS_XLSX = "trigger_flags.xlsx"
PRIORITY_REQUEUE_XLSX = "priority_requeue.xlsx"
REQUEUE_MIN_STRENGTH = "Strong"  # "Strong" or "Moderate"
SCAN_WINDOW_DAYS = 7

SEC_USER_AGENT = "Voiant Research contact@yourcompany.com"  # EDIT THIS

REQUEST_DELAY_SECONDS = 0.3
REQUEST_TIMEOUT_SECONDS = 15

# EDGAR forms to check -- 8-K for material events, S-1/S-1-A for IPO prep.
EDGAR_TARGET_FORMS = ("8-K", "S-1", "S-1/A")

# --------------------------------------------------------------------------
# TRIGGER CATEGORIES -- maps directly to Section 5's own taxonomy, plus the
# two additions Faheem asked for (new Sales Ops / RevOps leadership hires).
#
# ROLE_CHANGE_CATEGORIES need BOTH a role term AND an action verb in the
# SAME headline to count -- matching on the action verb alone (e.g.
# "appoints") used to categorize ANY leadership hire, including something
# totally unrelated like "Company appoints new VP of Marketing," as a
# CFO/CRO change. Requiring both together fixes that false-positive.
#
# "strength" for the other, non-role-change categories means: what does a
# single NEWS hit in this category earn on its own, before any EDGAR
# corroboration.
# --------------------------------------------------------------------------

ACTION_VERBS = [
    "steps down", "resigns", "departure", "departs",
    "names new", "appoints", "hires new",
]

ROLE_CHANGE_CATEGORIES = {
    "cfo_cro_change": {
        "strength": "moderate_alone",  # promoted to Strong only if EDGAR also hit
        "role_terms": ["cfo", "chief financial officer", "cro", "chief revenue officer"],
    },
    "new_sales_ops_revops_leadership": {
        "strength": "moderate_alone",  # VP/Director level rarely triggers an 8-K, so this
                                        # will almost always land Moderate via news alone --
                                        # that's expected, not a bug.
        "role_terms": [
            "vp of sales operations", "vp sales operations", "director of sales operations",
            "director sales operations", "head of sales operations",
            "vp of revenue operations", "vp revenue operations", "vp revops",
            "director of revenue operations", "director revops", "head of revenue operations",
        ],
    },
}

TRIGGER_CATEGORIES = {
    "ipo_scaling": {
        "strength": "strong",  # a single well-corroborated hit counts on its own
        "keywords": [
            "ipo", "initial public offering", "files for ipo", "goes public",
            "going public", "s-1 filing", "funding round", "series c",
            "series d", "series e", "raises $", "scaling rapidly",
        ],
    },
    "forecasting_chaos": {
        "strength": "weak",
        "keywords": [
            "forecasting challenges", "missed guidance", "cut guidance",
            "guidance cut", "lowered guidance", "cuts forecast",
        ],
    },
    "territory_quota_miss": {
        "strength": "weak",
        "keywords": [
            "territory realignment", "quota miss", "sales miss",
            "missed sales targets", "quota misalignment",
        ],
    },
    "rep_attrition": {
        "strength": "weak",
        "keywords": [
            "sales rep turnover", "sales team attrition", "high sales turnover",
            "sales reps leaving", "sales talent exodus",
        ],
    },
    "workforce_reduction": {
        "strength": "weak",
        "keywords": [
            "layoff", "layoffs", "job cuts", "workforce reduction",
        ],
    },
    # "restructuring"/"revamp" language alone doesn't cleanly mean any one
    # thing -- kept as its own honest bucket instead of folded into
    # cfo_cro_change, where it used to inflate that category regardless
    # of whether an exec change was even involved.
    "restructuring_language": {
        "strength": "weak",
        "keywords": ["restructuring", "restructure", "overhaul", "revamp"],
    },
}

# Flat list used for the actual RSS query -- broad net (role terms + action
# verbs + every other category's keywords), categorization/the AND-logic
# for role changes happens after a hit comes back, in categorize_title().
ALL_NEWS_KEYWORDS = (
    ACTION_VERBS
    + [term for cat in ROLE_CHANGE_CATEGORIES.values() for term in cat["role_terms"]]
    + [kw for cat in TRIGGER_CATEGORIES.values() for kw in cat["keywords"]]
)

_log_handlers = [logging.StreamHandler(sys.stdout)]
try:
    _log_handlers.insert(0, logging.FileHandler("scanner.log"))
except OSError:
    pass  # read-only filesystem (e.g. Vercel) -- stdout only is fine

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=_log_handlers,
)
log = logging.getLogger(__name__)


def categorize_title(title_lower):
    """Returns (category_name, matched_text) for a headline that already
    contains the company name, or (None, None) if nothing genuinely
    matches. Role-change categories (CFO/CRO change, new Sales Ops/RevOps
    leadership) require BOTH a role term AND an action verb in the SAME
    headline -- matching on the action verb alone used to categorize any
    unrelated hire/departure (e.g. a new VP of Marketing) as a CFO/CRO
    change, since "appoints" alone was one of that category's keywords."""
    matched_verb = next((v for v in ACTION_VERBS if v in title_lower), None)
    if matched_verb:
        for cat_name, cat in ROLE_CHANGE_CATEGORIES.items():
            role_term = next((r for r in cat["role_terms"] if r in title_lower), None)
            if role_term:
                return cat_name, f"{role_term} + {matched_verb}"

    for cat_name, cat in TRIGGER_CATEGORIES.items():
        matched = next((k for k in cat["keywords"] if k in title_lower), None)
        if matched:
            return cat_name, matched

    return None, None


# --------------------------------------------------------------------------
# SOURCE 1: SEC EDGAR submissions API, filtered by the company's own CIK
# --------------------------------------------------------------------------

_CIK_LOOKUP_CACHE = None


def _load_cik_lookup():
    global _CIK_LOOKUP_CACHE
    if _CIK_LOOKUP_CACHE is not None:
        return _CIK_LOOKUP_CACHE

    url = "https://www.sec.gov/files/company_tickers.json"
    headers = {"User-Agent": SEC_USER_AGENT}
    try:
        resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        log.warning(f"Could not download SEC company_tickers.json, EDGAR checks will be skipped this run: {e}")
        _CIK_LOOKUP_CACHE = {}
        return _CIK_LOOKUP_CACHE

    lookup = {}
    for entry in data.values():
        title = str(entry.get("title", "")).strip().lower()
        cik = str(entry.get("cik_str", "")).zfill(10)
        if title:
            lookup[title] = cik
    _CIK_LOOKUP_CACHE = lookup
    return lookup


def _resolve_cik(company_name):
    lookup = _load_cik_lookup()
    name_lower = company_name.strip().lower()

    if name_lower in lookup:
        return lookup[name_lower]

    candidates = [
        (title, cik) for title, cik in lookup.items()
        if name_lower in title or title in name_lower
    ]
    if len(candidates) == 1:
        return candidates[0][1]
    if len(candidates) > 1:
        candidates.sort(key=lambda pair: len(pair[0]))
        return candidates[0][1]
    return None


def check_edgar_filings(company_name, window_days, target_forms=EDGAR_TARGET_FORMS):
    """Returns a list of dicts: [{date, form, link, title, category}, ...]
    for filings actually FILED BY this company (by CIK) in the scan
    window, restricted to target_forms."""
    cik = _resolve_cik(company_name)
    if cik is None:
        log.warning(f"[{company_name}] Could not resolve a CIK -- skipping EDGAR check for this company.")
        return []

    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    headers = {"User-Agent": SEC_USER_AGENT}

    try:
        resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        log.warning(f"[{company_name}] EDGAR submissions lookup failed, skipping this source for this company: {e}")
        return []

    end_date = datetime.now(timezone.utc).date()
    start_date = end_date - timedelta(days=window_days)

    recent = data.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    filing_dates = recent.get("filingDate", [])
    accession_numbers = recent.get("accessionNumber", [])
    primary_docs = recent.get("primaryDocument", [])
    items_list = recent.get("items", [])  # e.g. "5.02,9.01" -- SEC's own item codes per 8-K

    hits = []
    for i, (form, filed_date, accession, primary_doc) in enumerate(zip(forms, filing_dates, accession_numbers, primary_docs)):
        if form not in target_forms:
            continue
        try:
            filed = datetime.strptime(filed_date, "%Y-%m-%d").date()
        except ValueError:
            continue
        if not (start_date <= filed <= end_date):
            continue

        accession_nodash = accession.replace("-", "")
        link = f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accession_nodash}/{primary_doc}"

        if form.startswith("S-1"):
            category = "ipo_scaling"
            title = f"{form} filed by {company_name} (CIK {cik})"
        elif form == "8-K":
            item_str = items_list[i] if i < len(items_list) else ""
            item_codes = [c.strip() for c in item_str.split(",") if c.strip()]
            # Item 5.02 = "Departure/Election/Appointment of Directors or
            # Officers" -- SEC's own official code for exactly the exec
            # change this category is meant to catch. Any OTHER 8-K item
            # (earnings, M&A, agreements, etc) is a real event but not
            # this specific trigger, so it gets its own honest bucket
            # instead of being lumped in as a false exec-change signal.
            if "5.02" in item_codes:
                category = "exec_change_8k"
                title = f"8-K Item 5.02 (Officer/Director Change) filed by {company_name} (CIK {cik})"
            else:
                category = "material_event_8k_other"
                title = f"8-K filed by {company_name} (CIK {cik}) -- item(s): {item_str or 'unspecified'}"
        else:
            category = "material_event_8k_other"
            title = f"{form} filed by {company_name} (CIK {cik})"

        hits.append({
            "date": filed_date,
            "form": form,
            "link": link,
            "title": title,
            "category": category,
        })

    return hits


# --------------------------------------------------------------------------
# SOURCE 2: Google News RSS (free, no key, built for syndication)
# --------------------------------------------------------------------------

def check_google_news(company_name, window_days, keywords):
    """Returns a list of dicts: [{date, title, link, matched_keyword,
    category}, ...] for news items in the scan window whose HEADLINE
    contains both the company name and one of our trigger keywords."""
    keyword_query = " OR ".join(f'"{k}"' for k in keywords)
    query = f'"{company_name}" ({keyword_query})'
    url = f"https://news.google.com/rss/search?q={quote(query)}&hl=en-US&gl=US&ceid=US:en"

    try:
        feed = feedparser.parse(url)
    except Exception as e:
        log.warning(f"[{company_name}] Google News RSS check failed: {e}")
        return []

    cutoff = datetime.now(timezone.utc) - timedelta(days=window_days)
    hits = []
    name_lower = company_name.strip().lower()

    for entry in feed.entries:
        published = getattr(entry, "published_parsed", None)
        if published is None:
            continue
        published_dt = datetime(*published[:6], tzinfo=timezone.utc)
        if published_dt < cutoff:
            continue

        title_lower = entry.title.lower()
        if name_lower not in title_lower:
            continue  # reject off-topic matches where the name appears elsewhere in the article, not the headline

        category, matched = categorize_title(title_lower)
        if category:
            hits.append({
                "date": published_dt.date().isoformat(),
                "title": entry.title,
                "link": entry.link,
                "matched_keyword": matched,
                "category": category,
            })

    return hits


# --------------------------------------------------------------------------
# MAIN SCAN LOOP
# --------------------------------------------------------------------------

def load_accounts(path):
    """Raises instead of sys.exit() -- only the CLI __main__ path below
    calls this today, but sys.exit() raising SystemExit (not a normal
    Exception) is a footgun for any future caller that wraps this in a
    try/except Exception expecting a clean error instead of a dead thread.
    See the same fix in Account_research_automation.py's load_accounts()
    for a case where that actually bit us."""
    if not Path(path).exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    if "Company" not in df.columns:
        raise ValueError(f"Input file must have a 'Company' column: {path}")

    dupes = df["Company"].astype(str).str.strip()
    dupe_names = dupes[dupes.duplicated()].unique().tolist()
    if dupe_names:
        log.warning(f"Duplicate company row(s) in {path}, keeping only the first occurrence of each: {dupe_names}")
        df = df.loc[~dupes.duplicated()].reset_index(drop=True)
    return df


def scan_company(company_name):
    """Returns a list of flag rows (dicts) for this company -- empty list
    means nothing new was found."""
    rows = []

    edgar_hits = check_edgar_filings(company_name, SCAN_WINDOW_DAYS)
    for hit in edgar_hits:
        rows.append({
            "Company": company_name,
            "Source Type": f"SEC EDGAR {hit['form']}",
            "Category": hit["category"],
            "What": hit["title"],
            "Date": hit["date"],
            "Source Link": hit["link"],
        })

    time.sleep(REQUEST_DELAY_SECONDS)

    news_hits = check_google_news(company_name, SCAN_WINDOW_DAYS, ALL_NEWS_KEYWORDS)
    for hit in news_hits:
        rows.append({
            "Company": company_name,
            "Source Type": f"Google News (matched: {hit['matched_keyword']})",
            "Category": hit["category"],
            "What": hit["title"],
            "Date": hit["date"],
            "Source Link": hit["link"],
        })

    time.sleep(REQUEST_DELAY_SECONDS)

    return rows


def classify_strength(company_flags):
    """Classifies by WHICH categories were found, not just hit count --
    matches Section 5's actual Strong/Weak structure rather than treating
    every hit identically.

    material_event_8k_other is deliberately never checked here -- it's
    still recorded in trigger_flags.xlsx for visibility, but a generic
    8-K (earnings, an agreement, an exhibit) isn't evidence of an exec
    change or any other defined trigger, so it doesn't score as one.
    Only an actual Item 5.02 filing (exec_change_8k) counts as Strong on
    EDGAR's own say-so."""
    has_edgar_exec_change = any(
        f["Source Type"].startswith("SEC EDGAR") and f["Category"] == "exec_change_8k"
        for f in company_flags
    )
    has_edgar_s1 = any(
        f["Source Type"].startswith("SEC EDGAR") and f["Category"] == "ipo_scaling"
        for f in company_flags
    )
    has_ipo_news = any(f["Category"] == "ipo_scaling" and f["Source Type"].startswith("Google News") for f in company_flags)

    # CFO/CRO change and new Sales Ops/RevOps leadership, found via news
    # only (no EDGAR corroboration) -- both land at Moderate, matching the
    # brief's existing caution that a leadership change alone isn't proof
    # of a restructuring mandate. VP/Director hires essentially never get
    # an EDGAR Item 5.02 filing (that's for officers/directors), so this
    # category will almost always land here, not get promoted to Strong --
    # that's expected, not a bug.
    role_change_news_categories = set(
        f["Category"] for f in company_flags
        if f["Category"] in ROLE_CHANGE_CATEGORIES and f["Source Type"].startswith("Google News")
    )

    weak_categories_found = set(
        f["Category"] for f in company_flags
        if f["Category"] in ("forecasting_chaos", "territory_quota_miss", "rep_attrition", "workforce_reduction", "restructuring_language")
    )

    if has_edgar_exec_change or has_edgar_s1 or has_ipo_news:
        return "Strong"
    if len(weak_categories_found) >= 2:
        return "Strong"  # mirrors "confirmed platform + two weak triggers" bar
    if role_change_news_categories or len(weak_categories_found) == 1:
        return "Moderate"
    return "None"


def main():
    if "contact@yourcompany.com" in SEC_USER_AGENT:
        log.warning(
            "SEC_USER_AGENT still has the placeholder email. Edit this "
            "before relying on the EDGAR results. Continuing anyway for this run."
        )

    try:
        accounts_df = load_accounts(INPUT_XLSX)
    except (FileNotFoundError, ValueError) as e:
        log.error(str(e))
        sys.exit(1)
    all_flags = []
    flags_by_company = {}

    total = len(accounts_df)
    for idx, row in accounts_df.iterrows():
        company = str(row["Company"]).strip()
        log.info(f"[{idx + 1}/{total}] Scanning {company}...")
        flags = scan_company(company)
        if flags:
            categories_found = sorted(set(f["Category"] for f in flags))
            log.info(f"  -> {len(flags)} item(s) found for {company}: {categories_found}")
            all_flags.extend(flags)
            flags_by_company[company] = flags
        else:
            log.info(f"  -> nothing new for {company}")

    if all_flags:
        flags_df = pd.DataFrame(all_flags)
        flags_df.to_excel(TRIGGER_FLAGS_XLSX, index=False)
        log.info(f"{len(all_flags)} flagged item(s) across {flags_df['Company'].nunique()} company/companies. Written to {TRIGGER_FLAGS_XLSX}")
    else:
        pd.DataFrame(columns=["Company", "Source Type", "Category", "What", "Date", "Source Link"]).to_excel(TRIGGER_FLAGS_XLSX, index=False)
        log.info(f"Nothing new found for any company this scan. Empty {TRIGGER_FLAGS_XLSX} written.")

    strength_by_company = {c: classify_strength(f) for c, f in flags_by_company.items()}
    accepted_strengths = {"Strong"} if REQUEUE_MIN_STRENGTH == "Strong" else {"Strong", "Moderate"}
    qualifying_companies = [c for c, s in strength_by_company.items() if s in accepted_strengths]

    if qualifying_companies:
        context_df = accounts_df[accounts_df["Company"].isin(qualifying_companies)].copy()
        context_df["Scanner Strength"] = context_df["Company"].map(strength_by_company)
        context_df["Scanner Categories"] = context_df["Company"].map(
            lambda c: ", ".join(sorted(set(f["Category"] for f in flags_by_company.get(c, []))))
        )
        context_df["Scanner Reason"] = context_df["Company"].map(
            lambda c: "\n".join(f"• {f['What']}" for f in flags_by_company.get(c, []))
        )
        log.info(
            f"{len(qualifying_companies)} account(s) meet the '{REQUEUE_MIN_STRENGTH}+' bar -- "
            f"written to {PRIORITY_REQUEUE_XLSX}."
        )
    else:
        context_df = pd.DataFrame(columns=list(accounts_df.columns) + ["Scanner Strength", "Scanner Categories", "Scanner Reason"])
        log.info(f"No account met the '{REQUEUE_MIN_STRENGTH}+' bar this scan. Empty {PRIORITY_REQUEUE_XLSX} written.")

    # Same Prospect List / Read Me / Reliability Key structure as the
    # post-research reference-format export -- Score/Band/Use
    # Case/Decision-Maker are genuinely blank pre-research, Status is
    # "Pending Research" (same convention the reference file already uses
    # for its own unresearched rows).
    report_export.write_scanner_reference_excel(context_df, PRIORITY_REQUEUE_XLSX)
    xlsx_style.style_xlsx(TRIGGER_FLAGS_XLSX)

    log.info("DONE.")


if __name__ == "__main__":
    main()