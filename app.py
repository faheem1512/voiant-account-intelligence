"""
Voiant Prospect Intelligence -- local web app (per Voiant_App_Technical_Spec)

Two buttons for Faheem, no terminal, no VS Code:
  1. "Run Scanner"           -- free, all ~595 companies, concurrent (8 workers)
  2. "Run Account Research"  -- paid (Batches API), cost estimate + confirm gate first

This file WRAPS scanner.py and Account_research_automation.py. It does not
reimplement their research logic, scoring formula, trigger rubric, or
Sales-vs-Finance classification -- those are imported and called as-is,
per the build spec's explicit "do not rebuild the research logic" and
"no changes to the scoring formula" instructions.

Run: python app.py
Then open http://localhost:5057 in a browser.
"""

import os
import shutil
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template_string, request, send_file

import scanner
import Account_research_automation as ara
import contacts_tracker
import report_export
import xlsx_style
from industry_mapping import to_icp_bucket

app = Flask(__name__)

MASTER_LIST_XLSX = "Forbes_2000_master.xlsx"
SCANNER_WORKERS = 8

DEFAULT_SCANNER_CHUNK_SIZE = 100

_state_lock = threading.Lock()
scanner_state = {
    "status": "idle", "scanned": 0, "total": 0, "flagged_count": 0, "error": None, "finished_at": None,
    "chunk_size": DEFAULT_SCANNER_CHUNK_SIZE, "chunks_done": 0, "chunks_total": 0,
}
research_state = {"status": "idle", "result": None, "error": None}

_scanner_companies = []
_scanner_master_df = None
_scanner_next_index = 0
_scanner_flags_by_company = {}
_scanner_all_flags = []


def _tmp(filename):
    """Write output files to /tmp on Vercel (read-only root fs), local dir otherwise."""
    if os.environ.get("VERCEL"):
        return f"/tmp/{filename}"
    return filename


def load_enriched_master():
    df = pd.read_excel(MASTER_LIST_XLSX)
    df.columns = [str(c).strip() for c in df.columns]

    companies = df["Company"].astype(str).str.strip()
    dupe_names = companies[companies.duplicated()].unique().tolist()
    if dupe_names:
        scanner.log.warning(f"Duplicate company row(s) in {MASTER_LIST_XLSX}, keeping only the first occurrence of each: {dupe_names}")
        df = df.loc[~companies.duplicated()].reset_index(drop=True)
        companies = df["Company"].astype(str).str.strip()

    out = pd.DataFrame({
        "Company": companies,
        "ICP_Bucket": df["Industry"].apply(to_icp_bucket) if "Industry" in df.columns else "",
        "Headquarters": df["Headquarters"] if "Headquarters" in df.columns else "",
    })
    out["Employees"] = None
    return out


def _write_scanner_outputs(master_df, flags_by_company, all_flags):
    if all_flags:
        pd.DataFrame(all_flags).to_excel(scanner.TRIGGER_FLAGS_XLSX, index=False)
    else:
        pd.DataFrame(columns=["Company", "Source Type", "Category", "What", "Date", "Source Link"]).to_excel(scanner.TRIGGER_FLAGS_XLSX, index=False)

    strength_by_company = {c: scanner.classify_strength(f) for c, f in flags_by_company.items()}
    accepted = {"Strong"} if scanner.REQUEUE_MIN_STRENGTH == "Strong" else {"Strong", "Moderate"}
    qualifying = [c for c, s in strength_by_company.items() if s in accepted]

    if qualifying:
        context_df = master_df[master_df["Company"].isin(qualifying)].copy()
        context_df["Scanner Strength"] = context_df["Company"].map(strength_by_company)
        context_df["Scanner Categories"] = context_df["Company"].map(
            lambda c: ", ".join(sorted(set(f["Category"] for f in flags_by_company.get(c, []))))
        )
        context_df["Scanner Reason"] = context_df["Company"].map(
            lambda c: "\n".join(f"• {f['What']}" for f in flags_by_company.get(c, []))
        )
    else:
        context_df = pd.DataFrame(columns=list(master_df.columns) + ["Scanner Strength", "Scanner Categories", "Scanner Reason"])

    report_export.write_scanner_reference_excel(context_df, scanner.PRIORITY_REQUEUE_XLSX)

    xlsx_style.style_xlsx(scanner.TRIGGER_FLAGS_XLSX)
    return len(qualifying)


def _run_scanner_chunk():
    global _scanner_next_index
    try:
        scanner._load_cik_lookup()

        with _state_lock:
            chunk_size = scanner_state["chunk_size"]
            start = _scanner_next_index
        end = min(start + chunk_size, len(_scanner_companies))
        chunk = _scanner_companies[start:end]

        def _scan_one(company):
            return company, scanner.scan_company(company)

        with ThreadPoolExecutor(max_workers=SCANNER_WORKERS) as pool:
            futures = [pool.submit(_scan_one, c) for c in chunk]
            for future in as_completed(futures):
                company, flags = future.result()
                if flags:
                    _scanner_flags_by_company[company] = flags
                    _scanner_all_flags.extend(flags)
                with _state_lock:
                    scanner_state["scanned"] += 1

        _scanner_next_index = end
        flagged_count = _write_scanner_outputs(_scanner_master_df, _scanner_flags_by_company, _scanner_all_flags)

        is_last_chunk = _scanner_next_index >= len(_scanner_companies)
        with _state_lock:
            scanner_state.update(
                status="done" if is_last_chunk else "paused",
                flagged_count=flagged_count,
                chunks_done=scanner_state["chunks_done"] + 1,
                finished_at=time.time() if is_last_chunk else None,
            )

    except BaseException as e:
        with _state_lock:
            scanner_state.update(status="error", error=f"{e}\n{traceback.format_exc()}")


def _start_scanner_chunk_thread():
    thread = threading.Thread(target=_run_scanner_chunk, daemon=True)
    thread.start()


@app.route("/api/scanner/run", methods=["POST"])
def scanner_run():
    global _scanner_companies, _scanner_master_df, _scanner_next_index, _scanner_flags_by_company, _scanner_all_flags

    with _state_lock:
        if scanner_state["status"] == "running":
            return jsonify({"error": "Scanner already running"}), 409
        scanner_state["status"] = "running"

    body = request.get_json(silent=True) or {}
    chunk_size = int(body.get("chunk_size", DEFAULT_SCANNER_CHUNK_SIZE))

    _scanner_master_df = load_enriched_master()
    _scanner_companies = _scanner_master_df["Company"].tolist()
    _scanner_next_index = 0
    _scanner_flags_by_company = {}
    _scanner_all_flags = []

    with _state_lock:
        scanner_state.update(
            status="running", scanned=0, total=len(_scanner_companies), flagged_count=0, error=None,
            finished_at=None, chunk_size=chunk_size, chunks_done=0,
            chunks_total=-(-len(_scanner_companies) // chunk_size),
        )

    _start_scanner_chunk_thread()
    return jsonify({"started": True})


@app.route("/api/scanner/continue", methods=["POST"])
def scanner_continue():
    with _state_lock:
        if scanner_state["status"] != "paused":
            return jsonify({"error": "Scanner is not paused -- nothing to continue"}), 409
        scanner_state["status"] = "running"
    _start_scanner_chunk_thread()
    return jsonify({"started": True})


@app.route("/api/scanner/status")
def scanner_status():
    with _state_lock:
        return jsonify(dict(scanner_state))


@app.route("/api/scanner/download/<name>")
def scanner_download(name):
    # FIX: trigger_flags.xlsx -- serve directly from scanner output
    if name == "trigger_flags.xlsx":
        path = scanner.TRIGGER_FLAGS_XLSX
        if not Path(path).exists():
            return jsonify({"error": "not found"}), 404
        return send_file(path, as_attachment=True)

    # FIX: priority_requeue.xlsx -- was incorrectly reading from ara.OUTPUT_XLSX
    # (Account Research results). Corrected to serve the scanner's own output directly.
    if name == "priority_requeue.xlsx":
        if not Path(scanner.PRIORITY_REQUEUE_XLSX).exists():
            return jsonify({"error": "not found"}), 404
        return send_file(scanner.PRIORITY_REQUEUE_XLSX, as_attachment=True)

    return jsonify({"error": "not found"}), 404


@app.route("/api/research/preview")
def research_preview():
    try:
        plan = ara.plan_research(ara.INPUT_XLSX)
        return jsonify(plan)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/research/run", methods=["POST"])
def research_run():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY is not set -- check .env in this folder."}), 400

    with _state_lock:
        if research_state["status"] == "running":
            return jsonify({"error": "Research already running"}), 409
        research_state.update(status="running", result=None, error=None)

    thread = threading.Thread(target=_run_research_job, daemon=True)
    thread.start()
    return jsonify({"started": True})


def _run_research_job():
    try:
        result = ara.run(ara.INPUT_XLSX)
        with _state_lock:
            research_state.update(status="done", result=result)
    except BaseException as e:
        with _state_lock:
            research_state.update(status="error", error=f"{e}\n{traceback.format_exc()}")


@app.route("/api/research/status")
def research_status():
    with _state_lock:
        snapshot = dict(research_state)
    batch_snapshot = ara.current_batch_status()
    if batch_snapshot:
        snapshot["batch"] = batch_snapshot
    return jsonify(snapshot)


@app.route("/api/research/download")
def research_download_raw():
    if not Path(ara.OUTPUT_XLSX).exists():
        return jsonify({"error": "not found"}), 404
    # FIX: copy to /tmp before styling -- Vercel root fs is read-only,
    # xlsx_style modifies the file in-place so we can't style the original directly.
    tmp_path = _tmp("Account_Research_Results_styled.xlsx")
    shutil.copy2(ara.OUTPUT_XLSX, tmp_path)
    xlsx_style.style_xlsx(tmp_path)
    return send_file(tmp_path, as_attachment=True, download_name="Account_Research_Results.xlsx")


@app.route("/api/research/download/contacts")
def research_download_contacts():
    if not Path(contacts_tracker.CONTACTS_TRACKER_XLSX).exists():
        return jsonify({"error": "not found"}), 404
    return send_file(contacts_tracker.CONTACTS_TRACKER_XLSX, as_attachment=True)


@app.route("/api/research/download/reference")
def research_download_reference():
    if not Path(ara.OUTPUT_XLSX).exists():
        return jsonify({"error": "not found"}), 404
    results_df = pd.read_excel(ara.OUTPUT_XLSX)
    master_df = load_enriched_master()
    # FIX: write to /tmp -- Vercel root fs is read-only
    out_path = _tmp("Prospect_List_reference_format.xlsx")
    report_export.write_reference_excel(results_df, out_path, master_df)
    return send_file(out_path, as_attachment=True, download_name="Prospect_List_reference_format.xlsx")


# --------------------------------------------------------------------------
# Demo: last Faheem-reviewed batch
# --------------------------------------------------------------------------

DEMO_REVIEWED_XLSX = "Faheem_Review_10_Account_Batch.xlsx"


@app.route("/api/demo/reference")
def demo_download_reference():
    if not Path(DEMO_REVIEWED_XLSX).exists():
        return jsonify({"error": f"{DEMO_REVIEWED_XLSX} not found"}), 404
    results_df = pd.read_excel(DEMO_REVIEWED_XLSX)
    # FIX: write to /tmp -- Vercel root fs is read-only
    out_path = _tmp("Prospect_List_reference_format_demo.xlsx")
    report_export.write_reference_excel(results_df, out_path)
    return send_file(out_path, as_attachment=True, download_name="Prospect_List_reference_format_demo.xlsx")


# --------------------------------------------------------------------------
# Home page
# --------------------------------------------------------------------------

HOME_HTML = """
<!doctype html>
<title>Voiant Prospect Intelligence</title>
<style>
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif; max-width: 760px; margin: 40px auto; color: #1a1a1a; }
  h1 { font-size: 22px; }
  .panel { border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin-bottom: 24px; }
  .panel h2 { margin-top: 0; font-size: 17px; }
  button { font-size: 15px; padding: 10px 18px; border-radius: 6px; border: none; cursor: pointer; }
  .primary { background: #2563eb; color: white; }
  .primary:disabled { background: #93b4ee; cursor: not-allowed; }
  .status { margin-top: 12px; font-size: 14px; color: #444; white-space: pre-wrap; }
  .warn { color: #b45309; }
  .err { color: #b91c1c; }
  a.dl { display: inline-block; margin-top: 8px; margin-right: 14px; font-size: 14px; }
  .cost-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 12px; margin-top: 12px; font-size: 14px; }
</style>

<h1>Voiant Prospect Intelligence</h1>

<div class="panel">
  <h2>1. Run Scanner <span style="font-weight:normal;color:#666;">(free, all ~595 companies)</span></h2>
  <p style="color:#555;font-size:14px;">Checks SEC EDGAR filings + Google News for every company. No API cost either way -- chunking below is just so you can look at partial results as it goes, not to save money.</p>
  <label style="font-size:14px;color:#444;">Chunk size: <input type="number" id="chunkSize" value="100" min="10" max="595" style="width:70px;"></label>
  <br><br>
  <button class="primary" id="scannerBtn" onclick="runScanner()">Run Scanner</button>
  <button class="primary" id="continueBtn" onclick="continueScanner()" style="display:none;">Continue Next Chunk</button>
  <div class="status" id="scannerStatus"></div>
</div>

<div class="panel">
  <h2>2. Run Account Research <span style="font-weight:normal;color:#666;">(paid -- Claude API)</span></h2>
  <p style="color:#555;font-size:14px;">Deep research on companies that need it. Costs real money. Shows an estimate before running anything.</p>
  <button class="primary" id="previewBtn" onclick="previewResearch()">Check Cost</button>
  <div id="costBox"></div>
  <div class="status" id="researchStatus"></div>
</div>

<div class="panel">
  <h2>Preview: last reviewed batch <span style="font-weight:normal;color:#666;">(10 accounts, Faheem sign-off)</span></h2>
  <p style="color:#555;font-size:14px;">Real, reviewed research already completed -- no API call, no cost. Use this to see the reference-format output while credits are unavailable.</p>
  <a class="dl" href="/api/demo/reference">Download in reference format</a>
</div>

<script>
async function runScanner() {
  document.getElementById('scannerBtn').disabled = true;
  document.getElementById('continueBtn').style.display = 'none';
  const chunkSize = parseInt(document.getElementById('chunkSize').value, 10) || 100;
  await fetch('/api/scanner/run', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({chunk_size: chunkSize}),
  });
  pollScanner();
}

async function continueScanner() {
  document.getElementById('continueBtn').disabled = true;
  await fetch('/api/scanner/continue', {method: 'POST'});
  pollScanner();
}

async function pollScanner() {
  const r = await fetch('/api/scanner/status');
  const s = await r.json();
  const el = document.getElementById('scannerStatus');
  const downloadLinks =
    `<a class="dl" href="/api/scanner/download/trigger_flags.xlsx">Download trigger_flags.xlsx</a>` +
    `<a class="dl" href="/api/scanner/download/priority_requeue.xlsx">Download priority_requeue.xlsx</a>`;
  if (s.status === 'running') {
    document.getElementById('continueBtn').style.display = 'none';
    el.textContent = `Scanning chunk ${s.chunks_done + 1}/${s.chunks_total}... ${s.scanned}/${s.total} companies checked.`;
    setTimeout(pollScanner, 1500);
  } else if (s.status === 'paused') {
    el.innerHTML = `Chunk ${s.chunks_done}/${s.chunks_total} done -- ${s.scanned}/${s.total} companies checked so far, ` +
      `${s.flagged_count} flagged. Results below are current up to this point.<br>${downloadLinks}`;
    document.getElementById('scannerBtn').disabled = false;
    document.getElementById('continueBtn').disabled = false;
    document.getElementById('continueBtn').style.display = 'inline-block';
  } else if (s.status === 'done') {
    document.getElementById('continueBtn').style.display = 'none';
    el.innerHTML = `All ${s.total} companies scanned. ${s.flagged_count} flagged with new activity.<br>${downloadLinks}`;
    document.getElementById('scannerBtn').disabled = false;
  } else if (s.status === 'error') {
    document.getElementById('continueBtn').style.display = 'none';
    el.innerHTML = '<span class="err">Error: ' + s.error + '</span>';
    document.getElementById('scannerBtn').disabled = false;
  }
}

async function previewResearch() {
  const r = await fetch('/api/research/preview');
  const p = await r.json();
  if (p.error) {
    document.getElementById('costBox').innerHTML = '<div class="cost-box err">' + p.error + '</div>';
    return;
  }
  document.getElementById('costBox').innerHTML =
    `<div class="cost-box">` +
    `<b>${p.needs_research_count}</b> companies need research (${p.reused_from_cache_count} already cached, no cost).<br>` +
    `Estimated cost: <b>~$${p.estimated_cost_usd}</b> (rough estimate, $${p.cost_per_account_usd}/company).<br><br>` +
    `<button class="primary" onclick="confirmResearch()">Confirm and Run</button>` +
    `</div>`;
}

async function confirmResearch() {
  document.getElementById('previewBtn').disabled = true;
  await fetch('/api/research/run', {method: 'POST'});
  document.getElementById('costBox').innerHTML = '';
  pollResearch();
}

async function pollResearch() {
  const r = await fetch('/api/research/status');
  const s = await r.json();
  const el = document.getElementById('researchStatus');
  if (s.status === 'running') {
    let extra = '';
    if (s.batch && s.batch.processing_status) {
      extra = ` (batch status: ${s.batch.processing_status})`;
    }
    el.textContent = 'Researching...' + extra + ' This can take a while -- safe to close this tab and come back.';
    setTimeout(pollResearch, 3000);
  } else if (s.status === 'done') {
    const res = s.result;
    el.innerHTML = `Done. ${res.freshly_researched_count} freshly researched, ${res.from_cache_count} reused from cache, ` +
      `${res.outreach_ready_count} Outreach Ready, ${res.needs_review_count} flagged for review.<br>` +
      `<a class="dl" href="/api/research/download">Download full results</a>` +
      `<a class="dl" href="/api/research/download/reference">Download in reference format</a>` +
      `<a class="dl" href="/api/research/download/contacts">Download contacts outreach tracker</a>`;
    document.getElementById('previewBtn').disabled = false;
  } else if (s.status === 'error') {
    el.innerHTML = '<span class="err">Error: ' + s.error + '</span>';
    document.getElementById('previewBtn').disabled = false;
  }
}

pollScanner();
pollResearch();
</script>
"""


@app.route("/")
def home():
    return render_template_string(HOME_HTML)


if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5057))
    app.run(host="0.0.0.0", port=port, threaded=True, debug=False)