# Voiant Prospect Intelligence — Codebase Handoff

**Owner:** Faheem Moosa (handed off from Zain Rashid, original builder)  
**Deployed on:** Vercel (project: voiant-account-intelligence)  
**GitHub:** github.com/faheem1512/voiant-account-intelligence  

---

## What This System Does

This is a B2B sales intelligence tool that identifies outreach-ready accounts
from the Forbes Global 2000 US company list. It looks for companies using
Anaplan, Pigment, or Varicent for **sales-side planning** (not Finance/FP&A —
this distinction is critical and must be preserved in all future changes).

The system has two parts:

1. **Free Scanner** — checks ~595 companies against SEC EDGAR filings and
   Google News RSS every morning. Costs nothing. Outputs a priority list of
   companies worth researching.

2. **Claude Project Research** — Faheem pastes the priority list into a Claude
   Project with embedded research instructions. Claude researches each company
   and returns a scored, sourced table ready for outreach decisions.

The Vercel dashboard is for viewing results and finding ICP prospects — it
shows scanner results and lets you download files. Faheem runs the scanner
locally and pushes the output files to GitHub, which auto-updates the Vercel
URL. (The Run Scanner / Run Research buttons in the dashboard work locally but
not reliably on Vercel — its filesystem is read-only and background jobs are
cut off. Run those locally.)

---

## The Three Target Platforms

Only these three matter. Everything else (Xactly, Excel, Adaptive, Workday)
is out of scope:

- **Anaplan**
- **Pigment**
- **Varicent**

---

## Ideal Client Profile (ICP)

- Global 2000 US enterprise, operating in 2+ countries
- Uses one of the three platforms **for Sales Ops / RevOps** (not Finance only)
- Industry (in priority order): Technology & SaaS → Healthcare & MedTech →
  Financial Services & Insurance

---

## File Structure

```
voiant-account-intelligence/
│
├── app.py                          # Flask web app (the Vercel dashboard)
├── scanner.py                      # Free trigger scanner (SEC EDGAR + Google News)
├── Account_research_automation.py  # Paid Claude API research layer (legacy)
├── report_export.py                # Formats output into reference Excel format
├── contacts_tracker.py             # Tracks named contacts found in research
├── industry_mapping.py             # Maps Forbes industry labels to ICP buckets
├── xlsx_style.py                   # Applies formatting to Excel output files
│
├── Forbes_2000_master.xlsx         # ~595 US Global 2000 companies (the input list)
├── Faheem_Review_10_Account_Batch.xlsx   # 10 reviewed accounts (demo reference)
├── Prospect_List_reference_format.xlsx  # Output template / column reference
├── trigger_flags.xlsx              # Scanner output: every trigger hit found
├── priority_requeue.xlsx           # Scanner output: companies worth researching
│
├── Procfile                        # Tells Railway/Vercel: web: python app.py
├── requirements.txt                # Python dependencies
├── vercel.json                     # Vercel deployment config
```

---

## How to Run Locally

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Create a .env file in the repo folder with:
APP_PASSWORD=your_password_here
SECRET_KEY=your_secret_key_here

# 3. Run the app
python app.py

# 4. Open in browser
http://localhost:5057
```

---

## Daily Morning Workflow

1. **Faheem runs the scanner locally:**
   ```bash
   python scanner.py
   ```
   This generates `trigger_flags.xlsx` and `priority_requeue.xlsx`

2. **Faheem pushes results to GitHub** (auto-updates Faheem's Vercel link):
   ```bash
   git add trigger_flags.xlsx priority_requeue.xlsx
   git commit -m "update scanner results"
   git push
   ```

3. **Open the Vercel URL** (log in with `APP_PASSWORD`) — download the priority list

4. **Paste the priority list into the Claude Project** — Claude
   researches each company using the embedded instructions and returns
   a scored table

5. **Review Band A and B accounts** and do manual outreach

---

## How to Swap the Company List

This is the most likely change. To use a different input list instead of
the Forbes Global 2000:

**Step 1** — Replace `Forbes_2000_master.xlsx` with your new Excel file.
Your file must have at minimum these columns:
- `Company` (company name)
- `Industry` (free text — mapped to ICP buckets via `industry_mapping.py`)
- `Headquarters` (optional but useful)

**Step 2** — Update the filename reference in `app.py` (`MASTER_LIST_XLSX`):
```python
MASTER_LIST_XLSX = "Forbes_2000_master.xlsx"  # ← change to your filename
```

**Step 3** — Update the filename reference in `scanner.py` (`INPUT_XLSX`). Note: running
`python scanner.py` uses this file, while the dashboard uses `MASTER_LIST_XLSX`;
point both at the same list if you want consistent results:
```python
INPUT_XLSX = "test_40_accounts_READY.xlsx"  # ← change to your filename
```

**Step 4** — If your industry labels differ from Forbes categories, update
the mapping in `industry_mapping.py`. The current mapping:
```python
"IT Software & Services" → "Tech & SaaS"
"Drugs & Biotechnology" → "Healthcare & MedTech"
"Banking"               → "FinServ & Insurance"
# etc.
```
Add or change entries to match your new list's industry labels.

**Step 5** — Commit and push. Vercel redeploys automatically.

---

## How to Change the ICP Industries

The three target industries are defined in two places:

**1. `industry_mapping.py`** — controls which Forbes industries map to ICP
buckets vs. "Other". Edit `FORBES_INDUSTRY_TO_ICP_BUCKET` to add/remove
industry mappings.

**2. `Account_research_automation.py`** (line ~100):
```python
ICP_INDUSTRIES = {"tech & saas", "healthcare & medtech", "finserv & insurance"}
```
Update this set to match any new ICP industry names.

---

## How to Change the Scoring Formula

The priority scoring formula lives in `Account_research_automation.py` in the
`compute_priority_score()` function. Current weights (per the brief):

| Factor             | Weight |
|--------------------|--------|
| Trigger strength   | 30%    |
| Platform confidence| 25%    |
| Company size       | 20%    |
| Industry fit       | 15%    |
| Trigger recency    | 10%    |

Priority bands: A (8–10) · B (6–7.9) · C (4–5.9) · D (<4)

---

## The Critical Rule — Sales Ops vs Finance

**This must never be changed or removed:**

A company only qualifies as an ICP match if the **Sales Operations or Revenue
Operations team** is the one using the platform. A Finance/FP&A-only deployment
is explicitly NOT a match, even if the platform is confirmed.

This distinction is enforced in:
- The Claude Project research instructions (separate document)
- `Account_research_automation.py` system prompt
- The `Platform Confidence` rating logic

Do not remove or soften this rule in any future version.

---

## Deployment (Vercel)

The app is deployed at the Vercel URL above. Every push to the `main` branch
on GitHub triggers an automatic redeploy (takes ~60 seconds).

Environment variables are set in the Vercel dashboard (not in code):
- `APP_PASSWORD` — login password for the web interface (required on Vercel;
  the app refuses to serve pages if it is missing)
- `SECRET_KEY` — Flask session signing key (set a long random string, or logins
  will reset between requests)
- `ANTHROPIC_API_KEY` — only needed if you use the paid API research button;
  leave unset to keep it disabled

To redeploy manually: Vercel dashboard → Deployments → Redeploy.

---

## Known Limitations

These are gaps that cannot be resolved with free data sources:

- **Repeated job-posting detection** — no free historical job posting data
- **G2/Gartner/TrustRadius monitoring** — no free API
- **LinkedIn people-search** — deliberately avoided (ToS)
- **Employee headcount** — defaults to lowest size bucket (data not reliably free)
- **Glassdoor signals** — no free API; Google News is a partial proxy only

---

## Key Files NOT to Change Without Understanding the Impact

| File | Why it's sensitive |
|------|--------------------|
| `industry_mapping.py` | Changes here affect which companies are treated as ICP vs Other across the entire pipeline |
| `Account_research_automation.py` system prompt | The Sales-vs-Finance distinction and scoring rubric live here |
| `report_export.py` | Column names match the reference format Faheem uses for outreach decisions |
| `vercel.json` | Changing this can break the deployment |
| `requirements.txt` | Only contains what the app needs — don't run pip freeze to regenerate it |
