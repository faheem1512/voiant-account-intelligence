"""
Forbes Global 2000 raw industry text -> Voiant ICP bucket.

WHY THIS EXISTS
----------------
The master account list (Forbes 2000 list USA companies.xlsx) only has a
free-text 'Industry' column (27 distinct Forbes categories -- "Banking",
"IT Software & Services", etc). The existing, unchanged scoring formula
in Account_research_automation.py (_industry_points) expects one of the
three literal ICP bucket strings from the brief:
  "Tech & SaaS" | "Healthcare & MedTech" | "FinServ & Insurance"
anything else scores as non-ICP (3pts instead of 9pts).

This is a data-prep crosswalk ONLY -- it does not change the scoring
formula, its weights, or its point values. Those are untouched, per the
build spec's "no changes to the scoring formula" instruction. This just
supplies the input field the existing formula already expects.

JUDGMENT CALL FLAG: which Forbes categories count as "in" vs "Other" is
a real judgment call (e.g. is "Media" Tech & SaaS? is "Diversified
Financials" FinServ?). The mapping below is a reasonable first pass, not
a decision from Faheem -- flag this file for his review rather than
treating it as settled.
"""

FORBES_INDUSTRY_TO_ICP_BUCKET = {
    # -- Tech & SaaS --
    "IT Software & Services": "Tech & SaaS",
    "Technology Hardware & Equipment": "Tech & SaaS",
    "Semiconductors": "Tech & SaaS",
    "Telecommunications Services": "Tech & SaaS",

    # -- Healthcare & MedTech --
    "Drugs & Biotechnology": "Healthcare & MedTech",
    "Health Care Equipment & Services": "Healthcare & MedTech",

    # -- FinServ & Insurance --
    "Banking": "FinServ & Insurance",
    "Insurance": "FinServ & Insurance",
    "Diversified Financials": "FinServ & Insurance",

    # Everything else (Aerospace & Defense, Business Services & Supplies,
    # Capital Goods, Chemicals, Conglomerates, Construction, Consumer
    # Durables, Food Markets, Food/Drink & Tobacco, Hotels/Restaurants &
    # Leisure, Household & Personal Products, Materials, Media, Oil & Gas
    # Operations, Retailing, Trading Companies, Transportation, Utilities)
    # falls through to "Other" via .get()'s default below -- matches the
    # brief's "deprioritize everything else" instruction.
}


def to_icp_bucket(forbes_industry) -> str:
    if not isinstance(forbes_industry, str):
        return "Other"  # covers NaN (pandas gives float('nan') for blank cells) and None
    return FORBES_INDUSTRY_TO_ICP_BUCKET.get(forbes_industry.strip(), "Other")
