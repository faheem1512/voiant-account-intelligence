"""
Dry-run test for scanner.py -- mocks EDGAR and Google News responses to
verify parsing, categorization, and strength classification WITHOUT
hitting the real network.
"""

import sys
from unittest.mock import patch, MagicMock
from datetime import datetime, timedelta, timezone

import scanner


def test_edgar_8k_item_5_02_is_exec_change():
    print("Testing EDGAR 8-K parsing: Item 5.02 -> exec_change_8k (via CIK submissions API)...")
    scanner._CIK_LOOKUP_CACHE = {"mckesson corp": "0000927653"}

    fake_response = MagicMock()
    fake_response.raise_for_status = MagicMock()
    today = datetime.now(timezone.utc).date()
    fake_response.json.return_value = {
        "filings": {"recent": {
            "form": ["8-K", "10-Q"],
            "filingDate": [today.isoformat(), (today - timedelta(days=100)).isoformat()],
            "accessionNumber": ["0000927653-26-000123", "0000927653-26-000099"],
            "primaryDocument": ["form8k.htm", "form10q.htm"],
            "items": ["5.02", ""],
        }}
    }

    with patch("scanner.requests.get", return_value=fake_response):
        hits = scanner.check_edgar_filings("McKesson", 7)

    assert len(hits) == 1, f"FAIL: expected 1 hit (8-K only, 10-Q not in target forms), got {len(hits)}"
    assert hits[0]["form"] == "8-K"
    assert hits[0]["category"] == "exec_change_8k", f"FAIL: Item 5.02 should categorize as exec_change_8k, got {hits[0]['category']}"
    print("  PASS: 8-K with Item 5.02 correctly categorized as an actual exec change, 10-Q correctly excluded")


def test_edgar_8k_other_item_is_not_exec_change():
    print("Testing EDGAR 8-K parsing: a non-5.02 item -> material_event_8k_other, NOT exec change...")
    scanner._CIK_LOOKUP_CACHE = {"mckesson corp": "0000927653"}

    fake_response = MagicMock()
    fake_response.raise_for_status = MagicMock()
    today = datetime.now(timezone.utc).date()
    fake_response.json.return_value = {
        "filings": {"recent": {
            "form": ["8-K"],
            "filingDate": [today.isoformat()],
            "accessionNumber": ["0000927653-26-000124"],
            "primaryDocument": ["form8k.htm"],
            "items": ["2.02,9.01"],  # earnings results, not an exec change
        }}
    }

    with patch("scanner.requests.get", return_value=fake_response):
        hits = scanner.check_edgar_filings("McKesson", 7)

    assert len(hits) == 1
    assert hits[0]["category"] == "material_event_8k_other", (
        f"FAIL: a non-5.02 8-K must NOT be labeled as an exec change, got {hits[0]['category']}"
    )
    print("  PASS: non-5.02 8-K correctly labeled material_event_8k_other, not confused with an exec change")


def test_edgar_s1_detected_as_ipo_category():
    print("Testing EDGAR S-1 detection (IPO category)...")
    scanner._CIK_LOOKUP_CACHE = {"newco inc": "0001111111"}

    fake_response = MagicMock()
    fake_response.raise_for_status = MagicMock()
    today = datetime.now(timezone.utc).date()
    fake_response.json.return_value = {
        "filings": {"recent": {
            "form": ["S-1"],
            "filingDate": [today.isoformat()],
            "accessionNumber": ["0001111111-26-000001"],
            "primaryDocument": ["s1.htm"],
        }}
    }

    with patch("scanner.requests.get", return_value=fake_response):
        hits = scanner.check_edgar_filings("NewCo Inc", 7)

    assert len(hits) == 1
    assert hits[0]["form"] == "S-1"
    assert hits[0]["category"] == "ipo_scaling", f"FAIL: S-1 should categorize as ipo_scaling, got {hits[0]['category']}"
    print("  PASS: S-1 filing correctly detected and categorized as ipo_scaling")


def test_cik_resolution_failure_skips_gracefully():
    print("Testing CIK resolution failure handling...")
    scanner._CIK_LOOKUP_CACHE = {}  # empty lookup -- nothing will resolve
    hits = scanner.check_edgar_filings("Totally Unknown Company XYZ", 7)
    assert hits == [], "FAIL: should return empty list when CIK can't be resolved, not crash"
    print("  PASS: unresolvable CIK handled gracefully (no crash)")


def test_news_headline_must_contain_company_name():
    print("Testing news headline-must-contain-name filter (rejects off-topic matches)...")
    now = datetime.now(timezone.utc)
    recent_time = (now - timedelta(days=2)).timetuple()

    fake_feed = MagicMock()
    fake_feed.entries = [
        MagicMock(title="Bank of America comments on International Paper layoffs",
                   link="http://example.com/1", published_parsed=recent_time),  # name in headline, on-topic-ish
        MagicMock(title="International Paper announces layoffs; Bank of America downgrades stock",
                   link="http://example.com/2", published_parsed=recent_time),  # Bank of America IS in headline here too, so this would pass -- need a case where it's NOT
    ]
    # Better test: a case where the company name does NOT appear in the headline at all
    fake_feed.entries.append(
        MagicMock(title="International Paper announces major layoffs",
                   link="http://example.com/3", published_parsed=recent_time)
    )

    with patch("scanner.feedparser.parse", return_value=fake_feed):
        hits = scanner.check_google_news("Bank of America", 7, scanner.ALL_NEWS_KEYWORDS)

    for hit in hits:
        assert "bank of america" in hit["title"].lower(), f"FAIL: hit headline doesn't contain company name: {hit['title']}"
    print(f"  PASS: {len(hits)} hit(s) found, all had 'Bank of America' actually in the headline")


def test_categorize_title():
    print("Testing headline-to-category mapping...")
    assert scanner.categorize_title("acme files for ipo")[0] == "ipo_scaling"
    assert scanner.categorize_title("acme announces layoffs")[0] == "workforce_reduction"
    assert scanner.categorize_title("acme cites forecasting challenges")[0] == "forecasting_chaos"
    assert scanner.categorize_title("acme announces territory realignment")[0] == "territory_quota_miss"
    assert scanner.categorize_title("acme sees sales rep turnover spike")[0] == "rep_attrition"
    assert scanner.categorize_title("acme undergoing restructuring")[0] == "restructuring_language"
    print("  PASS: all non-role-change categories map correctly")


def test_categorize_title_role_change_requires_both_role_and_verb():
    print("Testing role-change categories require BOTH a role term AND an action verb...")
    assert scanner.categorize_title("acme names new cfo")[0] == "cfo_cro_change"
    assert scanner.categorize_title("acme appoints new vp of sales operations")[0] == "new_sales_ops_revops_leadership"
    # the actual bug this fixes: "appoints" alone used to categorize ANY
    # hire as a CFO/CRO change, including one totally unrelated
    result = scanner.categorize_title("acme appoints new vp of marketing")
    assert result[0] is None, f"FAIL: an unrelated VP of Marketing hire must NOT match any role-change category, got {result}"
    print("  PASS: correctly requires role+verb together, no more false positives on unrelated hires")


def test_classify_strength_edgar_exec_change_alone_is_strong():
    print("Testing classify_strength: EDGAR 8-K Item 5.02 (exec change) alone -> Strong...")
    flags = [{"Source Type": "SEC EDGAR 8-K", "Category": "exec_change_8k"}]
    assert scanner.classify_strength(flags) == "Strong"
    print("  PASS")


def test_classify_strength_generic_8k_alone_is_not_strong():
    print("Testing classify_strength: a generic 8-K (not an exec change) alone -> None, NOT Strong...")
    flags = [{"Source Type": "SEC EDGAR 8-K", "Category": "material_event_8k_other"}]
    result = scanner.classify_strength(flags)
    assert result == "None", f"FAIL: a generic material 8-K must not auto-qualify as Strong, got {result}"
    print("  PASS: generic 8-Ks (earnings, agreements, etc) no longer inflate every filing to Strong")


def test_classify_strength_new_sales_ops_leadership_news_is_moderate():
    print("Testing classify_strength: new Sales Ops/RevOps leadership hire via news alone -> Moderate...")
    flags = [{"Source Type": "Google News (matched: x)", "Category": "new_sales_ops_revops_leadership"}]
    assert scanner.classify_strength(flags) == "Moderate"
    print("  PASS")


def test_classify_strength_edgar_s1_alone_is_strong():
    print("Testing classify_strength: EDGAR S-1 alone -> Strong...")
    flags = [{"Source Type": "SEC EDGAR S-1", "Category": "ipo_scaling"}]
    assert scanner.classify_strength(flags) == "Strong"
    print("  PASS")


def test_classify_strength_single_ipo_news_is_strong():
    print("Testing classify_strength: single IPO news hit -> Strong (no EDGAR needed)...")
    flags = [{"Source Type": "Google News (matched: ipo)", "Category": "ipo_scaling"}]
    assert scanner.classify_strength(flags) == "Strong"
    print("  PASS")


def test_classify_strength_single_cfo_news_only_is_moderate_not_strong():
    print("Testing classify_strength: single CFO news hit with NO EDGAR corroboration -> Moderate, not Strong...")
    flags = [{"Source Type": "Google News (matched: cfo)", "Category": "cfo_cro_change"}]
    result = scanner.classify_strength(flags)
    assert result == "Moderate", f"FAIL: expected Moderate (uncorroborated CFO mention shouldn't auto-Strong), got {result}"
    print("  PASS: correctly avoids the 'any CFO news = Strong' mistake we fixed in the paid script")


def test_classify_strength_two_distinct_weak_categories_is_strong():
    print("Testing classify_strength: 2 DIFFERENT weak categories -> Strong...")
    flags = [
        {"Source Type": "Google News (matched: forecasting challenges)", "Category": "forecasting_chaos"},
        {"Source Type": "Google News (matched: territory realignment)", "Category": "territory_quota_miss"},
    ]
    assert scanner.classify_strength(flags) == "Strong"
    print("  PASS")


def test_classify_strength_two_hits_same_weak_category_is_only_moderate():
    print("Testing classify_strength: 2 hits of the SAME weak category -> still only Moderate...")
    flags = [
        {"Source Type": "Google News (matched: layoff)", "Category": "workforce_reduction"},
        {"Source Type": "Google News (matched: job cuts)", "Category": "workforce_reduction"},
    ]
    result = scanner.classify_strength(flags)
    assert result == "Moderate", f"FAIL: 2 hits of the SAME category shouldn't count as 2 distinct weak triggers, got {result}"
    print("  PASS: correctly distinguishes 'distinct categories' from 'raw hit count'")


def test_classify_strength_single_weak_category_is_moderate():
    print("Testing classify_strength: exactly 1 weak category -> Moderate...")
    flags = [{"Source Type": "Google News (matched: layoff)", "Category": "workforce_reduction"}]
    assert scanner.classify_strength(flags) == "Moderate"
    print("  PASS")


def test_classify_strength_nothing_is_none():
    print("Testing classify_strength: no flags -> None...")
    assert scanner.classify_strength([]) == "None"
    print("  PASS")


if __name__ == "__main__":
    tests = [
        test_edgar_8k_item_5_02_is_exec_change,
        test_edgar_8k_other_item_is_not_exec_change,
        test_edgar_s1_detected_as_ipo_category,
        test_cik_resolution_failure_skips_gracefully,
        test_news_headline_must_contain_company_name,
        test_categorize_title,
        test_categorize_title_role_change_requires_both_role_and_verb,
        test_classify_strength_edgar_exec_change_alone_is_strong,
        test_classify_strength_generic_8k_alone_is_not_strong,
        test_classify_strength_edgar_s1_alone_is_strong,
        test_classify_strength_single_ipo_news_is_strong,
        test_classify_strength_single_cfo_news_only_is_moderate_not_strong,
        test_classify_strength_new_sales_ops_leadership_news_is_moderate,
        test_classify_strength_two_distinct_weak_categories_is_strong,
        test_classify_strength_two_hits_same_weak_category_is_only_moderate,
        test_classify_strength_single_weak_category_is_moderate,
        test_classify_strength_nothing_is_none,
    ]
    for t in tests:
        t()
    print()
    print("=" * 60)
    print("ALL SCANNER LOGIC TESTS PASSED (no real network calls made)")
    print("=" * 60)