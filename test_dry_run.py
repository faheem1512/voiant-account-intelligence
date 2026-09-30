"""
Dry-run test for account_research_automation.py (BATCH VERSION)

Mocks the Anthropic Batches API so you can verify the parsing, caching,
scoring, and Excel output logic work correctly WITHOUT spending any API
credits or needing a real key. Run this before your first real batch run.

This replaces the old test_dry_run.py / test_run.py, which mocked the
single-call API and will throw AttributeError against this version of
the script (research_account() no longer exists -- replaced by the
batch submit/poll/fetch functions below).

Usage:
    python test_dry_run.py
"""

import sys
import json
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta, timezone

import pandas as pd

import account_research_automation as ara

MOCK_RESPONSES = {
    "Alphabet": {
        "company": "Alphabet", "platform": "Anaplan", "platform_confidence": "Confirmed",
        "department": "Sales-Ops (GCBP)", "named_contact": "", "contact_title": "",
        "contact_linkedin_url": "", "trigger_event": "None identified", "trigger_strength": "None",
        "trigger_source_and_date": "", "weak_trigger_count": 0, "trigger_recency_bucket": "N/A",
        "outreach_ready": False, "low_confidence_flag": "", "source_links": "cloud.google.com/blog/x",
        "research_summary": "Anaplan confirmed sales-side.",
    },
    "JPMorganChase": {
        "company": "JPMorganChase", "platform": "Anaplan", "platform_confidence": "Inferred",
        "department": "Finance-only", "named_contact": "", "contact_title": "", "contact_linkedin_url": "",
        "trigger_event": "None identified", "trigger_strength": "None", "trigger_source_and_date": "",
        "weak_trigger_count": 0, "trigger_recency_bucket": "N/A", "outreach_ready": False,
        "low_confidence_flag": "", "source_links": "anaplan.com/blog/x",
        "research_summary": "Anaplan confirmed but Finance-led.",
    },
}


def make_fake_batch_entry(custom_id, company):
    """Mimics one entry from client.messages.batches.results(batch_id)."""
    payload = MOCK_RESPONSES.get(company)
    entry = MagicMock()
    entry.custom_id = custom_id

    if payload is None:
        # Simulate a batch request that errored server-side
        entry.result.type = "errored"
        entry.result.error = "simulated_error"
    else:
        entry.result.type = "succeeded"
        text_block = MagicMock()
        text_block.type = "text"
        text_block.text = json.dumps(payload)
        entry.result.message.content = [text_block]

    return entry


def run_dry_run():
    print("=" * 70)
    print("DRY RUN (batch version) - no real API calls, no cost")
    print("=" * 70)

    test_df = pd.DataFrame({
        "Company": ["Alphabet", "JPMorganChase", "Unmocked Corp"],
        "ICP_Bucket": ["Tech & SaaS", "FinServ & Insurance", "Tech & SaaS"],
        "Headquarters": ["Mountain View, US", "New York, US", "Nowhere, US"],
    })
    test_input_path = "test_input_accounts.xlsx"
    test_df.to_excel(test_input_path, index=False)
    ara.INPUT_XLSX = test_input_path
    ara.OUTPUT_XLSX = "test_output_results.xlsx"
    ara.TRIGGER_FLAGS_XLSX = "test_flags_that_dont_exist.xlsx"  # simulate no scanner run yet

    import os
    for f in [ara.OUTPUT_XLSX, ara.BATCH_STATE_FILE]:
        if os.path.exists(f):
            os.remove(f)

    id_to_company = {"acct_0": ("Alphabet", "Tech & SaaS", "Mountain View, US"),
                      "acct_1": ("JPMorganChase", "FinServ & Insurance", "New York, US"),
                      "acct_2": ("Unmocked Corp", "Tech & SaaS", "Nowhere, US")}

    fake_batch = MagicMock()
    fake_batch.id = "batch_test_123"
    fake_batch.processing_status = "ended"

    fake_results = [make_fake_batch_entry(cid, ctx[0]) for cid, ctx in id_to_company.items()]

    with patch.object(ara, "build_client") as mock_build_client, \
         patch.object(ara, "_batches_client") as mock_batches_client:

        mock_client = MagicMock()
        mock_build_client.return_value = mock_client

        mock_batches = MagicMock()
        mock_batches.create.return_value = fake_batch
        mock_batches.retrieve.return_value = fake_batch
        mock_batches.results.return_value = fake_results
        mock_batches_client.return_value = mock_batches

        ara.main()

    output_df = pd.read_excel(ara.OUTPUT_XLSX)
    print("\nRESULTS TABLE")
    print(output_df[["Company", "Platform (Anaplan/Pigment/Varicent)",
                      "Platform Confidence (Confirmed/Inferred/Unconfirmed)",
                      "NEEDS_MANUAL_REVIEW", "From Cache (No New API Call)"]].to_string(index=False))

    checks_passed = True
    alphabet_row = output_df[output_df["Company"] == "Alphabet"].iloc[0]
    if alphabet_row["Platform Confidence (Confirmed/Inferred/Unconfirmed)"] == "Confirmed":
        print("PASS: Alphabet correctly parsed as Confirmed via batch path")
    else:
        print("FAIL: Alphabet should be Confirmed"); checks_passed = False

    jpm_row = output_df[output_df["Company"] == "JPMorganChase"].iloc[0]
    if jpm_row["Platform Confidence (Confirmed/Inferred/Unconfirmed)"] == "Inferred":
        print("PASS: JPMorganChase correctly parsed as Inferred via batch path")
    else:
        print("FAIL: JPMorganChase should be Inferred"); checks_passed = False

    unmocked_row = output_df[output_df["Company"] == "Unmocked Corp"].iloc[0]
    if unmocked_row["NEEDS_MANUAL_REVIEW"] == True:
        print("PASS: Errored batch entry correctly flagged for manual review")
    else:
        print("FAIL: Errored entry should be flagged"); checks_passed = False

    print()
    print("SECOND RUN: verifying cache reuse (Alphabet/JPMorganChase should NOT")
    print("be re-researched; Unmocked Corp SHOULD be retried since it errored)")

    fake_batch_2 = MagicMock()
    fake_batch_2.id = "batch_test_456"
    fake_batch_2.processing_status = "ended"
    retry_results = [make_fake_batch_entry("acct_0", "Unmocked Corp")]  # still errors again

    with patch.object(ara, "build_client") as mock_build_client, \
         patch.object(ara, "_batches_client") as mock_batches_client:
        mock_client = MagicMock()
        mock_build_client.return_value = mock_client
        mock_batches = MagicMock()
        mock_batches.create.return_value = fake_batch_2
        mock_batches.retrieve.return_value = fake_batch_2
        mock_batches.results.return_value = retry_results
        mock_batches_client.return_value = mock_batches
        ara.main()

        submitted_companies_count = mock_batches.create.call_args[1]["requests"] if mock_batches.create.called else []
        if mock_batches.create.called and len(submitted_companies_count) == 1:
            print("PASS: only the 1 errored account was resubmitted -- Alphabet/JPMorganChase correctly reused from cache")
        elif not mock_batches.create.called:
            print("FAIL: expected exactly 1 resubmission (the errored account), but none was submitted")
            checks_passed = False
        else:
            print(f"FAIL: expected exactly 1 resubmission, got {len(submitted_companies_count)}")
            checks_passed = False

    print("\n" + "=" * 70)
    print("ALL CHECKS PASSED" if checks_passed else "SOME CHECKS FAILED")
    print("=" * 70)
    return checks_passed


if __name__ == "__main__":
    sys.exit(0 if run_dry_run() else 1)