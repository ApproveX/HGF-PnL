from pathlib import Path

import polars as pl

from hgf_pnl.extractors.chargeback_pdf import (
    extract_chargeback_pdf,
    profile_chargeback_pdf,
    write_profile_artifacts,
)


SAMPLE = Path("sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf")


def test_extracts_chargeback_monthly_summary() -> None:
    result = extract_chargeback_pdf(SAMPLE)

    assert result.subject == "OG | Chargeback Report - 03. March 2026"
    assert not result.warnings

    monthly = result.monthly_summary
    current = monthly.filter((pl.col("year") == 2026) & (pl.col("month_name") == "March"))
    assert current.height == 6

    values = {row["category"]: row for row in current.to_dicts()}
    assert values["allowance"]["amount"] == -65092
    assert values["allowance"]["percent_of_total"] == 0.625
    assert values["penalty"]["amount"] == -7821
    assert values["provision"]["amount"] == -15661
    assert values["return"]["amount"] == -15596
    assert values["software_fees"]["amount"] == -34
    assert values["grand_total"]["amount"] == -104205
    assert values["grand_total"]["source_page"] == 2
    assert values["grand_total"]["source_line"] == 3


def test_extracts_customer_detail_block() -> None:
    result = extract_chargeback_pdf(SAMPLE)
    detail = result.customer_detail

    assert detail.height == 18
    non_total = detail.filter(~pl.col("is_total_row"))
    assert non_total.height == 14
    assert round(non_total.select(pl.col("amount").sum()).item(), 2) == -88_541

    burlington = detail.filter(pl.col("customer") == "Burlington PO").to_dicts()[0]
    assert burlington["department"] == "B&M"
    assert burlington["amount"] == -1884

    grand_total = detail.filter(pl.col("customer") == "Grand Total").to_dicts()[0]
    assert grand_total["amount"] == -88543
    assert grand_total["is_total_row"] is True


def test_extracts_reconciliation_block() -> None:
    result = extract_chargeback_pdf(SAMPLE)
    reconciliation = result.reconciliation

    assert reconciliation.height == 7
    walmart = reconciliation.filter(pl.col("customer") == "Walmart Marketplace").to_dicts()[0]
    assert walmart["cb_report_amount"] == -711.37
    assert walmart["qb_amount"] == -647.62
    assert walmart["difference"] == -63.75
    assert "reversals are ignored" in walmart["note"]

    grand_total = reconciliation.filter(pl.col("customer") == "Grand Total").to_dicts()[0]
    assert grand_total["cb_report_amount"] == -104204.61
    assert grand_total["qb_amount"] == -99389.89
    assert grand_total["difference"] == -4814.72


def test_profiles_pdf_and_suggested_config_extracts(tmp_path: Path) -> None:
    profile = profile_chargeback_pdf(SAMPLE)

    assert profile.page_count == 2
    assert len(profile.lines) > 100
    assert profile.monthly_line_candidates
    assert profile.anchor_candidates
    assert profile.table_summaries
    assert profile.suggested_config.target_month_name == "March"
    assert profile.suggested_config.target_year == 2026

    write_profile_artifacts(profile, tmp_path)
    assert (tmp_path / "chargeback_pdf_profile.md").exists()
    assert (tmp_path / "chargeback_pdf_raw_text.txt").exists()
    assert (tmp_path / "chargeback_pdf_suggested_config.json").exists()

    extraction = extract_chargeback_pdf(SAMPLE, profile.suggested_config)
    current = extraction.monthly_summary.filter(
        (pl.col("year") == 2026)
        & (pl.col("month_name") == "March")
        & (pl.col("category") == "grand_total")
    )
    assert current.select(pl.col("amount")).item() == -104205
