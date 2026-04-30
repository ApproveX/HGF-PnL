from pathlib import Path
import json

import polars as pl

from hgf_pnl.extractors.th_revenue import THRevenueConfig, extract_th_revenue


SAMPLE = Path("sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx")


def test_extracts_th_revenue_report_sample() -> None:
    result = extract_th_revenue(SAMPLE)

    assert set(result.sheets) == {"summary", "details", "usa_stock"}
    assert result.sheets["summary"].sheet_name == "Summary "
    assert result.sheets["summary"].header_row == 2
    assert result.sheets["details"].header_row == 1
    assert result.sheets["usa_stock"].header_row == 1

    summary = result.account_summary
    assert summary.height == 11

    total = summary.filter(pl.col("is_total_row")).to_dicts()[0]
    assert total["account"] == "Total"
    assert round(total["revenue"], 2) == 1_242_393.40
    assert round(total["total_cost"], 3) == 750_574.496
    assert round(total["gross_margin_pct"], 6) == 0.395864
    assert json.loads(total["formula_status"])["revenue"] == "ok"
    assert json.loads(total["formula_status"])["total_cost"] == "ok"

    po_details = result.po_details
    assert po_details.height == 38
    non_total_po = po_details.filter(~pl.col("is_total_row"))
    assert round(non_total_po.select(pl.col("revenue").sum()).item(), 2) == 1_242_393.40

    usa_stock = result.usa_stock
    assert usa_stock.height == 13
    assert round(usa_stock.filter(~pl.col("is_total_row")).select(pl.col("revenue").sum()).item(), 2) == 31_473.30


def test_can_exclude_total_rows() -> None:
    result = extract_th_revenue(SAMPLE, THRevenueConfig(include_total_rows=False))

    assert not result.account_summary["is_total_row"].any()
    assert not result.po_details["is_total_row"].any()
    assert result.account_summary.height == 10
    assert result.po_details.height == 37
    assert result.usa_stock.height == 12
