from pathlib import Path

import polars as pl

from hgf_pnl.extractors.monthly_revenue import MonthlyRevenueConfig, extract_monthly_revenue


SAMPLE = Path("sample_files/Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx")


def test_extracts_monthly_revenue_summary_and_sales() -> None:
    result = extract_monthly_revenue(SAMPLE)

    assert result.sheets == {
        "summary": "March-Revenue ",
        "shopify": "Shopify",
        "refunds": "Refunds",
        "coupons": "Coupons",
    }
    assert not result.warnings

    summary = result.summary
    revenue_total = summary.filter(
        (pl.col("section") == "REVENUE") & (pl.col("label") == "Grand Total")
    ).to_dicts()[0]
    refund_total = summary.filter(
        (pl.col("section") == "REFUNDS") & (pl.col("label") == "Grand Total")
    ).to_dicts()[0]
    assert revenue_total["amount"] == 163_408.05
    assert refund_total["amount"] == 10_114.56

    sales = result.sales
    assert sales.height == 159
    by_channel = {
        row["channel"]: round(row["net_sales"], 2)
        for row in sales.group_by("channel")
        .agg(pl.col("net_sales").sum().alias("net_sales"))
        .to_dicts()
    }
    assert by_channel == {"DTC": 146_599.00, "WS": 16_809.05}


def test_extracts_refunds_and_can_exclude_rows_without_amount() -> None:
    result = extract_monthly_revenue(SAMPLE)
    refunds = result.refunds

    assert refunds.height == 19
    assert refunds.filter(pl.col("has_amount")).height == 16
    assert round(refunds.select(pl.col("amount").sum()).item(), 2) == 10_114.56

    by_division = {
        row["division"]: round(row["amount"], 2)
        for row in refunds.filter(pl.col("has_amount"))
        .group_by("division")
        .agg(pl.col("amount").sum().alias("amount"))
        .to_dicts()
    }
    assert by_division == {"OG-DTC": 9_940.81, "OG-WS": 173.75}

    amount_only = extract_monthly_revenue(
        SAMPLE,
        MonthlyRevenueConfig(include_refund_rows_without_amount=False),
    )
    assert amount_only.refunds.height == 16


def test_extracts_coupon_rows() -> None:
    result = extract_monthly_revenue(SAMPLE)
    coupons = result.coupons

    assert coupons.height == 19
    assert round(coupons.select(pl.col("total").sum()).item(), 2) == 17_212.05

    order = coupons.filter(pl.col("order") == "#930619").to_dicts()[0]
    assert order["date"] == "2026-03-23"
    assert order["customer"] == "Alexa Harris-Ralff"
    assert order["total"] == 982.5
