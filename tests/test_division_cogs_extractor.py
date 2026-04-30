from pathlib import Path

import polars as pl

from hgf_pnl.extractors.division_cogs import extract_division_cogs


SAMPLE = Path("sample_files/Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx")


def test_extracts_division_cogs_year_matrix() -> None:
    result = extract_division_cogs(SAMPLE)
    matrix = result.matrix

    assert result.year_sheets == ["2018", "2019", "2020", "2021", "2022", "2023", "2026", "2025", "2024"]
    assert not result.warnings
    assert matrix.height == 3276

    march_2026_total = matrix.filter(
        (pl.col("year") == 2026)
        & (pl.col("month_num") == 3)
        & (pl.col("type") == "COGS")
        & (pl.col("channel") == "Total")
    ).to_dicts()[0]
    assert round(march_2026_total["amount"], 2) == 213_273.23
    assert march_2026_total["source_cell"] == "U16"
    assert march_2026_total["calculation_status"] == "ok"

    online_usa = matrix.filter(
        (pl.col("year") == 2026)
        & (pl.col("month_num") == 3)
        & (pl.col("type") == "COGS")
        & (pl.col("channel") == "Online - USA")
    ).to_dicts()[0]
    assert round(online_usa["amount"], 2) == 181_629.67


def test_extracts_division_cogs_partner_details() -> None:
    result = extract_division_cogs(SAMPLE)
    details = result.partner_details

    assert result.partner_detail_sheets == [
        "2022 Partner Details",
        "2023 Partner Details",
        "2026 Partner Details",
        "2025 Partner Details",
        "2024 Partner Details",
    ]
    assert details.height == 6020

    d2c = {
        row["measure"]: row["amount"]
        for row in details.filter(
            (pl.col("year") == 2026)
            & (pl.col("month_num") == 3)
            & (pl.col("partner") == "D2C")
        ).to_dicts()
    }
    assert d2c == {
        "cogs": 19_658.51,
        "material_cost": 15_596.9,
        "labor_cost": 4_061.61,
    }


def test_preserves_unsupported_partner_detail_formula_status() -> None:
    result = extract_division_cogs(SAMPLE)
    details = result.partner_details

    row = details.filter(
        (pl.col("sheet") == "2024 Partner Details")
        & (pl.col("month_num") == 3)
        & (pl.col("partner") == "Amazon DS - USA")
        & (pl.col("measure") == "cogs")
    ).to_dicts()[0]

    assert row["amount"] is None
    assert row["raw_value"] == "#N/A"
    assert row["formula"].startswith("=VLOOKUP")
    assert row["calculation_status"] == "unsupported"
    assert row["calculation_detail"] == "VLOOKUP"
