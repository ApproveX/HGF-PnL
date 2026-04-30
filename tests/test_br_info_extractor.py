from pathlib import Path

import polars as pl

from hgf_pnl.extractors.br_info import BRInfoConfig, extract_br_info


SAMPLE = Path("sample_files/BR Info.xlsx")


def test_extracts_br_info_manual_march_overrides() -> None:
    result = extract_br_info(SAMPLE)
    overrides = result.overrides

    assert result.sheet_name == "2026"
    assert result.header_row == 2
    assert result.year == 2026
    assert not result.warnings
    assert overrides.height == 9
    assert set(overrides["month_name"]) == {"March"}
    assert round(overrides.select(pl.col("value").sum()).item(), 2) == 578_002.00

    values = {row["override_name"]: row for row in overrides.to_dicts()}
    assert values["AllPopArt Sales"]["value"] == 1833.0
    assert values["AllPopArt Sales"]["source_cell"] == "D3"
    assert values["AllPopArt Returns and Allowances"]["value"] == -346.0
    assert values["Employee Benefits"]["value"] == 10898.0
    assert values["Equipment Leasing"]["value"] == 12172.0
    assert values["LOC Interest"]["value"] == 3170.0
    assert values["Online Sales"]["value"] == 542805.0


def test_can_include_blank_month_values() -> None:
    result = extract_br_info(SAMPLE, BRInfoConfig(include_blank_values=True))
    overrides = result.overrides

    assert overrides.height == 108
    march = overrides.filter(pl.col("month_name") == "March")
    april = overrides.filter(pl.col("month_name") == "April")
    assert march.select(pl.col("value").is_not_null().sum()).item() == 9
    assert april.select(pl.col("value").is_null().sum()).item() == 9
