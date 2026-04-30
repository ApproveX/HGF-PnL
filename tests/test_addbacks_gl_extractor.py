from pathlib import Path

import polars as pl

from hgf_pnl.extractors.addbacks_gl import (
    AddbacksGLConfig,
    DeclaredTotal,
    RowGroupRule,
    extract_addbacks_gl,
)


SAMPLE = Path("sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx")


def test_extracts_reviewed_gl_addbacks_from_comments() -> None:
    result = extract_addbacks_gl(
        SAMPLE,
        AddbacksGLConfig(
            declared_totals=[DeclaredTotal(group_name="addbacks", amount=23_195.00)]
        ),
    )

    assert result.sheet_name == "NEW MONTH"
    assert result.header_row == 5
    assert result.header_map["expected_account"] == 10
    assert result.header_map["expected_department"] == 11
    assert result.header_map["comments"] == 12

    addbacks = result.group("addbacks")
    assert addbacks.height == 122
    assert round(addbacks.select(pl.col("amount").sum()).item(), 2) == 23_195.16

    reconciliation = result.reconciliations[0]
    assert reconciliation["status"] == "ok"
    assert round(reconciliation["difference"], 2) == 0.16


def test_extracts_email_instruction_color_groups() -> None:
    result = extract_addbacks_gl(SAMPLE)

    summaries = {row["group_name"]: row for row in result.group_summaries}
    assert summaries["red_addback_color_rows"]["row_count"] == 123
    assert round(summaries["red_addback_color_rows"]["amount_total"], 2) == 23_248.67
    assert summaries["unknown_charges"]["row_count"] == 1
    assert round(summaries["unknown_charges"]["amount_total"], 2) == 213.90
    assert summaries["account_department_edits"]["row_count"] == 139
    assert round(summaries["account_department_edits"]["amount_total"], 2) == 49_765.75

    unknown = result.group("unknown_charges").to_dicts()[0]
    assert unknown["account_section"] == "6400 Software & Web Services"
    assert unknown["account"] == "Accounts Payable"
    assert unknown["memo_description"] == "INV_CA_325973"

    assert "Red/pink row total differs from comment-based addback total by 53.51" in result.warnings


def test_agent_can_override_row_group_rules() -> None:
    result = extract_addbacks_gl(
        SAMPLE,
        AddbacksGLConfig(
            row_group_rules=[
                RowGroupRule(
                    name="yellow_with_expected_department",
                    match_mode="all",
                    fill_colors=["FFFFFF00"],
                    nonblank_columns=["expected_department"],
                )
            ]
        ),
    )

    group = result.group("yellow_with_expected_department")
    assert group.height == 131
    assert group.select(pl.col("expected_department").str.len_chars().min()).item() > 0
