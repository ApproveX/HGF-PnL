from pathlib import Path

import polars as pl

from hgf_pnl.extractors.payroll_journal import extract_payroll_journal


SAMPLE = Path("sample_files/PAYROLL/Payroll Journal_March 2026.xlsx")


def test_extracts_payroll_employee_sections() -> None:
    result = extract_payroll_journal(SAMPLE)
    employees = result.employees

    assert result.payroll_sheet == "Payroll"
    assert result.distribution_sheet == "Payroll Distribution"
    assert employees.height == 43
    assert round(employees.select(pl.col("gross_pay").sum()).item(), 2) == 241_161.25

    by_section = {
        row["section"]: round(row["gross_pay"], 2)
        for row in employees.group_by("section")
        .agg(pl.col("gross_pay").sum().alias("gross_pay"))
        .to_dicts()
    }
    assert by_section == {
        "Production": 64_635.26,
        "Sales Dept": 44_909.04,
        "IT": 15_846.15,
        "Art": 32_961.56,
        "Corp": 64_655.38,
        "Ops": 18_153.86,
    }


def test_extracts_employee_allocation_rows() -> None:
    result = extract_payroll_journal(SAMPLE)
    allocations = result.allocations

    assert allocations.height == 24
    assert round(allocations.select(pl.col("amount").sum()).item(), 2) == 93_716.75

    alexis = allocations.filter(
        (pl.col("employee_name") == "Koczwara Alexis Buffy")
        & (pl.col("allocation_category") == "Trend House")
    ).to_dicts()[0]
    assert alexis["amount"] == 4307.688
    assert alexis["percent_of_gross"] == 0.7
    assert alexis["calculation_status"] == "ok"

    employees = result.employees
    allocated = employees.filter(pl.col("allocated_total").is_not_null())
    assert allocated.height == 14
    assert allocated.select(pl.col("allocation_difference").abs().max()).item() == 0


def test_extracts_department_allocation_summary_totals() -> None:
    result = extract_payroll_journal(SAMPLE)
    summaries = result.allocation_summaries

    assert summaries.height == 13

    art_th = summaries.filter(
        (pl.col("department") == "Art") & (pl.col("allocation_category") == "TH")
    ).to_dicts()[0]
    assert art_th["amount"] == 13_007.702
    assert art_th["source_cell"] == "J47"
    assert art_th["source_kind"] == "allocation_total_row"
    assert art_th["calculation_status"] == "ok"

    art_general = summaries.filter(
        (pl.col("department") == "Art") & (pl.col("allocation_category") == "General")
    ).to_dicts()[0]
    assert art_general["amount"] == 5_923.084
    assert art_general["source_cell"] == "M47"

    it_online = summaries.filter(
        (pl.col("department") == "IT") & (pl.col("allocation_category") == "Online")
    ).to_dicts()[0]
    assert it_online["amount"] == 1_230.768
    assert it_online["source_cell"] == "H38"

    totals = summaries.filter(pl.col("is_total_category"))
    by_department = {
        row["department"]: round(row["amount"], 2)
        for row in totals.select("department", "amount").to_dicts()
    }
    assert by_department == {
        "Sales Dept": 44_909.04,
        "IT": 15_846.15,
        "Art": 32_961.56,
    }


def test_extracts_distribution_and_workbook_index_formula_refs() -> None:
    result = extract_payroll_journal(SAMPLE)
    distribution = result.distribution

    assert distribution.height == 17
    corp_total = distribution.filter(
        (pl.col("block") == "Payroll Corp") & (pl.col("label") == "Total")
    ).to_dicts()[0]
    assert round(corp_total["amount"], 2) == 241_161.25
    assert corp_total["calculation_status"] == "derived"

    th = distribution.filter(
        (pl.col("block") == "Lital Allocation in G&A Exp") & (pl.col("label") == "TH")
    ).to_dicts()[0]
    assert th["formula"] == "=+C53*L57"
    assert th["cached_amount"] == 1544.0
    assert th["calculated_amount"] == 1544.0
    assert th["amount"] == 1544.0
    assert th["calculation_status"] == "ok"
    assert th["source_cell"] == "M57"

    corp = distribution.filter(
        (pl.col("block") == "Lital Allocation in G&A Exp") & (pl.col("label") == "CORP")
    ).to_dicts()[0]
    assert corp["cached_amount"] == 6948.0
    assert corp["calculated_amount"] == 6948.0
    assert corp["amount"] == 6948.0
    assert corp["source_cell"] == "M59"

    total = distribution.filter(
        (pl.col("block") == "Lital Allocation in G&A Exp") & (pl.col("label") == "Total")
    ).to_dicts()[0]
    assert total["amount"] == 15_440.0
    assert total["calculation_status"] == "derived"
