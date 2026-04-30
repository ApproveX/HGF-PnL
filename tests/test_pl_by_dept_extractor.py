from pathlib import Path

from hgf_pnl.extractors.pl_by_dept import PLByDeptConfig, extract_pl_by_dept


SAMPLE = Path("sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx")


def test_extracts_profit_and_loss_by_department_sample() -> None:
    result = extract_pl_by_dept(SAMPLE)

    assert result.sheet_name == "Profit and Loss by Department"
    assert result.header_row == 5
    assert result.line_item_column == 1
    assert result.report_title == "Hotgoldfish Corp | Profit and Loss by Department"
    assert result.report_period == "March 2026"
    assert "Corporate Dept" in result.department_columns.values()
    assert "Total" in result.department_columns.values()

    rows = result.rows
    assert rows

    apa_sales = [
        row
        for row in rows
        if row["line_item"] == "Sales - APA" and row["department"] == "All Pop Art"
    ]
    assert apa_sales
    assert apa_sales[0]["amount"] == 2786.55
    assert apa_sales[0]["cached_amount"] == 0
    assert apa_sales[0]["calculated_amount"] == 2786.55
    assert apa_sales[0]["amount_cell"] == "B8"
    assert apa_sales[0]["formula"] == "=2786.55"

    total_sales = [
        row
        for row in rows
        if row["line_item"] == "Total Sales" and row["department"] == "Total"
    ]
    assert total_sales
    assert total_sales[0]["is_total_column"] is True


def test_can_exclude_total_columns() -> None:
    result = extract_pl_by_dept(SAMPLE, PLByDeptConfig(include_total_columns=False))

    departments = set(result.department_columns.values())
    assert "Total" not in departments
    assert "Total Z-COMPANY" not in departments
    assert all(row["department"] not in {"Total", "Total Z-COMPANY"} for row in result.rows)
