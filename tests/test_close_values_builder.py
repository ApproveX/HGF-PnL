from pathlib import Path

from hgf_pnl.extractors.br_info import extract_br_info
from hgf_pnl.extractors.chargeback_pdf import extract_chargeback_pdf
from hgf_pnl.extractors.division_cogs import extract_division_cogs
from hgf_pnl.extractors.monthly_revenue import extract_monthly_revenue
from hgf_pnl.extractors.payroll_journal import extract_payroll_journal
from hgf_pnl.extractors.pl_by_dept import extract_pl_by_dept
from hgf_pnl.extractors.th_revenue import extract_th_revenue
from hgf_pnl.pipeline.close_values import build_consolidated_values, flatten_values


WORKPAPERS = Path("sample_files/Workpapers MARCH/DATA")


def test_builds_march_values_with_accountant_corrections() -> None:
    pl_by_dept = extract_pl_by_dept(WORKPAPERS / "Profit and Loss By Dept.xlsx").to_polars()
    br_info = extract_br_info(Path("sample_files/BR Info.xlsx")).overrides
    monthly_revenue = extract_monthly_revenue(
        WORKPAPERS / "DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx"
    )
    division_cogs = extract_division_cogs(
        WORKPAPERS / "INTERNAL - Division COGS 2019 - Current (26).xlsx"
    )
    th_revenue = extract_th_revenue(WORKPAPERS / "TH March 2026 Revenue Report.xlsx")
    payroll = extract_payroll_journal(Path("sample_files/PAYROLL/Payroll Journal_March 2026.xlsx"))
    chargeback = extract_chargeback_pdf(WORKPAPERS / "- OG _ Chargeback Report - 03. March 2026.pdf")

    result = build_consolidated_values(
        pl_by_dept=pl_by_dept,
        br_info=br_info,
        monthly_revenue_summary=monthly_revenue.summary,
        monthly_revenue_sales=monthly_revenue.sales,
        monthly_revenue_refunds=monthly_revenue.refunds,
        division_cogs_matrix=division_cogs.matrix,
        th_revenue_summary=th_revenue.account_summary,
        payroll_employees=payroll.employees,
        payroll_allocation_summaries=payroll.allocation_summaries,
        payroll_distribution=payroll.distribution,
        chargeback_customer_detail=chargeback.customer_detail,
        year=2026,
        month_num=3,
    )

    assert result.warnings == []
    values = flatten_values(result.values)
    assert len(values) >= 120

    assert round(values["raw_master.gl.consulting_expense"], 2) == 51_295.21
    assert round(values["raw_master.consulting.corp"], 2) == 51_295.21
    assert values["raw_master.gl.hr_recruiting"] == 5_317.87
    assert "raw_master.gl.travel" not in values
    assert "raw_master.gl.advertising_marketing" not in values

    assert values["raw_master.gl.bank_fees"] == 799.0
    assert values["raw_master.gl.bank_fees_adjustment"] == 0.0
    assert values["raw_master.gl.merchant_account_fees"] == 4_619.0
    assert values["raw_master.gl.merchant_account_fees_adjustment"] == 0.0
    assert values["raw_master.gl.equipment_lease"] == 12_172.0
    assert values["raw_master.gl.equipment_lease_adjustment"] == 0.0

    assert round(values["raw_master.sales.dtc"], 2) == 163_408.05
    assert values["raw_master.sales.online"] == 542_805.0
    assert values["raw_master.returns.dtc"] == -10_114.56
    assert values["raw_master.returns.trend_house"] == -2_011.0

    assert round(values["raw_cogs.current_month.cogs.online_usa"], 2) == 190_442.49
    assert values["raw_cogs.current_month.cogs.online"] == 0.0
    assert values["raw_cogs.shipping_actual.online"] == 887.53
    assert values["raw_cogs.shipping_actual.online_usa"] == 2_636.24
    assert values["raw_cogs.current_month.cogs.og_dtc"] == 19_457.04
    assert values["raw_cogs.current_month.cogs.og_dtc_returns"] == 1_713.3
    assert values["raw_cogs.shipping_for_samples.current_month"] == 307.46
    assert round(values["raw_cogs.tariffs"], 2) == 73_451.30

    assert round(values["raw_payroll.production"], 2) == 64_635.26
    assert values["raw_payroll.lital_allocation.corp"] == 6_948.0
    assert values["raw_payroll.allocation_breakdowns.art.trend_house"] == 13_007.702
    assert values["raw_payroll.allocation_breakdowns.it.general"] == 13_384.614
