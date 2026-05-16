from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.pipeline.close_values import (
    build_consolidated_values,
    read_frame,
    write_values_json,
)


app = typer.Typer(help="Build consolidated_values.json from reviewed HGF extractor outputs.")
console = Console()


@app.command()
def build(
    output: Path = typer.Argument(..., help="Output consolidated values JSON."),
    year: int | None = typer.Option(None, "--year", help="Optional year filter."),
    month_num: int | None = typer.Option(None, "--month-num", help="Optional month number filter."),
    pl_by_dept: Path | None = typer.Option(None, "--pl-by-dept", help="pl_by_dept CSV/JSON/Parquet."),
    br_info: Path | None = typer.Option(None, "--br-info", help="br_info CSV/JSON/Parquet."),
    monthly_revenue_summary: Path | None = typer.Option(
        None,
        "--monthly-revenue-summary",
        help="monthly_revenue_summary CSV/JSON/Parquet.",
    ),
    monthly_revenue_sales: Path | None = typer.Option(
        None,
        "--monthly-revenue-sales",
        help="monthly_revenue_sales CSV/JSON/Parquet.",
    ),
    monthly_revenue_refunds: Path | None = typer.Option(
        None,
        "--monthly-revenue-refunds",
        help="monthly_revenue_refunds CSV/JSON/Parquet.",
    ),
    division_cogs_matrix: Path | None = typer.Option(
        None,
        "--division-cogs-matrix",
        help="division_cogs_matrix CSV/JSON/Parquet.",
    ),
    th_revenue_summary: Path | None = typer.Option(
        None,
        "--th-revenue-summary",
        help="th_revenue_summary CSV/JSON/Parquet.",
    ),
    payroll_employees: Path | None = typer.Option(
        None,
        "--payroll-employees",
        help="payroll_employees CSV/JSON/Parquet.",
    ),
    payroll_allocation_summaries: Path | None = typer.Option(
        None,
        "--payroll-allocation-summaries",
        help="payroll_allocation_summaries CSV/JSON/Parquet.",
    ),
    payroll_distribution: Path | None = typer.Option(
        None,
        "--payroll-distribution",
        help="payroll_distribution CSV/JSON/Parquet.",
    ),
    chargeback_customer_detail: Path | None = typer.Option(
        None,
        "--chargeback-customer-detail",
        help="chargeback_customer_detail CSV/JSON/Parquet.",
    ),
) -> None:
    result = build_consolidated_values(
        pl_by_dept=read_frame(pl_by_dept),
        br_info=read_frame(br_info),
        monthly_revenue_summary=read_frame(monthly_revenue_summary),
        monthly_revenue_sales=read_frame(monthly_revenue_sales),
        monthly_revenue_refunds=read_frame(monthly_revenue_refunds),
        division_cogs_matrix=read_frame(division_cogs_matrix),
        th_revenue_summary=read_frame(th_revenue_summary),
        payroll_employees=read_frame(payroll_employees),
        payroll_allocation_summaries=read_frame(payroll_allocation_summaries),
        payroll_distribution=read_frame(payroll_distribution),
        chargeback_customer_detail=read_frame(chargeback_customer_detail),
        year=year,
        month_num=month_num,
    )
    write_values_json(result, output)

    table = Table(title="Consolidated Values")
    table.add_column("Metric")
    table.add_column("Count", justify="right")
    table.add_row("populated_keys", str(result.populated_key_count))
    table.add_row("warnings", str(len(result.warnings)))
    console.print(table)
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")
    console.print(f"Wrote {output}")


if __name__ == "__main__":
    app()
