from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.payroll_journal import (
    PayrollJournalConfig,
    extract_payroll_journal,
    write_default_config,
)


app = typer.Typer(help="Extract payroll journal workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to a payroll journal workbook."),
    output_dir: Path | None = typer.Option(None, "--output-dir", "-o", help="Write output files."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    no_calculate_formulas: bool = typer.Option(
        False,
        "--no-calculate-formulas",
        help="Disable default in-memory formula evaluation.",
    ),
    preserve_zero_allocations: bool = typer.Option(
        False,
        "--preserve-zero-allocations",
        help="Keep allocation rows whose amount is zero.",
    ),
    use_distribution_sheet: bool = typer.Option(
        False,
        "--use-distribution-sheet",
        help="Use the intermediary Payroll Distribution sheet instead of deriving distribution from Payroll.",
    ),
) -> None:
    if init_config is not None:
        write_default_config(init_config)
        console.print(f"Wrote default config to {init_config}")
        raise typer.Exit()

    extractor_config = PayrollJournalConfig.from_json_file(config)
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False
    if preserve_zero_allocations:
        extractor_config.preserve_zero_allocations = True
    if use_distribution_sheet:
        extractor_config.derive_distribution_from_payroll_sheet = False

    result = extract_payroll_journal(path, extractor_config)
    employees = result.employees
    allocations = result.allocations
    allocation_summaries = result.allocation_summaries
    distribution = result.distribution

    console.print(f"File: {path}")
    console.print(f"Payroll sheet: {result.payroll_sheet}")
    console.print(f"Distribution sheet: {result.distribution_sheet or '(not found)'}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Payroll Journal Extraction")
    table.add_column("Table")
    table.add_column("Rows", justify="right")
    table.add_column("Amount", justify="right")
    table.add_row("employees", str(employees.height), f"{sum_column(employees, 'gross_pay'):,.2f}")
    table.add_row("allocations", str(allocations.height), f"{sum_column(allocations, 'amount'):,.2f}")
    table.add_row(
        "allocation_summaries",
        str(allocation_summaries.height),
        f"{sum_column(allocation_summaries, 'amount'):,.2f}",
    )
    table.add_row("distribution", str(distribution.height), f"{sum_column(distribution, 'amount'):,.2f}")
    console.print(table)

    if not employees.is_empty():
        by_section = (
            employees.group_by("section")
            .agg(pl.col("gross_pay").sum().alias("gross_pay"))
            .sort("section")
        )
        console.print(by_section)

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(employees, output_dir / f"payroll_employees.{format}", format)
        write_frame(allocations, output_dir / f"payroll_allocations.{format}", format)
        write_frame(allocation_summaries, output_dir / f"payroll_allocation_summaries.{format}", format)
        write_frame(distribution, output_dir / f"payroll_distribution.{format}", format)
        console.print(f"Wrote outputs to {output_dir}")


def sum_column(df: pl.DataFrame, column: str) -> float:
    if df.is_empty() or column not in df.columns:
        return 0.0
    value = df.select(pl.col(column).fill_null(0).sum()).item()
    return float(value or 0)


def write_frame(df: pl.DataFrame, path: Path, format: str) -> None:
    match format.lower():
        case "csv":
            df.write_csv(path)
        case "json":
            path.write_text(df.write_json(), encoding="utf-8")
        case "parquet":
            df.write_parquet(path)
        case other:
            raise typer.BadParameter(f"Unsupported format: {other}")


if __name__ == "__main__":
    app()
