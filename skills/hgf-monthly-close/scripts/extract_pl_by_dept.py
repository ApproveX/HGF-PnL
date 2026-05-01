from __future__ import annotations

from pathlib import Path

import typer
import polars as pl
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.pl_by_dept import (
    PLByDeptConfig,
    extract_pl_by_dept,
    write_default_config,
)


app = typer.Typer(help="Extract a department matrix P&L workbook.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to a Profit and Loss by Department workbook."),
    output: Path | None = typer.Option(None, "--output", "-o", help="Write extracted rows."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    no_totals: bool = typer.Option(False, "--no-totals", help="Exclude total columns."),
    no_calculate_formulas: bool = typer.Option(
        False,
        "--no-calculate-formulas",
        help="Disable default in-memory formula evaluation.",
    ),
) -> None:
    if init_config is not None:
        write_default_config(init_config)
        console.print(f"Wrote default config to {init_config}")
        raise typer.Exit()

    extractor_config = PLByDeptConfig.from_json_file(config)
    if no_totals:
        extractor_config.include_total_columns = False
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False

    result = extract_pl_by_dept(path, extractor_config)
    df = result.to_polars()

    console.print(f"File: {path}")
    console.print(f"Sheet: {result.sheet_name}")
    console.print(f"Title: {result.report_title or '(not detected)'}")
    console.print(f"Period: {result.report_period or '(not detected)'}")
    console.print(f"Header row: {result.header_row}")
    console.print(f"Line item column: {result.line_item_column}")
    console.print(f"Departments: {', '.join(result.department_columns.values())}")
    console.print(f"Rows extracted: {df.height}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    summary = (
        df.group_by("department")
        .agg(
            [
                pl.col("amount").sum().alias("amount_sum"),
                pl.col("line_item").n_unique().alias("line_items"),
            ]
        )
        .sort("department")
    )
    table = Table(title="Department Summary")
    table.add_column("Department")
    table.add_column("Line Items", justify="right")
    table.add_column("Amount Sum", justify="right")
    for row in summary.iter_rows(named=True):
        table.add_row(
            str(row["department"]),
            str(row["line_items"]),
            f'{row["amount_sum"]:,.2f}',
        )
    console.print(table)

    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        match format.lower():
            case "csv":
                df.write_csv(output)
            case "json":
                output.write_text(df.write_json(), encoding="utf-8")
            case "parquet":
                df.write_parquet(output)
            case other:
                raise typer.BadParameter(f"Unsupported format: {other}")
        console.print(f"Wrote {output}")


if __name__ == "__main__":
    app()
