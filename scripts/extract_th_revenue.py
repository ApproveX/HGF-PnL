from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.th_revenue import (
    THRevenueConfig,
    extract_th_revenue,
    write_default_config,
)


app = typer.Typer(help="Extract Trend House revenue report workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to a TH revenue report workbook."),
    output_dir: Path | None = typer.Option(None, "--output-dir", "-o", help="Write output files."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    no_totals: bool = typer.Option(False, "--no-totals", help="Exclude total rows."),
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

    extractor_config = THRevenueConfig.from_json_file(config)
    if no_totals:
        extractor_config.include_total_rows = False
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False

    result = extract_th_revenue(path, extractor_config)

    console.print(f"File: {path}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Extracted Sheets")
    table.add_column("Role")
    table.add_column("Sheet")
    table.add_column("Header Row", justify="right")
    table.add_column("Rows", justify="right")
    table.add_column("Revenue", justify="right")
    table.add_column("Total Cost", justify="right")
    for role, sheet in result.sheets.items():
        df = sheet.to_polars()
        non_total = df.filter(~pl.col("is_total_row"))
        revenue = sum_column(non_total, "revenue")
        total_cost = sum_column(non_total, "total_cost")
        table.add_row(
            role,
            sheet.sheet_name,
            str(sheet.header_row),
            str(df.height),
            f"{revenue:,.2f}",
            f"{total_cost:,.2f}",
        )
    console.print(table)

    summary = result.account_summary
    po_details = result.po_details
    usa_stock = result.usa_stock
    console.print(f"Account summary rows: {summary.height}")
    console.print(f"PO detail rows: {po_details.height}")
    console.print(f"USA stock rows: {usa_stock.height}")
    console.print(
        "PO detail revenue excluding totals: "
        f"{sum_column(po_details.filter(~pl.col('is_total_row')), 'revenue'):,.2f}"
    )

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(summary, output_dir / f"th_revenue_summary.{format}", format)
        write_frame(po_details, output_dir / f"th_revenue_po_details.{format}", format)
        write_frame(usa_stock, output_dir / f"th_revenue_usa_stock.{format}", format)
        write_frame(result.all_rows(), output_dir / f"th_revenue_all_rows.{format}", format)
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
