from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.monthly_revenue import (
    MonthlyRevenueConfig,
    extract_monthly_revenue,
    write_default_config,
)


app = typer.Typer(help="Extract DTC/WS monthly revenue workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to a monthly revenue workbook."),
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
    exclude_refund_rows_without_amount: bool = typer.Option(
        False,
        "--exclude-refund-rows-without-amount",
        help="Drop refund detail rows that do not carry an amount.",
    ),
) -> None:
    if init_config is not None:
        write_default_config(init_config)
        console.print(f"Wrote default config to {init_config}")
        raise typer.Exit()

    extractor_config = MonthlyRevenueConfig.from_json_file(config)
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False
    if exclude_refund_rows_without_amount:
        extractor_config.include_refund_rows_without_amount = False

    result = extract_monthly_revenue(path, extractor_config)
    summary = result.summary
    sales = result.sales
    refunds = result.refunds
    coupons = result.coupons

    console.print(f"File: {path}")
    console.print(f"Sheets: {result.sheets}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Monthly Revenue Extraction")
    table.add_column("Table")
    table.add_column("Rows", justify="right")
    table.add_column("Amount", justify="right")
    table.add_row("summary", str(summary.height), f"{sum_column(summary, 'amount'):,.2f}")
    table.add_row("sales", str(sales.height), f"{sum_column(sales, 'net_sales'):,.2f}")
    table.add_row("refunds", str(refunds.height), f"{sum_column(refunds, 'amount'):,.2f}")
    table.add_row("coupons", str(coupons.height), f"{sum_column(coupons, 'total'):,.2f}")
    console.print(table)

    if not sales.is_empty():
        console.print(
            sales.group_by("channel")
            .agg(pl.col("net_sales").sum().alias("net_sales"))
            .sort("channel")
        )
    if not refunds.is_empty():
        console.print(
            refunds.filter(pl.col("has_amount"))
            .group_by("division")
            .agg(pl.col("amount").sum().alias("amount"))
            .sort("division")
        )

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(summary, output_dir / f"monthly_revenue_summary.{format}", format)
        write_frame(sales, output_dir / f"monthly_revenue_sales.{format}", format)
        write_frame(refunds, output_dir / f"monthly_revenue_refunds.{format}", format)
        write_frame(coupons, output_dir / f"monthly_revenue_coupons.{format}", format)
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
