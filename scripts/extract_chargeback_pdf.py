from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.chargeback_pdf import (
    ChargebackPDFConfig,
    extract_chargeback_pdf,
    write_default_config,
)


app = typer.Typer(help="Extract OG chargeback report email PDFs.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to chargeback report PDF."),
    output_dir: Path | None = typer.Option(None, "--output-dir", "-o", help="Write output files."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
) -> None:
    if init_config is not None:
        write_default_config(init_config)
        console.print(f"Wrote default config to {init_config}")
        raise typer.Exit()

    extractor_config = ChargebackPDFConfig.from_json_file(config)
    result = extract_chargeback_pdf(path, extractor_config)
    monthly = result.monthly_summary
    customer_detail = result.customer_detail
    reconciliation = result.reconciliation

    console.print(f"File: {path}")
    console.print(f"Subject: {result.subject or '(not detected)'}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Chargeback Extraction")
    table.add_column("Table")
    table.add_column("Rows", justify="right")
    table.add_column("Key Amount", justify="right")
    table.add_row("monthly_summary", str(monthly.height), f"{target_month_total(monthly, extractor_config):,.2f}")
    table.add_row("customer_detail", str(customer_detail.height), f"{named_total(customer_detail, 'Grand Total', 'amount'):,.2f}")
    table.add_row(
        "reconciliation",
        str(reconciliation.height),
        f"{named_total(reconciliation, 'Grand Total', 'difference'):,.2f}",
    )
    console.print(table)

    current = monthly.filter(
        (pl.col("year") == extractor_config.target_year)
        & (pl.col("month_name") == extractor_config.target_month_name)
    )
    if not current.is_empty():
        console.print(current.select(["category", "amount", "percent_of_total", "source_page", "source_line"]))

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(monthly, output_dir / f"chargeback_monthly_summary.{format}", format)
        write_frame(customer_detail, output_dir / f"chargeback_customer_detail.{format}", format)
        write_frame(reconciliation, output_dir / f"chargeback_reconciliation.{format}", format)
        if result.notes:
            write_frame(pl.DataFrame(result.notes), output_dir / f"chargeback_notes.{format}", format)
        console.print(f"Wrote outputs to {output_dir}")


def sum_column(df: pl.DataFrame, column: str) -> float:
    if df.is_empty() or column not in df.columns:
        return 0.0
    value = df.select(pl.col(column).fill_null(0).sum()).item()
    return float(value or 0)


def target_month_total(df: pl.DataFrame, config: ChargebackPDFConfig) -> float:
    if df.is_empty():
        return 0.0
    filtered = df.filter(
        (pl.col("year") == config.target_year)
        & (pl.col("month_name") == config.target_month_name)
        & (pl.col("category") == "grand_total")
    )
    if filtered.is_empty():
        return 0.0
    return float(filtered.select(pl.col("amount")).item())


def named_total(df: pl.DataFrame, name: str, amount_column: str) -> float:
    if df.is_empty() or "customer" not in df.columns:
        return 0.0
    filtered = df.filter(pl.col("customer") == name)
    if filtered.is_empty():
        return 0.0
    return float(filtered.select(pl.col(amount_column)).item())


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
