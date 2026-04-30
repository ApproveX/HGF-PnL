from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.division_cogs import (
    DivisionCOGSConfig,
    extract_division_cogs,
    write_default_config,
)


app = typer.Typer(help="Extract Division COGS workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to Division COGS workbook."),
    output_dir: Path | None = typer.Option(None, "--output-dir", "-o", help="Write output files."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    no_totals: bool = typer.Option(False, "--no-totals", help="Exclude total columns from year matrix rows."),
    include_zero_amounts: bool = typer.Option(False, "--include-zero-amounts", help="Keep zero-value rows."),
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

    extractor_config = DivisionCOGSConfig.from_json_file(config)
    if no_totals:
        extractor_config.include_total_columns = False
    if include_zero_amounts:
        extractor_config.include_zero_amounts = True
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False

    result = extract_division_cogs(path, extractor_config)
    matrix = result.matrix
    partner_details = result.partner_details

    console.print(f"File: {path}")
    console.print(f"Year sheets: {', '.join(result.year_sheets)}")
    console.print(f"Partner detail sheets: {', '.join(result.partner_detail_sheets)}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Division COGS Extraction")
    table.add_column("Table")
    table.add_column("Rows", justify="right")
    table.add_column("Amount", justify="right")
    table.add_row("matrix", str(matrix.height), f"{sum_column(matrix, 'amount'):,.2f}")
    table.add_row("partner_details", str(partner_details.height), f"{sum_column(partner_details, 'amount'):,.2f}")
    console.print(table)

    if not matrix.is_empty():
        current_cogs = matrix.filter((pl.col("year") == 2026) & (pl.col("type") == "COGS"))
        if not current_cogs.is_empty():
            console.print(
                current_cogs.group_by("month_name")
                .agg(
                    pl.col("month_num").min().alias("month_num"),
                    pl.col("amount").sum().alias("amount"),
                )
                .sort("month_num")
                .select(["month_name", "amount"])
            )

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(matrix, output_dir / f"division_cogs_matrix.{format}", format)
        write_frame(partner_details, output_dir / f"division_cogs_partner_details.{format}", format)
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
