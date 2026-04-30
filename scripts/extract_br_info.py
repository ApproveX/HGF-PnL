from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.br_info import (
    BRInfoConfig,
    extract_br_info,
    write_default_config,
)


app = typer.Typer(help="Extract BR Info manual override workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to a BR Info workbook."),
    output: Path | None = typer.Option(None, "--output", "-o", help="Write output file."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    include_blank_values: bool = typer.Option(
        False,
        "--include-blank-values",
        help="Keep blank month/override cells.",
    ),
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

    extractor_config = BRInfoConfig.from_json_file(config)
    if include_blank_values:
        extractor_config.include_blank_values = True
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False

    result = extract_br_info(path, extractor_config)
    overrides = result.overrides

    console.print(f"File: {path}")
    console.print(f"Sheet: {result.sheet_name}")
    console.print(f"Header row: {result.header_row}")
    console.print(f"Year: {result.year or '(not detected)'}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="BR Info Overrides")
    table.add_column("Rows", justify="right")
    table.add_column("Months")
    table.add_column("Value Total", justify="right")
    months = ", ".join(overrides.select("month_name").unique().sort("month_name").to_series().to_list())
    table.add_row(str(overrides.height), months, f"{sum_column(overrides, 'value'):,.2f}")
    console.print(table)
    if not overrides.is_empty():
        console.print(overrides.select(["year", "month_name", "override_name", "value", "source_cell"]))

    if output is not None:
        output.parent.mkdir(parents=True, exist_ok=True)
        write_frame(overrides, output, format)
        console.print(f"Wrote {output}")


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
