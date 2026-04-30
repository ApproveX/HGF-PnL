from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.addbacks_gl import (
    AddbacksGLConfig,
    DeclaredTotal,
    extract_addbacks_gl,
    write_default_config,
)


app = typer.Typer(help="Extract reviewed GL addback workbooks.")
console = Console()


@app.command()
def extract(
    path: Path = typer.Argument(..., help="Path to reviewed GL workbook."),
    output_dir: Path | None = typer.Option(None, "--output-dir", "-o", help="Write output files."),
    format: str = typer.Option("csv", "--format", help="csv, json, or parquet."),
    config: Path | None = typer.Option(None, "--config", help="Optional extractor config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    declared_addbacks_total: float | None = typer.Option(
        None,
        "--declared-addbacks-total",
        help="Total parsed from email/PDF instructions, used for reconciliation.",
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

    extractor_config = AddbacksGLConfig.from_json_file(config)
    if declared_addbacks_total is not None:
        extractor_config.declared_totals.append(
            DeclaredTotal(group_name="addbacks", amount=declared_addbacks_total, tolerance=1.0)
        )
    if no_calculate_formulas:
        extractor_config.calculate_formulas = False

    result = extract_addbacks_gl(path, extractor_config)
    ledger = result.ledger
    groups = result.groups
    summaries = result.summaries

    console.print(f"File: {path}")
    console.print(f"Sheet: {result.sheet_name}")
    console.print(f"Header row: {result.header_row}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Reviewed GL Groups")
    table.add_column("Group")
    table.add_column("Rows", justify="right")
    table.add_column("Amount", justify="right")
    for row in result.group_summaries:
        table.add_row(
            row["group_name"],
            str(row["row_count"]),
            f"{row['amount_total']:,.2f}",
        )
    console.print(table)

    if result.reconciliations:
        recon = Table(title="Declared Total Reconciliation")
        recon.add_column("Group")
        recon.add_column("Declared", justify="right")
        recon.add_column("Extracted", justify="right")
        recon.add_column("Difference", justify="right")
        recon.add_column("Status")
        for row in result.reconciliations:
            recon.add_row(
                row["group_name"],
                f"{row['declared_total']:,.2f}",
                f"{row['extracted_total']:,.2f}",
                f"{row['difference']:,.2f}",
                row["status"],
            )
        console.print(recon)

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        write_frame(ledger, output_dir / f"addbacks_gl_ledger.{format}", format)
        write_frame(groups, output_dir / f"addbacks_gl_groups.{format}", format)
        write_frame(summaries, output_dir / f"addbacks_gl_summaries.{format}", format)
        console.print(f"Wrote outputs to {output_dir}")


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
