from __future__ import annotations

from pathlib import Path

import polars as pl
import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.writers.consolidated_pnl import (
    ConsolidatedPNLWriterConfig,
    load_values_json,
    write_consolidated_pnl,
    write_default_config,
    write_example_values_from_workbook,
)


app = typer.Typer(help="Write HGF consolidated P&L workbooks from validated raw values.")
console = Console()


@app.command()
def write(
    template: Path = typer.Argument(..., help="Path to consolidated P&L template workbook."),
    output: Path = typer.Argument(..., help="Path for the generated workbook."),
    values: Path | None = typer.Option(None, "--values", "-v", help="Input values JSON."),
    config: Path | None = typer.Option(None, "--config", "-c", help="Optional writer config JSON."),
    init_config: Path | None = typer.Option(
        None,
        "--init-config",
        help="Write a default config JSON to this path and exit.",
    ),
    example_values_from: Path | None = typer.Option(
        None,
        "--example-values-from",
        help="Read configured cell values from a completed workbook into --values path and exit.",
    ),
) -> None:
    if init_config is not None:
        write_default_config(init_config)
        console.print(f"Wrote default config to {init_config}")
        raise typer.Exit()

    writer_config = ConsolidatedPNLWriterConfig.from_json_file(config)

    if example_values_from is not None:
        if values is None:
            raise typer.BadParameter("--values is required with --example-values-from")
        extracted_values = write_example_values_from_workbook(
            example_values_from,
            values,
            writer_config,
        )
        console.print(f"Wrote example values to {values}")
        console.print(f"Top-level keys: {', '.join(sorted(extracted_values))}")
        raise typer.Exit()

    result = write_consolidated_pnl(
        template_path=template,
        output_path=output,
        values=load_values_json(values),
        config=writer_config,
    )

    console.print(f"Template: {template}")
    console.print(f"Output: {output}")
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Consolidated P&L Writer")
    table.add_column("Metric")
    table.add_column("Count", justify="right")
    table.add_row("written_cells", str(len(result.written_cells)))
    table.add_row("skipped_cells", str(len(result.skipped_cells)))
    table.add_row("validations", str(len(result.validation_results)))
    table.add_row("validation_failures", str(count_validation_failures(result.validation_results)))
    console.print(table)

    if result.validation_results:
        console.print(pl.DataFrame(result.validation_results).select(["name", "status", "cell"]))


def count_validation_failures(rows: list[dict[str, object]]) -> int:
    return sum(1 for row in rows if row.get("status") != "ok")


if __name__ == "__main__":
    app()
