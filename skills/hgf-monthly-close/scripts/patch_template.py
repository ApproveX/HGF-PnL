from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.pipeline.template_patch import patch_consolidated_template


app = typer.Typer(help="Patch known consolidated-template formulas before writing HGF P&L values.")
console = Console()


@app.command()
def patch(
    source: Path = typer.Argument(..., help="Source consolidated template workbook."),
    output: Path = typer.Argument(..., help="Patched template output path."),
    full_sheet: str | None = typer.Option(None, "--full-sheet", help="Exact FULL report sheet name."),
    unhide_all_sheets: bool = typer.Option(
        False,
        "--unhide-all-sheets",
        help="Set every worksheet visible in the patched copy.",
    ),
) -> None:
    result = patch_consolidated_template(
        source,
        output,
        full_sheet_name=full_sheet,
        unhide_all_sheets=unhide_all_sheets,
    )

    console.print(f"Source: {result.source_path}")
    console.print(f"Output: {result.output_path}")
    console.print(f"FULL sheet: {result.full_sheet_name!r}")

    table = Table(title="Template Patches")
    table.add_column("Sheet")
    table.add_column("Cell")
    table.add_column("Old")
    table.add_column("New")
    table.add_column("Note")
    for patch_row in result.patches:
        table.add_row(
            patch_row.sheet_name,
            patch_row.cell,
            str(patch_row.old_value),
            str(patch_row.new_value),
            patch_row.note,
        )
    console.print(table)


if __name__ == "__main__":
    app()
