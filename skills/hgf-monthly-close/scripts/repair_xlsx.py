from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.pipeline.repair import iter_xlsx_files, repair_xlsx_file


app = typer.Typer(help="Repair common XLSX zip-footer corruption in staged close workbooks.")
console = Console()


@app.command()
def repair(
    path: Path = typer.Argument(..., help="XLSX file or directory to scan."),
    output_dir: Path | None = typer.Option(
        None,
        "--output-dir",
        "-o",
        help="Write repaired/copied files under this directory, preserving relative paths.",
    ),
    in_place: bool = typer.Option(
        False,
        "--in-place",
        help="Modify files in place. Use only on staged copies unless the user explicitly approved it.",
    ),
    backup: bool = typer.Option(
        True,
        "--backup/--no-backup",
        help="Create .bak files for in-place repairs.",
    ),
) -> None:
    if output_dir is None and not in_place:
        raise typer.BadParameter("--output-dir is required unless --in-place is set")
    if output_dir is not None and in_place:
        raise typer.BadParameter("Use either --output-dir or --in-place, not both")

    files = iter_xlsx_files(path)
    if not files:
        console.print("[yellow]No XLSX files found.[/yellow]")
        raise typer.Exit()

    root = path if path.is_dir() else path.parent
    rows = []
    for file in files:
        output_path = None
        if output_dir is not None:
            output_path = output_dir / file.relative_to(root)
        result = repair_xlsx_file(file, output_path, in_place=in_place, backup=backup)
        rows.append(result)

    table = Table(title="XLSX Repair")
    table.add_column("Status")
    table.add_column("Source")
    table.add_column("Output")
    table.add_column("Bytes", justify="right")
    for row in rows:
        table.add_row(
            row.status,
            str(row.source_path),
            str(row.output_path),
            f"{row.original_size:,} -> {row.repaired_size:,}",
        )
    console.print(table)


if __name__ == "__main__":
    app()
