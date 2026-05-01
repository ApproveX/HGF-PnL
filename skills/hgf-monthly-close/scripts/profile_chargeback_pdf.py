from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.extractors.chargeback_pdf import (
    profile_chargeback_pdf,
    write_profile_artifacts,
)


app = typer.Typer(help="Profile chargeback PDFs before configuring extraction.")
console = Console()


@app.command()
def profile(
    path: Path = typer.Argument(..., help="Path to chargeback report PDF."),
    output_dir: Path = typer.Option(
        Path("tmp/chargeback_pdf_profile"),
        "--output-dir",
        "-o",
        help="Directory for raw text, profile, and suggested config.",
    ),
) -> None:
    result = profile_chargeback_pdf(path)
    write_profile_artifacts(result, output_dir)

    console.print(f"File: {path}")
    console.print(f"Pages: {result.page_count}")
    console.print(f"Text lines: {len(result.lines)}")
    console.print(f"Monthly candidates: {len(result.monthly_line_candidates)}")
    console.print(f"Anchor candidates: {len(result.anchor_candidates)}")
    console.print(f"Tables: {len(result.table_summaries)}")

    table = Table(title="Profile Artifacts")
    table.add_column("Artifact")
    table.add_column("Path")
    for name in [
        "chargeback_pdf_profile.md",
        "chargeback_pdf_profile.json",
        "chargeback_pdf_raw_text.txt",
        "chargeback_pdf_suggested_config.json",
    ]:
        table.add_row(name, str(output_dir / name))
    console.print(table)


if __name__ == "__main__":
    app()
