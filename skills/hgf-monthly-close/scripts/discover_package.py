from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.pipeline.discovery import discover_package, discovery_summary
from hgf_pnl.pipeline.manifest import manifest_from_discovery, manifest_summary


app = typer.Typer(help="Discover and classify HGF close package files.")
console = Console()


@app.command()
def scan(
    root: Path = typer.Argument(..., help="Close package root directory."),
    discovery_output: Path | None = typer.Option(
        None,
        "--discovery-output",
        help="Write discovery JSON to this path.",
    ),
    manifest_output: Path | None = typer.Option(
        None,
        "--manifest-output",
        help="Write initial run manifest JSON to this path.",
    ),
    inspect_workbooks: bool = typer.Option(
        False,
        "--inspect-workbooks",
        help="Open workbooks to inspect sheet names for classification refinement.",
    ),
    include_zone_identifier: bool = typer.Option(
        False,
        "--include-zone-identifier",
        help="Include browser Zone.Identifier sidecar files.",
    ),
    include_temp_files: bool = typer.Option(
        False,
        "--include-temp-files",
        help="Include temporary Office lock files.",
    ),
) -> None:
    discovery = discover_package(
        root,
        inspect_workbooks=inspect_workbooks,
        include_zone_identifier=include_zone_identifier,
        include_temp_files=include_temp_files,
    )
    summary = discovery_summary(discovery)
    console.print(f"Root: {summary['root_path']}")
    console.print(f"Files: {summary['file_count']}")
    for warning in discovery.warnings:
        console.print(f"[yellow]Warning:[/yellow] {warning}")

    table = Table(title="Classification Summary")
    table.add_column("Role")
    table.add_column("Count", justify="right")
    for role, count in summary["by_role"].items():
        table.add_row(role, str(count))
    console.print(table)

    component_table = Table(title="Extractor / Writer Matches")
    component_table.add_column("Component")
    component_table.add_column("Count", justify="right")
    for component, count in summary["by_extractor"].items():
        component_table.add_row(component, str(count))
    console.print(component_table)

    if discovery_output is not None:
        discovery.to_json_file(discovery_output)
        console.print(f"Wrote discovery JSON to {discovery_output}")

    if manifest_output is not None:
        manifest = manifest_from_discovery(discovery)
        manifest.to_json_file(manifest_output)
        console.print(f"Wrote run manifest to {manifest_output}")
        console.print(manifest_summary(manifest))


if __name__ == "__main__":
    app()
