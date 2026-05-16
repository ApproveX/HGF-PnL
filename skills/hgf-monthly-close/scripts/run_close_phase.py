from __future__ import annotations

from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from hgf_pnl.pipeline.close_phases import (
    PhaseResult,
    run_core_phase,
    run_extract_phase,
    run_inputs_phase,
    run_values_phase,
)
from hgf_pnl.pipeline.manifest import load_manifest, manifest_summary


app = typer.Typer(help="Run HGF monthly close phases with explicit approval gates.")
console = Console()


@app.command("inputs")
def inputs_phase(
    input_root: Path = typer.Argument(..., help="Close package input folder."),
    run_dir: Path = typer.Argument(..., help="Run directory for manifests/reports/artifacts."),
    year: int | None = typer.Option(None, "--year", help="Close year."),
    month: int | None = typer.Option(None, "--month", help="Close month number."),
    period_label: str | None = typer.Option(None, "--period-label", help="Display period label."),
    no_inspect_workbooks: bool = typer.Option(
        False,
        "--no-inspect-workbooks",
        help="Skip opening workbooks during discovery.",
    ),
    stage_repaired_copy: bool = typer.Option(
        False,
        "--stage-repaired-copy",
        help="Copy inputs into the run directory and repair XLSX files there before discovery.",
    ),
) -> None:
    result = run_inputs_phase(
        input_root,
        run_dir,
        year=year,
        month=month,
        period_label=period_label,
        inspect_workbooks=not no_inspect_workbooks,
        stage_repaired_copy=stage_repaired_copy,
    )
    print_result(result)


@app.command("extract")
def extract_phase(
    manifest: Path = typer.Argument(..., help="Path to run_manifest.json."),
    run_dir: Path | None = typer.Option(None, "--run-dir", help="Override run artifact directory."),
    declared_addbacks_total: float | None = typer.Option(
        None,
        "--declared-addbacks-total",
        help="Addbacks total parsed from the PDF/email instructions.",
    ),
) -> None:
    result = run_extract_phase(
        manifest,
        run_dir,
        declared_addbacks_total=declared_addbacks_total,
    )
    print_result(result)


@app.command("values")
def values_phase(
    manifest: Path = typer.Argument(..., help="Path to run_manifest.json."),
    run_dir: Path | None = typer.Option(None, "--run-dir", help="Override run artifact directory."),
) -> None:
    result = run_values_phase(manifest, run_dir)
    print_result(result)


@app.command("core")
def core_phase(
    manifest: Path = typer.Argument(..., help="Path to run_manifest.json."),
    run_dir: Path | None = typer.Option(None, "--run-dir", help="Override run artifact directory."),
    output_workbook: Path | None = typer.Option(None, "--output-workbook", help="Output workbook path."),
    full_report_sheet: str | None = typer.Option(
        None,
        "--full-report-sheet",
        help="Exact FULL report sheet name. Defaults to auto-detecting the workbook's FULL sheet.",
    ),
    unhide_all_sheets: bool = typer.Option(
        False,
        "--unhide-all-sheets",
        help="Make all sheets visible in the patched template before writing.",
    ),
) -> None:
    result = run_core_phase(
        manifest,
        run_dir,
        output_workbook=output_workbook,
        full_report_sheet=full_report_sheet,
        unhide_all_sheets=unhide_all_sheets,
    )
    print_result(result)


@app.command("status")
def status(manifest: Path = typer.Argument(..., help="Path to run_manifest.json.")) -> None:
    run_manifest = load_manifest(manifest)
    summary = manifest_summary(run_manifest)
    console.print(f"Manifest: {manifest}")
    console.print(f"Run ID: {summary['run_id']}")
    console.print(f"Status: {summary['status']}")
    console.print(f"Period: {run_manifest.period_label or '(unknown)'}")
    console.print(f"Events: {summary['event_count']}")
    if run_manifest.values_path:
        console.print(f"Values: {run_manifest.values_path}")
    if run_manifest.output_workbook_path:
        console.print(f"Output workbook: {run_manifest.output_workbook_path}")
    if summary["warnings"]:
        console.print("[yellow]Warnings:[/yellow]")
        for warning in summary["warnings"]:
            console.print(f"- {warning}")


def print_result(result: PhaseResult) -> None:
    table = Table(title=f"Phase: {result.phase}")
    table.add_column("Field")
    table.add_column("Value")
    table.add_row("status", result.status)
    table.add_row("report", str(result.report_path))
    table.add_row("manifest", str(result.manifest_path or ""))
    table.add_row("artifacts", str(len(result.artifacts)))
    table.add_row("warnings", str(len(result.warnings)))
    table.add_row("next_phase", result.next_phase or "")
    console.print(table)
    if result.warnings:
        console.print("[yellow]Warnings:[/yellow]")
        for warning in result.warnings:
            console.print(f"- {warning}")
    console.print("Approval gate: review the report before running the next phase.")


if __name__ == "__main__":
    app()
