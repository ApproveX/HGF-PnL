from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import json
import shutil
from typing import Any

import polars as pl

from hgf_pnl.extractors.addbacks_gl import AddbacksGLConfig, DeclaredTotal, extract_addbacks_gl
from hgf_pnl.extractors.br_info import extract_br_info
from hgf_pnl.extractors.chargeback_pdf import extract_chargeback_pdf
from hgf_pnl.extractors.division_cogs import extract_division_cogs
from hgf_pnl.extractors.monthly_revenue import extract_monthly_revenue
from hgf_pnl.extractors.payroll_journal import extract_payroll_journal
from hgf_pnl.extractors.pl_by_dept import extract_pl_by_dept
from hgf_pnl.extractors.th_revenue import extract_th_revenue
from hgf_pnl.pipeline.close_values import build_consolidated_values, flatten_values, read_frame, write_values_json
from hgf_pnl.pipeline.discovery import discover_package, discovery_summary
from hgf_pnl.pipeline.manifest import (
    ManifestEvent,
    ManifestInput,
    RunManifest,
    load_manifest,
    manifest_from_discovery,
)
from hgf_pnl.pipeline.repair import iter_xlsx_files, repair_xlsx_file
from hgf_pnl.pipeline.template_patch import TemplatePatchResult, patch_consolidated_template
from hgf_pnl.writers.consolidated_pnl import (
    default_consolidated_pnl_writer_config,
    write_consolidated_pnl,
)


CORE_EXTRACTORS = [
    "pl_by_dept",
    "th_revenue",
    "payroll_journal",
    "br_info",
    "monthly_revenue",
    "division_cogs",
    "chargeback_pdf",
]


@dataclass
class PhaseResult:
    phase: str
    status: str
    report_path: Path
    manifest_path: Path | None = None
    artifacts: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    next_phase: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "phase": self.phase,
            "status": self.status,
            "report_path": str(self.report_path),
            "manifest_path": str(self.manifest_path) if self.manifest_path else None,
            "artifacts": [str(path) for path in self.artifacts],
            "warnings": self.warnings,
            "next_phase": self.next_phase,
        }


def run_inputs_phase(
    input_root: Path,
    run_dir: Path,
    *,
    year: int | None = None,
    month: int | None = None,
    period_label: str | None = None,
    inspect_workbooks: bool = True,
    stage_repaired_copy: bool = False,
) -> PhaseResult:
    run_dir.mkdir(parents=True, exist_ok=True)
    discovery_root = input_root
    repair_rows: list[dict[str, Any]] = []
    warnings: list[str] = []

    if stage_repaired_copy:
        staged_root = run_dir / "staged-inputs"
        stage_package_copy(input_root, staged_root)
        for workbook in iter_xlsx_files(staged_root):
            result = repair_xlsx_file(workbook, in_place=True, backup=False)
            repair_rows.append(result.to_dict())
        discovery_root = staged_root

    discovery = discover_package(discovery_root, inspect_workbooks=inspect_workbooks)
    manifest = manifest_from_discovery(
        discovery,
        period_label=period_label,
        year=year,
        month=month,
    )
    missing = missing_core_extractors(manifest)
    duplicates = duplicate_components(manifest)
    if missing:
        warnings.append(f"Missing core extractor inputs: {', '.join(missing)}")
    for component, count in duplicates.items():
        warnings.append(f"Multiple selected candidates for {component}: {count}")
    manifest.warnings.extend(warnings)
    manifest.status = "needs_review" if warnings else "discovered"

    discovery_path = run_dir / "discovery.json"
    manifest_path = run_dir / "run_manifest.json"
    discovery.to_json_file(discovery_path)
    manifest.to_json_file(manifest_path)

    report_path = run_dir / "01_inputs_review.md"
    write_report(
        report_path,
        "Inputs Review",
        [
            ("Run", report_kv({"input_root": input_root, "discovery_root": discovery_root, "run_dir": run_dir})),
            ("Discovery Summary", fenced_json(discovery_summary(discovery))),
            ("Repair Summary", report_repair_rows(repair_rows)),
            ("Selected Inputs", report_selected_inputs(manifest)),
            ("Review Warnings", report_list(warnings)),
            ("Approval Gate", "Review file classifications, duplicates, and missing inputs before running `extract`."),
        ],
    )

    return PhaseResult(
        phase="inputs",
        status=manifest.status,
        report_path=report_path,
        manifest_path=manifest_path,
        artifacts=[discovery_path, manifest_path],
        warnings=warnings,
        next_phase="extract",
    )


def run_extract_phase(
    manifest_path: Path,
    run_dir: Path | None = None,
    *,
    declared_addbacks_total: float | None = None,
) -> PhaseResult:
    manifest = load_manifest(manifest_path)
    run_dir = run_dir or manifest_path.parent
    run_dir.mkdir(parents=True, exist_ok=True)
    artifacts: list[Path] = []
    warnings: list[str] = []
    summaries: list[dict[str, Any]] = []

    for extractor in CORE_EXTRACTORS + ["addbacks_gl"]:
        item = select_component_input(manifest, extractor=extractor)
        if item is None:
            if extractor in CORE_EXTRACTORS:
                warnings.append(f"Missing selected input for {extractor}")
            continue
        try:
            output_paths, summary = run_one_extractor(
                extractor,
                Path(item.path),
                run_dir,
                declared_addbacks_total=declared_addbacks_total,
            )
            item.output_paths = [str(path) for path in output_paths]
            artifacts.extend(output_paths)
            summaries.append(summary)
        except Exception as exc:
            warning = f"{extractor}: {type(exc).__name__}: {exc}"
            warnings.append(warning)
            item.warnings.append(warning)

    manifest.status = "needs_review" if warnings else "extracted"
    manifest.events.append(
        ManifestEvent(
            event_type="extract",
            status="completed" if not warnings else "needs_review",
            output_paths=[str(path) for path in artifacts],
            summary={"extractor_count": len(summaries)},
            warnings=warnings,
        )
    )
    manifest.to_json_file(manifest_path)

    report_path = run_dir / "02_extraction_review.md"
    write_report(
        report_path,
        "Extraction Review",
        [
            ("Extractor Outputs", report_extractor_summaries(summaries)),
            ("Warnings", report_list(warnings)),
            (
                "Approval Gate",
                "Review extractor totals, BR Info rows, chargeback totals, addbacks, and warnings before running `values`.",
            ),
        ],
    )
    return PhaseResult(
        phase="extract",
        status=manifest.status,
        report_path=report_path,
        manifest_path=manifest_path,
        artifacts=artifacts,
        warnings=warnings,
        next_phase="values",
    )


def run_values_phase(manifest_path: Path, run_dir: Path | None = None) -> PhaseResult:
    manifest = load_manifest(manifest_path)
    run_dir = run_dir or manifest_path.parent
    artifacts = expected_extractor_artifacts(run_dir)
    result = build_consolidated_values(
        pl_by_dept=read_frame_if_exists(artifacts["pl_by_dept"]),
        br_info=read_frame_if_exists(artifacts["br_info"]),
        monthly_revenue_summary=read_frame_if_exists(artifacts["monthly_revenue_summary"]),
        monthly_revenue_sales=read_frame_if_exists(artifacts["monthly_revenue_sales"]),
        monthly_revenue_refunds=read_frame_if_exists(artifacts["monthly_revenue_refunds"]),
        division_cogs_matrix=read_frame_if_exists(artifacts["division_cogs_matrix"]),
        th_revenue_summary=read_frame_if_exists(artifacts["th_revenue_summary"]),
        payroll_employees=read_frame_if_exists(artifacts["payroll_employees"]),
        payroll_allocation_summaries=read_frame_if_exists(artifacts["payroll_allocation_summaries"]),
        payroll_distribution=read_frame_if_exists(artifacts["payroll_distribution"]),
        chargeback_customer_detail=read_frame_if_exists(artifacts["chargeback_customer_detail"]),
        year=manifest.year,
        month_num=manifest.month,
    )
    values_path = run_dir / "consolidated_values.json"
    write_values_json(result, values_path)

    flat = flatten_values(result.values)
    manifest.values_path = str(values_path)
    manifest.status = "needs_review" if result.warnings else "reviewed"
    manifest.events.append(
        ManifestEvent(
            event_type="values",
            status="completed" if not result.warnings else "needs_review",
            output_paths=[str(values_path)],
            summary={"populated_keys": len(flat)},
            warnings=result.warnings,
        )
    )
    manifest.to_json_file(manifest_path)

    report_path = run_dir / "03_values_review.md"
    write_report(
        report_path,
        "Values Review",
        [
            ("Summary", report_kv({"populated_keys": len(flat), "values_path": values_path})),
            ("Priority Values", report_priority_values(flat)),
            ("Warnings", report_list(result.warnings)),
            (
                "Approval Gate",
                "Review mapped values, replacement BR Info rows, defaults, and warnings before running `core`.",
            ),
        ],
    )
    return PhaseResult(
        phase="values",
        status=manifest.status,
        report_path=report_path,
        manifest_path=manifest_path,
        artifacts=[values_path],
        warnings=result.warnings,
        next_phase="core",
    )


def run_core_phase(
    manifest_path: Path,
    run_dir: Path | None = None,
    *,
    output_workbook: Path | None = None,
    unhide_all_sheets: bool = False,
    full_report_sheet: str | None = None,
) -> PhaseResult:
    manifest = load_manifest(manifest_path)
    run_dir = run_dir or manifest_path.parent
    values_path = Path(manifest.values_path) if manifest.values_path else run_dir / "consolidated_values.json"
    if not values_path.exists():
        raise FileNotFoundError(f"Values JSON not found: {values_path}")

    template_item = select_template_input(manifest)
    if template_item is None:
        raise ValueError("No selected consolidated template/prior-output workbook found in manifest")

    patched_template = run_dir / "patched_template.xlsx"
    patch_result = patch_consolidated_template(
        Path(template_item.path),
        patched_template,
        unhide_all_sheets=unhide_all_sheets,
    )
    writer_config_path = run_dir / "consolidated_writer_config.json"
    writer_config = default_consolidated_pnl_writer_config()
    if full_report_sheet is not None:
        writer_config.full_report_sheet_name = full_report_sheet
    writer_config_path.write_text(writer_config.model_dump_json(indent=2) + "\n", encoding="utf-8")

    if output_workbook is None:
        output_workbook = run_dir / default_output_workbook_name(manifest)
    values = json.loads(values_path.read_text(encoding="utf-8"))
    write_result = write_consolidated_pnl(
        patched_template,
        output_workbook,
        values=values,
        config=writer_config,
    )

    warnings = list(write_result.warnings)
    validation_failures = [
        row for row in write_result.validation_results if row.get("status") != "ok"
    ]
    if validation_failures:
        warnings.append(f"Writer validation failures: {len(validation_failures)}")

    write_metadata_path = run_dir / "core_write_metadata.json"
    write_metadata_path.write_text(json.dumps(write_result.to_dict(), indent=2) + "\n", encoding="utf-8")
    patch_metadata_path = run_dir / "template_patch_metadata.json"
    patch_metadata_path.write_text(json.dumps(patch_result.to_dict(), indent=2) + "\n", encoding="utf-8")

    manifest.writer_config_path = str(writer_config_path)
    manifest.output_workbook_path = str(output_workbook)
    manifest.status = "needs_review" if warnings else "written"
    manifest.events.append(
        ManifestEvent(
            event_type="core_workbook",
            status="completed" if not warnings else "needs_review",
            output_paths=[str(output_workbook), str(write_metadata_path), str(patch_metadata_path)],
            summary={
                "written_cells": len(write_result.written_cells),
                "skipped_cells": len(write_result.skipped_cells),
                "validations": len(write_result.validation_results),
                "validation_failures": len(validation_failures),
            },
            warnings=warnings,
        )
    )
    manifest.to_json_file(manifest_path)

    report_path = run_dir / "04_core_workbook_review.md"
    write_report(
        report_path,
        "Core Workbook Review",
        [
            ("Template Patches", report_template_patches(patch_result)),
            (
                "Writer Summary",
                report_kv(
                    {
                        "output_workbook": output_workbook,
                        "written_cells": len(write_result.written_cells),
                        "skipped_cells": len(write_result.skipped_cells),
                        "validation_failures": len(validation_failures),
                        "metadata": write_metadata_path,
                    }
                ),
            ),
            ("Warnings", report_list(warnings)),
            (
                "Approval Gate",
                "Recalculate the workbook and review FULL/SIMPLE/raw outputs before running secondary-tab work.",
            ),
        ],
    )
    return PhaseResult(
        phase="core",
        status=manifest.status,
        report_path=report_path,
        manifest_path=manifest_path,
        artifacts=[output_workbook, write_metadata_path, patch_metadata_path],
        warnings=warnings,
        next_phase="secondary-tabs",
    )


def stage_package_copy(input_root: Path, staged_root: Path) -> None:
    if staged_root.exists():
        shutil.rmtree(staged_root)
    shutil.copytree(
        input_root,
        staged_root,
        ignore=lambda _dir, names: [
            name
            for name in names
            if name.endswith(":Zone.Identifier") or name.startswith("~$")
        ],
    )


def missing_core_extractors(manifest: RunManifest) -> list[str]:
    return [
        extractor
        for extractor in CORE_EXTRACTORS
        if select_component_input(manifest, extractor=extractor) is None
    ]


def duplicate_components(manifest: RunManifest) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in manifest.selected_inputs():
        component = item.extractor or (f"writer:{item.writer}" if item.writer else None)
        if component:
            counts[component] = counts.get(component, 0) + 1
    return {component: count for component, count in counts.items() if count > 1}


def select_component_input(manifest: RunManifest, *, extractor: str) -> ManifestInput | None:
    candidates = [
        item for item in manifest.selected_inputs() if item.extractor == extractor
    ]
    if not candidates:
        return None
    return sorted(candidates, key=lambda item: (item.confidence, item.role == "source_input"), reverse=True)[0]


def select_template_input(manifest: RunManifest) -> ManifestInput | None:
    candidates = [
        item
        for item in manifest.selected_inputs()
        if item.writer == "consolidated_pnl"
        and item.document_type in {"template_workbook", "deliverable_workbook", "workbook"}
    ]
    if not candidates:
        return None
    return sorted(
        candidates,
        key=lambda item: (
            item.role == "template",
            item.document_type == "template_workbook",
            item.confidence,
        ),
        reverse=True,
    )[0]


def run_one_extractor(
    extractor: str,
    path: Path,
    run_dir: Path,
    *,
    declared_addbacks_total: float | None,
) -> tuple[list[Path], dict[str, Any]]:
    if extractor == "pl_by_dept":
        result = extract_pl_by_dept(path)
        output = run_dir / "pl_by_dept.csv"
        write_frame(result.to_polars(), output)
        return [output], frame_summary(extractor, path, result.to_polars(), "amount", result.warnings)

    if extractor == "th_revenue":
        result = extract_th_revenue(path)
        out = run_dir / "th_revenue"
        paths = [
            write_frame(result.account_summary, out / "th_revenue_summary.csv"),
            write_frame(result.po_details, out / "th_revenue_po_details.csv"),
            write_frame(result.usa_stock, out / "th_revenue_usa_stock.csv"),
            write_frame(result.all_rows(), out / "th_revenue_all_rows.csv"),
        ]
        return paths, frame_summary(extractor, path, result.account_summary, "revenue", result.warnings)

    if extractor == "payroll_journal":
        result = extract_payroll_journal(path)
        out = run_dir / "payroll_journal"
        paths = [
            write_frame(result.employees, out / "payroll_employees.csv"),
            write_frame(result.allocations, out / "payroll_allocations.csv"),
            write_frame(result.allocation_summaries, out / "payroll_allocation_summaries.csv"),
            write_frame(result.distribution, out / "payroll_distribution.csv"),
        ]
        return paths, frame_summary(extractor, path, result.employees, "gross_pay", result.warnings)

    if extractor == "br_info":
        result = extract_br_info(path)
        output = run_dir / "br_info.csv"
        write_frame(result.overrides, output)
        return [output], frame_summary(extractor, path, result.overrides, "value", result.warnings)

    if extractor == "monthly_revenue":
        result = extract_monthly_revenue(path)
        out = run_dir / "monthly_revenue"
        paths = [
            write_frame(result.summary, out / "monthly_revenue_summary.csv"),
            write_frame(result.sales, out / "monthly_revenue_sales.csv"),
            write_frame(result.refunds, out / "monthly_revenue_refunds.csv"),
            write_frame(result.coupons, out / "monthly_revenue_coupons.csv"),
        ]
        return paths, frame_summary(extractor, path, result.sales, "net_sales", result.warnings)

    if extractor == "division_cogs":
        result = extract_division_cogs(path)
        out = run_dir / "division_cogs"
        paths = [
            write_frame(result.matrix, out / "division_cogs_matrix.csv"),
            write_frame(result.partner_details, out / "division_cogs_partner_details.csv"),
        ]
        return paths, frame_summary(extractor, path, result.matrix, "amount", result.warnings)

    if extractor == "chargeback_pdf":
        result = extract_chargeback_pdf(path)
        out = run_dir / "chargeback_pdf"
        paths = [
            write_frame(result.monthly_summary, out / "chargeback_monthly_summary.csv"),
            write_frame(result.customer_detail, out / "chargeback_customer_detail.csv"),
            write_frame(result.reconciliation, out / "chargeback_reconciliation.csv"),
        ]
        return paths, frame_summary(extractor, path, result.monthly_summary, "amount", result.warnings)

    if extractor == "addbacks_gl":
        config = AddbacksGLConfig()
        if declared_addbacks_total is not None:
            config.declared_totals.append(
                DeclaredTotal(group_name="addbacks", amount=declared_addbacks_total, tolerance=1.0)
            )
        result = extract_addbacks_gl(path, config)
        out = run_dir / "addbacks_gl"
        paths = [
            write_frame(result.ledger, out / "addbacks_gl_ledger.csv"),
            write_frame(result.groups, out / "addbacks_gl_groups.csv"),
            write_frame(result.summaries, out / "addbacks_gl_summaries.csv"),
        ]
        return paths, frame_summary(extractor, path, result.groups, "amount", result.warnings)

    raise ValueError(f"Unsupported extractor: {extractor}")


def expected_extractor_artifacts(run_dir: Path) -> dict[str, Path]:
    return {
        "pl_by_dept": run_dir / "pl_by_dept.csv",
        "br_info": run_dir / "br_info.csv",
        "monthly_revenue_summary": run_dir / "monthly_revenue" / "monthly_revenue_summary.csv",
        "monthly_revenue_sales": run_dir / "monthly_revenue" / "monthly_revenue_sales.csv",
        "monthly_revenue_refunds": run_dir / "monthly_revenue" / "monthly_revenue_refunds.csv",
        "division_cogs_matrix": run_dir / "division_cogs" / "division_cogs_matrix.csv",
        "th_revenue_summary": run_dir / "th_revenue" / "th_revenue_summary.csv",
        "payroll_employees": run_dir / "payroll_journal" / "payroll_employees.csv",
        "payroll_allocation_summaries": run_dir
        / "payroll_journal"
        / "payroll_allocation_summaries.csv",
        "payroll_distribution": run_dir / "payroll_journal" / "payroll_distribution.csv",
        "chargeback_customer_detail": run_dir / "chargeback_pdf" / "chargeback_customer_detail.csv",
    }


def read_frame_if_exists(path: Path) -> pl.DataFrame | None:
    if not path.exists():
        return None
    return read_frame(path)


def write_frame(df: pl.DataFrame, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.write_csv(path)
    return path


def frame_summary(
    extractor: str,
    path: Path,
    df: pl.DataFrame,
    amount_column: str,
    warnings: list[str],
) -> dict[str, Any]:
    amount = None
    if not df.is_empty() and amount_column in df.columns:
        amount = float(df.select(pl.col(amount_column).fill_null(0).sum()).item() or 0.0)
    return {
        "extractor": extractor,
        "source": str(path),
        "rows": df.height,
        "amount_column": amount_column,
        "amount_sum": amount,
        "warnings": warnings,
    }


def default_output_workbook_name(manifest: RunManifest) -> str:
    period = (manifest.period_label or "generated").replace(" ", "_").upper()
    return f"HGF_CONSOLIDATED_{period}_GENERATED.xlsx"


def write_report(path: Path, title: str, sections: list[tuple[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"# {title}", ""]
    for heading, body in sections:
        lines.extend([f"## {heading}", "", body.strip() or "_None._", ""])
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def report_kv(values: dict[str, Any]) -> str:
    return "\n".join(f"- **{key}:** `{value}`" for key, value in values.items())


def report_list(values: list[str]) -> str:
    if not values:
        return "_None._"
    return "\n".join(f"- {value}" for value in values)


def fenced_json(value: Any) -> str:
    return "```json\n" + json.dumps(value, indent=2, default=str) + "\n```"


def report_repair_rows(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "_No staging/repair pass was run._"
    lines = ["| Status | Source | Output | Bytes |", "|---|---|---|---:|"]
    for row in rows:
        lines.append(
            f"| {row['status']} | `{row['source_path']}` | `{row['output_path']}` | "
            f"{row['original_size']} -> {row['repaired_size']} |"
        )
    return "\n".join(lines)


def report_selected_inputs(manifest: RunManifest) -> str:
    lines = ["| Component | Role | Confidence | Path |", "|---|---|---:|---|"]
    for item in sorted(manifest.selected_inputs(), key=lambda entry: entry.relative_path):
        component = item.extractor or (f"writer:{item.writer}" if item.writer else item.document_type)
        lines.append(f"| {component} | {item.role} | {item.confidence:.2f} | `{item.relative_path}` |")
    return "\n".join(lines)


def report_extractor_summaries(summaries: list[dict[str, Any]]) -> str:
    if not summaries:
        return "_No extractors completed._"
    lines = ["| Extractor | Rows | Amount Sum | Source |", "|---|---:|---:|---|"]
    for row in summaries:
        amount = row["amount_sum"]
        amount_text = "" if amount is None else f"{amount:,.2f}"
        lines.append(
            f"| {row['extractor']} | {row['rows']} | {amount_text} | `{row['source']}` |"
        )
    return "\n".join(lines)


def report_priority_values(flat: dict[str, Any]) -> str:
    keys = [
        "raw_master.sales.online",
        "raw_master.sales.dtc",
        "raw_master.returns.dtc",
        "raw_master.returns.trend_house",
        "raw_master.gl.bank_fees",
        "raw_master.gl.bank_fees_adjustment",
        "raw_master.gl.merchant_account_fees",
        "raw_master.gl.merchant_account_fees_adjustment",
        "raw_cogs.current_month.cogs.online_usa",
        "raw_cogs.current_month.cogs.online",
        "raw_cogs.tariffs",
        "raw_payroll.production",
        "raw_payroll.lital_allocation.corp",
    ]
    lines = ["| Key | Value |", "|---|---:|"]
    for key in keys:
        value = flat.get(key, "")
        lines.append(f"| `{key}` | {value} |")
    return "\n".join(lines)


def report_template_patches(result: TemplatePatchResult) -> str:
    lines = ["| Sheet | Cell | Old | New | Note |", "|---|---|---|---|---|"]
    for patch in result.patches:
        lines.append(
            f"| {patch.sheet_name} | {patch.cell} | `{patch.old_value}` | "
            f"`{patch.new_value}` | {patch.note} |"
        )
    return "\n".join(lines)
