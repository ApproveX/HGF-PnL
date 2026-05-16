from pathlib import Path

from hgf_pnl.pipeline.close_phases import run_inputs_phase
from hgf_pnl.pipeline.manifest import load_manifest


def touch(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"sample")


def test_inputs_phase_writes_manifest_and_review_report(tmp_path: Path) -> None:
    input_root = tmp_path / "inputs"
    run_dir = tmp_path / "run"
    touch(input_root / "DATA" / "Profit and Loss By Dept.xlsx")
    touch(input_root / "DATA" / "- OG _ Chargeback Report - 03. March 2026.pdf")
    touch(input_root / "DATA" / "DTC & WS Monthly Revenue - report (03.01-03.31).xlsx")
    touch(input_root / "DATA" / "INTERNAL - Division COGS 2019 - Current (26).xlsx")
    touch(input_root / "DATA" / "TH March 2026 Revenue Report.xlsx")
    touch(input_root / "PAYROLL" / "Payroll Journal_March 2026.xlsx")
    touch(input_root / "BR Info.xlsx")
    touch(input_root / "HGF CONSOLIDATED_MARCH 2026 Template.xlsx")

    result = run_inputs_phase(
        input_root,
        run_dir,
        inspect_workbooks=False,
    )

    assert result.status == "discovered"
    assert result.report_path.exists()
    assert result.manifest_path is not None
    manifest = load_manifest(result.manifest_path)
    assert manifest.period_label == "March 2026"
    assert len(manifest.selected_inputs()) == 8
    assert "Approval Gate" in result.report_path.read_text(encoding="utf-8")


def test_inputs_phase_flags_missing_core_inputs(tmp_path: Path) -> None:
    input_root = tmp_path / "inputs"
    run_dir = tmp_path / "run"
    touch(input_root / "BR Info.xlsx")

    result = run_inputs_phase(
        input_root,
        run_dir,
        inspect_workbooks=False,
    )

    assert result.status == "needs_review"
    assert any("Missing core extractor inputs" in warning for warning in result.warnings)
    assert "Missing core extractor inputs" in result.report_path.read_text(encoding="utf-8")
