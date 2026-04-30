from pathlib import Path

from hgf_pnl.pipeline.discovery import classify_file, discover_package
from hgf_pnl.pipeline.manifest import manifest_from_discovery, manifest_summary


def touch(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"sample")


def test_discovers_and_classifies_known_package_files(tmp_path: Path) -> None:
    touch(tmp_path / "DATA" / "Profit and Loss By Dept.xlsx")
    touch(tmp_path / "DATA" / "- OG _ Chargeback Report - 03. March 2026.pdf")
    touch(tmp_path / "DATA" / "DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx")
    touch(tmp_path / "DATA" / "INTERNAL - Division COGS 2019 - Current (26).xlsx")
    touch(tmp_path / "DATA" / "TH March 2026 Revenue Report.xlsx")
    touch(tmp_path / "DATA" / "March Addbacks_$23,195.pdf")
    touch(tmp_path / "HGF GL_March_Sent April 13_DONE.xlsx")
    touch(tmp_path / "HGF GL_March_Sent April 13.xlsx")
    touch(tmp_path / "HGF CONSOLIDATED_MARCH 2026 Template.xlsx")
    touch(tmp_path / "Profit and Loss By Dept.xlsx:Zone.Identifier")

    discovery = discover_package(tmp_path)

    by_name = {file.file_name: file for file in discovery.files}
    assert "Profit and Loss By Dept.xlsx:Zone.Identifier" not in by_name
    assert by_name["Profit and Loss By Dept.xlsx"].extractor == "pl_by_dept"
    assert by_name["- OG _ Chargeback Report - 03. March 2026.pdf"].extractor == "chargeback_pdf"
    assert by_name["DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx"].extractor == "monthly_revenue"
    assert by_name["INTERNAL - Division COGS 2019 - Current (26).xlsx"].extractor == "division_cogs"
    assert by_name["TH March 2026 Revenue Report.xlsx"].extractor == "th_revenue"
    assert by_name["March Addbacks_$23,195.pdf"].role == "instruction"
    assert by_name["HGF GL_March_Sent April 13_DONE.xlsx"].extractor == "addbacks_gl"
    assert by_name["HGF GL_March_Sent April 13.xlsx"].role == "supporting_input"
    assert by_name["HGF CONSOLIDATED_MARCH 2026 Template.xlsx"].writer == "consolidated_pnl"


def test_manifest_from_discovery_records_inputs_and_period(tmp_path: Path) -> None:
    touch(tmp_path / "DATA" / "Payroll Journal_March 2026.xlsx")
    touch(tmp_path / "BR Info.xlsx")
    touch(tmp_path / "unknown.bin")

    discovery = discover_package(tmp_path)
    manifest = manifest_from_discovery(discovery)
    summary = manifest_summary(manifest)

    assert manifest.status == "discovered"
    assert manifest.period_label == "March 2026"
    assert manifest.year == 2026
    assert manifest.month == 3
    assert len(manifest.events) == 1
    assert summary["input_count"] == 3
    assert summary["selected_input_count"] == 2

    payroll = next(item for item in manifest.inputs if item.file_name.startswith("Payroll"))
    assert payroll.extractor == "payroll_journal"
    assert payroll.selected

    unknown = next(item for item in manifest.inputs if item.file_name == "unknown.bin")
    assert unknown.role == "unknown"
    assert not unknown.selected


def test_sheet_name_refinement_can_classify_reviewed_gl_sample() -> None:
    sample = Path("sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx")
    file = classify_file(sample.parent.resolve(), sample.resolve(), {"sheet_names": ["NEW MONTH"]})

    assert file.extractor == "addbacks_gl"
    assert file.confidence >= 0.95
    assert "workbook has reviewed GL NEW MONTH sheet" in file.reasons
