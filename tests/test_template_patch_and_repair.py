from pathlib import Path

from openpyxl import load_workbook

from hgf_pnl.pipeline.repair import repair_status, repair_xlsx_bytes
from hgf_pnl.pipeline.template_patch import patch_consolidated_template


TEMPLATE = Path("sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx")


def test_repairs_trailing_and_truncated_xlsx_eocd_bytes() -> None:
    valid = b"prefix" + b"PK\x05\x06" + (b"\x00" * 18)
    with_trailing_junk = valid + b"\x00\x00junk"
    truncated = valid[:-3]

    repaired_trailing = repair_xlsx_bytes(with_trailing_junk)
    repaired_truncated = repair_xlsx_bytes(truncated)

    assert repaired_trailing == valid
    assert repair_status(with_trailing_junk, repaired_trailing) == "truncated"
    assert repaired_truncated == valid
    assert repair_status(truncated, repaired_truncated) == "padded"


def test_patches_known_consolidated_template_cells(tmp_path: Path) -> None:
    output = tmp_path / "patched.xlsx"

    result = patch_consolidated_template(TEMPLATE, output)

    assert output.exists()
    assert result.full_sheet_name == "MARCH 2026 FULL "
    workbook = load_workbook(output, data_only=False)
    try:
        full = workbook["MARCH 2026 FULL "]
        master = workbook["RAW DATA_Master File"]

        assert full["AA31"].value == 0
        assert full["AW52"].value == "=+'RAW DATA_Master File'!B23"
        assert full["BH52"].value == 0
        assert full["AL55"].value == "=0.25*EB55"
        assert full["BH55"].value == "=0.2*EB55"
        assert master["B100"].value == (
            "=+'RAW DATA_COGS & Freight'!G5+'RAW DATA_COGS & Freight'!K5"
        )
    finally:
        workbook.close()
