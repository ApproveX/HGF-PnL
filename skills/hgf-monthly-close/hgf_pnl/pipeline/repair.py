from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import struct


EOCD_SIGNATURE = b"PK\x05\x06"
EOCD_MIN_SIZE = 22


@dataclass(frozen=True)
class XLSXRepairResult:
    source_path: Path
    output_path: Path
    status: str
    original_size: int
    repaired_size: int
    backup_path: Path | None = None

    def to_dict(self) -> dict[str, object]:
        return {
            "source_path": str(self.source_path),
            "output_path": str(self.output_path),
            "status": self.status,
            "original_size": self.original_size,
            "repaired_size": self.repaired_size,
            "backup_path": str(self.backup_path) if self.backup_path else None,
        }


def repair_xlsx_file(
    source_path: Path,
    output_path: Path | None = None,
    *,
    in_place: bool = False,
    backup: bool = False,
) -> XLSXRepairResult:
    """Repair common XLSX zip-footer corruption.

    The default is immutable-source friendly: when ``output_path`` is supplied,
    the repaired or copied workbook is written there and the source is untouched.
    Use ``in_place=True`` only for a staged copy or when the operator explicitly
    wants to modify the source file.
    """

    if output_path is not None and in_place:
        raise ValueError("Use either output_path or in_place, not both")
    if output_path is None and not in_place:
        raise ValueError("output_path is required unless in_place=True")

    source_path = source_path.expanduser().resolve()
    target_path = source_path if in_place else output_path.expanduser().resolve()  # type: ignore[union-attr]
    data = source_path.read_bytes()
    original_size = len(data)
    repaired = repair_xlsx_bytes(data)
    status = repair_status(data, repaired)

    backup_path: Path | None = None
    if in_place and backup and status != "ok":
        backup_path = source_path.with_name(f"{source_path.name}.bak")
        shutil.copy2(source_path, backup_path)

    target_path.parent.mkdir(parents=True, exist_ok=True)
    if target_path != source_path or status != "ok":
        target_path.write_bytes(repaired)

    return XLSXRepairResult(
        source_path=source_path,
        output_path=target_path,
        status=status,
        original_size=original_size,
        repaired_size=len(repaired),
        backup_path=backup_path,
    )


def repair_xlsx_bytes(data: bytes) -> bytes:
    pos = data.rfind(EOCD_SIGNATURE)
    if pos < 0:
        raise ValueError("No XLSX End-of-Central-Directory record found")

    available = len(data) - pos
    if available < EOCD_MIN_SIZE:
        return data + (b"\x00" * (EOCD_MIN_SIZE - available))

    eocd = data[pos : pos + EOCD_MIN_SIZE]
    try:
        *_, comment_length = struct.unpack("<IHHHHIIH", eocd)
    except struct.error as exc:
        raise ValueError("Malformed XLSX End-of-Central-Directory record") from exc

    end = pos + EOCD_MIN_SIZE + comment_length
    if end < len(data):
        return data[:end]
    return data


def repair_status(original: bytes, repaired: bytes) -> str:
    if len(repaired) > len(original):
        return "padded"
    if len(repaired) < len(original):
        return "truncated"
    return "ok"


def iter_xlsx_files(path: Path) -> list[Path]:
    path = path.expanduser()
    if path.is_file():
        return [path]
    return sorted(
        file
        for file in path.rglob("*.xlsx")
        if file.is_file() and not file.name.endswith(".bak")
    )
