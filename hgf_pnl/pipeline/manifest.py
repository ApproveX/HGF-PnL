from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
from typing import Any, Literal
from uuid import uuid4

from pydantic import BaseModel, Field

from hgf_pnl.pipeline.discovery import PackageDiscovery, PackageFile


ManifestStatus = Literal[
    "discovered",
    "configured",
    "extracted",
    "reviewed",
    "written",
    "validated",
    "needs_review",
    "failed",
]


class ManifestInput(BaseModel):
    input_id: str
    path: str
    relative_path: str
    file_name: str
    role: str
    document_type: str
    extractor: str | None = None
    writer: str | None = None
    confidence: float
    selected: bool = True
    config_path: str | None = None
    output_paths: list[str] = Field(default_factory=list)
    reasons: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)


class ManifestOverride(BaseModel):
    override_id: str = Field(default_factory=lambda: f"override_{uuid4().hex[:12]}")
    target: str
    original_value: Any = None
    override_value: Any
    reason: str
    source: str | None = None
    approved_by: str | None = None
    created_at: str = Field(default_factory=lambda: utc_now())


class ManifestEvent(BaseModel):
    event_id: str = Field(default_factory=lambda: f"event_{uuid4().hex[:12]}")
    event_type: str
    status: str
    started_at: str | None = None
    finished_at: str | None = None
    component: str | None = None
    input_ids: list[str] = Field(default_factory=list)
    config_path: str | None = None
    output_paths: list[str] = Field(default_factory=list)
    summary: dict[str, Any] = Field(default_factory=dict)
    warnings: list[str] = Field(default_factory=list)


class RunManifest(BaseModel):
    manifest_version: str = "1"
    run_id: str = Field(default_factory=lambda: f"run_{uuid4().hex[:12]}")
    package_root: str
    status: ManifestStatus = "discovered"
    period_label: str | None = None
    year: int | None = None
    month: int | None = None
    created_at: str = Field(default_factory=lambda: utc_now())
    updated_at: str = Field(default_factory=lambda: utc_now())
    inputs: list[ManifestInput] = Field(default_factory=list)
    overrides: list[ManifestOverride] = Field(default_factory=list)
    events: list[ManifestEvent] = Field(default_factory=list)
    writer_config_path: str | None = None
    values_path: str | None = None
    output_workbook_path: str | None = None
    validation_report_path: str | None = None
    warnings: list[str] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)

    def to_json_file(self, path: Path) -> None:
        self.updated_at = utc_now()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.model_dump_json(indent=2) + "\n", encoding="utf-8")

    def input_by_id(self, input_id: str) -> ManifestInput | None:
        for item in self.inputs:
            if item.input_id == input_id:
                return item
        return None

    def selected_inputs(self) -> list[ManifestInput]:
        return [item for item in self.inputs if item.selected]


def manifest_from_discovery(
    discovery: PackageDiscovery,
    *,
    period_label: str | None = None,
    year: int | None = None,
    month: int | None = None,
) -> RunManifest:
    inferred_period = infer_period(discovery.files)
    manifest = RunManifest(
        package_root=discovery.root_path,
        period_label=period_label or inferred_period.get("period_label"),
        year=year or inferred_period.get("year"),
        month=month or inferred_period.get("month"),
        inputs=[manifest_input_from_file(file) for file in discovery.files],
        warnings=list(discovery.warnings),
    )
    manifest.events.append(
        ManifestEvent(
            event_type="discovery",
            status="completed",
            started_at=discovery.discovered_at,
            finished_at=utc_now(),
            summary={
                "file_count": len(discovery.files),
                "selected_count": len(manifest.selected_inputs()),
            },
            warnings=list(discovery.warnings),
        )
    )
    return manifest


def manifest_input_from_file(file: PackageFile) -> ManifestInput:
    return ManifestInput(
        input_id=stable_input_id(file),
        path=file.path,
        relative_path=file.relative_path,
        file_name=file.file_name,
        role=file.role,
        document_type=file.document_type,
        extractor=file.extractor,
        writer=file.writer,
        confidence=file.confidence,
        selected=file.role not in {"unknown"},
        reasons=file.reasons,
        metadata=file.metadata,
    )


def stable_input_id(file: PackageFile) -> str:
    stem = file.relative_path.lower()
    normalized = "".join(char if char.isalnum() else "_" for char in stem).strip("_")
    return f"input_{normalized[:80]}"


def load_manifest(path: Path) -> RunManifest:
    return RunManifest.model_validate_json(path.read_text(encoding="utf-8"))


def write_manifest(path: Path, manifest: RunManifest) -> None:
    manifest.to_json_file(path)


def manifest_summary(manifest: RunManifest) -> dict[str, Any]:
    by_component: dict[str, int] = {}
    by_role: dict[str, int] = {}
    for item in manifest.inputs:
        by_role[item.role] = by_role.get(item.role, 0) + 1
        component = item.extractor or (f"writer:{item.writer}" if item.writer else None)
        if component:
            by_component[component] = by_component.get(component, 0) + 1
    return {
        "run_id": manifest.run_id,
        "status": manifest.status,
        "package_root": manifest.package_root,
        "period_label": manifest.period_label,
        "input_count": len(manifest.inputs),
        "selected_input_count": len(manifest.selected_inputs()),
        "event_count": len(manifest.events),
        "override_count": len(manifest.overrides),
        "by_role": dict(sorted(by_role.items())),
        "by_component": dict(sorted(by_component.items())),
        "warnings": manifest.warnings,
    }


def manifest_summary_json(manifest: RunManifest) -> str:
    return json.dumps(manifest_summary(manifest), indent=2) + "\n"


def infer_period(files: list[PackageFile]) -> dict[str, Any]:
    text = " ".join(file.file_name for file in files)
    year_match = None
    for match in re_find_years(text):
        year_match = match
        if match == 2026:
            break
    month_match = re_find_month(text)
    period_label = None
    if month_match and year_match:
        period_label = f"{month_name(month_match)} {year_match}"
    return {"year": year_match, "month": month_match, "period_label": period_label}


def re_find_years(text: str) -> list[int]:
    import re

    return [int(match) for match in re.findall(r"\b(20\d{2}|19\d{2})\b", text)]


def re_find_month(text: str) -> int | None:
    import re

    months = {
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
    }
    normalized = re.sub(r"[^a-z]+", " ", text.lower())
    for name, number in months.items():
        if re.search(rf"\b{name}\b", normalized):
            return number
    return None


def month_name(month_num: int) -> str:
    return datetime(2000, month_num, 1).strftime("%B")


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()
