from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import json
import re
from typing import Any

import pdfplumber
import polars as pl
from pydantic import BaseModel, Field


MONEY_RE = re.compile(r"-?\$?[\d,]+(?:\.\d+)?|\$0")
PCT_RE = re.compile(r"-?\d+(?:\.\d+)?%")
MONTHLY_LINE_RE = re.compile(
    r"^(?P<year>20\d{2})\s*\|\s*(?P<month_num>\d{1,2})\.\s*"
    r"\((?P<month_name>[^)]+)\)\s+(?P<body>.+)$"
)
CUSTOMER_AMOUNT_RE = re.compile(r"^(?P<label>.+?)\s+(?P<amount>-?\$?[\d,]+(?:\.\d+)?)$")
RECON_ROW_RE = re.compile(
    r"^(?P<customer>.+?)\s+"
    r"(?P<cb>-?\$?[\d,]+(?:\.\d+)?)\s+"
    r"(?P<qb>-|-?\$?[\d,]+(?:\.\d+)?)\s+"
    r"(?P<diff>-?\$?[\d,]+(?:\.\d+)?)"
    r"(?:\s+(?P<note>.*))?$"
)


CATEGORY_ORDER = [
    "allowance",
    "penalty",
    "provision",
    "return",
    "software_fees",
]


class ChargebackPDFConfig(BaseModel):
    """Agent-adjustable extraction rules for chargeback report email PDFs."""

    target_month_name: str | None = "March"
    target_year: int | None = 2026
    monthly_category_order: list[str] = Field(default_factory=lambda: CATEGORY_ORDER.copy())
    customer_detail_start_patterns: list[str] = Field(
        default_factory=lambda: [r"^Month Department Customer .*SUM of Deduction Amount$"]
    )
    customer_detail_stop_patterns: list[str] = Field(
        default_factory=lambda: [r"^Customer As per CB Report As per QB Difference$"]
    )
    reconciliation_start_patterns: list[str] = Field(
        default_factory=lambda: [r"^Customer As per CB Report As per QB Difference$"]
    )
    reconciliation_stop_patterns: list[str] = Field(
        default_factory=lambda: [r"^--$", r"^We appreciate your continued partnership!$"]
    )

    @classmethod
    def from_json_file(cls, path: Path | None) -> "ChargebackPDFConfig":
        if path is None:
            return cls()
        return cls.model_validate_json(path.read_text(encoding="utf-8"))


@dataclass(frozen=True)
class PDFLine:
    page: int
    line_number: int
    text: str


@dataclass
class ChargebackPDFExtraction:
    path: Path
    subject: str | None
    lines: list[PDFLine]
    monthly_summary_rows: list[dict[str, Any]]
    customer_detail_rows: list[dict[str, Any]]
    reconciliation_rows: list[dict[str, Any]]
    notes: list[dict[str, Any]]
    warnings: list[str] = field(default_factory=list)

    @property
    def monthly_summary(self) -> pl.DataFrame:
        return pl.DataFrame(self.monthly_summary_rows)

    @property
    def customer_detail(self) -> pl.DataFrame:
        return pl.DataFrame(self.customer_detail_rows)

    @property
    def reconciliation(self) -> pl.DataFrame:
        return pl.DataFrame(self.reconciliation_rows)

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "subject": self.subject,
            "warnings": self.warnings,
            "monthly_summary_rows": self.monthly_summary_rows,
            "customer_detail_rows": self.customer_detail_rows,
            "reconciliation_rows": self.reconciliation_rows,
            "notes": self.notes,
        }


@dataclass
class ChargebackPDFProfile:
    path: Path
    page_count: int
    lines: list[PDFLine]
    table_summaries: list[dict[str, Any]]
    monthly_line_candidates: list[dict[str, Any]]
    anchor_candidates: list[dict[str, Any]]
    suggested_config: ChargebackPDFConfig

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "page_count": self.page_count,
            "line_count": len(self.lines),
            "table_summaries": self.table_summaries,
            "monthly_line_candidates": self.monthly_line_candidates,
            "anchor_candidates": self.anchor_candidates,
            "suggested_config": self.suggested_config.model_dump(),
        }


def extract_chargeback_pdf(
    path: Path,
    config: ChargebackPDFConfig | None = None,
) -> ChargebackPDFExtraction:
    config = config or ChargebackPDFConfig()
    lines = extract_pdf_lines(path)
    subject = detect_subject(lines)
    monthly_summary_rows = parse_monthly_summary(lines, config)
    customer_detail_rows = parse_customer_detail(lines, config)
    reconciliation_rows, notes = parse_reconciliation(lines, config)
    warnings = validate_extraction(monthly_summary_rows, customer_detail_rows, reconciliation_rows, config)
    return ChargebackPDFExtraction(
        path=path,
        subject=subject,
        lines=lines,
        monthly_summary_rows=monthly_summary_rows,
        customer_detail_rows=customer_detail_rows,
        reconciliation_rows=reconciliation_rows,
        notes=notes,
        warnings=warnings,
    )


def profile_chargeback_pdf(path: Path) -> ChargebackPDFProfile:
    lines = extract_pdf_lines(path)
    table_summaries = extract_table_summaries(path)
    suggested_config = suggest_config(lines)
    return ChargebackPDFProfile(
        path=path,
        page_count=page_count(path),
        lines=lines,
        table_summaries=table_summaries,
        monthly_line_candidates=find_monthly_line_candidates(lines),
        anchor_candidates=find_anchor_candidates(lines),
        suggested_config=suggested_config,
    )


def extract_pdf_lines(path: Path) -> list[PDFLine]:
    lines: list[PDFLine] = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            for line_number, line in enumerate(text.splitlines(), start=1):
                normalized = normalize_text(line)
                if normalized:
                    lines.append(PDFLine(page_number, line_number, normalized))
    return lines


def page_count(path: Path) -> int:
    with pdfplumber.open(path) as pdf:
        return len(pdf.pages)


def extract_table_summaries(path: Path) -> list[dict[str, Any]]:
    summaries: list[dict[str, Any]] = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            for table_index, table in enumerate(page.extract_tables(), start=1):
                row_count = len(table)
                col_count = max((len(row) for row in table), default=0)
                preview = [
                    [normalize_text(cell) for cell in row[:8]]
                    for row in table[:6]
                ]
                summaries.append(
                    {
                        "page": page_number,
                        "table_index": table_index,
                        "row_count": row_count,
                        "column_count": col_count,
                        "preview": preview,
                    }
                )
    return summaries


def find_monthly_line_candidates(lines: list[PDFLine]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for line in lines:
        match = MONTHLY_LINE_RE.match(line.text)
        if not match:
            continue
        token_count = len(extract_amount_percent_tokens(match.group("body")))
        candidates.append(
            {
                "page": line.page,
                "line": line.line_number,
                "year": int(match.group("year")),
                "month_num": int(match.group("month_num")),
                "month_name": normalize_month_name(match.group("month_name")),
                "token_count": token_count,
                "text": line.text,
            }
        )
    return candidates


ANCHOR_HINTS = [
    "Report Name",
    "Month Department",
    "Customer",
    "Difference",
    "Grand Total",
    "Total",
    "SUM of Deduction",
    "As per CB Report",
    "As per QB",
]


def find_anchor_candidates(lines: list[PDFLine]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for line in lines:
        hits = [hint for hint in ANCHOR_HINTS if hint.lower() in line.text.lower()]
        if not hits:
            continue
        candidates.append(
            {
                "page": line.page,
                "line": line.line_number,
                "hints": hits,
                "text": line.text,
            }
        )
    return candidates


def suggest_config(lines: list[PDFLine]) -> ChargebackPDFConfig:
    config = ChargebackPDFConfig()
    subject = detect_subject(lines) or ""
    subject_match = re.search(r"(\d{2})\.\s*([A-Za-z]+)\s+(20\d{2})", subject)
    if subject_match:
        config.target_month_name = normalize_month_name(subject_match.group(2))
        config.target_year = int(subject_match.group(3))

    customer_anchor = first_matching_text(lines, [r"^Month Department Customer .*SUM of Deduction Amount$"])
    if customer_anchor:
        config.customer_detail_start_patterns = [f"^{re.escape(customer_anchor)}$"]

    reconciliation_anchor = first_matching_text(
        lines,
        [r"^Customer As per CB Report As per QB Difference$"],
    )
    if reconciliation_anchor:
        escaped = f"^{re.escape(reconciliation_anchor)}$"
        config.customer_detail_stop_patterns = [escaped]
        config.reconciliation_start_patterns = [escaped]

    stop_anchor = first_matching_text(lines, [r"^--$", r"^We appreciate your continued partnership!$"])
    if stop_anchor:
        config.reconciliation_stop_patterns = [f"^{re.escape(stop_anchor)}$"]
    return config


def first_matching_text(lines: list[PDFLine], patterns: list[str]) -> str | None:
    for line in lines:
        if matches_any(line.text, patterns):
            return line.text
    return None


def detect_subject(lines: list[PDFLine]) -> str | None:
    for line in lines:
        if line.text.startswith("OG | Chargeback Report"):
            return line.text
    for line in lines:
        if "Chargeback Report" in line.text:
            return line.text
    return None


def parse_monthly_summary(
    lines: list[PDFLine],
    config: ChargebackPDFConfig,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for line in lines:
        match = MONTHLY_LINE_RE.match(line.text)
        if not match:
            continue

        year = int(match.group("year"))
        month_num = int(match.group("month_num"))
        month_name = normalize_month_name(match.group("month_name"))
        body = match.group("body")

        tokens = extract_amount_percent_tokens(body)
        if len(tokens) < 2:
            continue

        amount_percent_pairs = tokens[:-1]
        grand_total = tokens[-1]["amount"]
        for idx, pair in enumerate(amount_percent_pairs):
            category = category_for_position(idx, len(amount_percent_pairs), config)
            rows.append(
                {
                    "year": year,
                    "month_num": month_num,
                    "month_name": month_name,
                    "category": category,
                    "category_position": idx,
                    "amount": pair["amount"],
                    "percent_of_total": pair.get("percent"),
                    "grand_total": grand_total,
                    "source_page": line.page,
                    "source_line": line.line_number,
                    "source_text": line.text,
                    "confidence": "high" if len(amount_percent_pairs) == len(CATEGORY_ORDER) else "medium",
                }
            )

        rows.append(
            {
                "year": year,
                "month_num": month_num,
                "month_name": month_name,
                "category": "grand_total",
                "category_position": len(amount_percent_pairs),
                "amount": grand_total,
                "percent_of_total": None,
                "grand_total": grand_total,
                "source_page": line.page,
                "source_line": line.line_number,
                "source_text": line.text,
                "confidence": "high",
            }
        )
    return rows


def extract_amount_percent_tokens(body: str) -> list[dict[str, float | None]]:
    parts = re.findall(r"-?\d+(?:\.\d+)?%|-?\$?[\d,]+(?:\.\d+)?|\$0", body)
    tokens: list[dict[str, float | None]] = []
    idx = 0
    while idx < len(parts):
        amount = parse_money(parts[idx])
        pct: float | None = None
        if idx + 1 < len(parts) and parts[idx + 1].endswith("%"):
            pct = parse_percent(parts[idx + 1])
            idx += 2
        else:
            idx += 1
        tokens.append({"amount": amount, "percent": pct})
    return tokens


def category_for_position(
    idx: int,
    pair_count: int,
    config: ChargebackPDFConfig,
) -> str:
    categories = config.monthly_category_order
    if pair_count == len(categories):
        return categories[idx]
    if pair_count == 4 and idx == 2:
        return "return_or_provision"
    if idx < len(categories):
        return categories[idx]
    return f"category_{idx + 1}"


def parse_customer_detail(
    lines: list[PDFLine],
    config: ChargebackPDFConfig,
) -> list[dict[str, Any]]:
    detail_lines = lines_between(
        lines,
        config.customer_detail_start_patterns,
        config.customer_detail_stop_patterns,
    )
    rows: list[dict[str, Any]] = []
    current_month: str | None = None
    current_department: str | None = None

    for line in detail_lines:
        match = CUSTOMER_AMOUNT_RE.match(line.text)
        if not match:
            continue
        label = match.group("label").strip()
        amount = parse_money(match.group("amount"))
        parts = label.split()
        is_total = False
        month: str | None = current_month
        department: str | None = current_department
        customer: str

        if len(parts) >= 3 and parts[0].lower() in MONTH_NAMES and parts[1] in {"B&M", "Online"}:
            month = normalize_month_name(parts[0])
            department = parts[1]
            customer = " ".join(parts[2:])
            current_month = month
            current_department = department
        elif len(parts) >= 2 and parts[-1].lower() == "total":
            is_total = True
            if parts[0].lower() in MONTH_NAMES:
                month = normalize_month_name(parts[0])
                department = None
                customer = "Total"
            elif parts[0].lower() == "grand":
                month = current_month
                department = None
                customer = "Grand Total"
            else:
                department = " ".join(parts[:-1])
                customer = "Total"
                current_department = department
        else:
            customer = label

        rows.append(
            {
                "month_name": month,
                "department": department,
                "customer": customer,
                "amount": amount,
                "is_total_row": is_total,
                "source_page": line.page,
                "source_line": line.line_number,
                "source_text": line.text,
            }
        )
    return rows


def parse_reconciliation(
    lines: list[PDFLine],
    config: ChargebackPDFConfig,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    recon_lines = lines_between(
        lines,
        config.reconciliation_start_patterns,
        config.reconciliation_stop_patterns,
    )
    rows: list[dict[str, Any]] = []
    notes: list[dict[str, Any]] = []

    for line in recon_lines:
        match = RECON_ROW_RE.match(line.text)
        if not match:
            notes.append(
                {
                    "source_page": line.page,
                    "source_line": line.line_number,
                    "text": line.text,
                }
            )
            continue

        customer = match.group("customer").strip()
        rows.append(
            {
                "customer": customer,
                "cb_report_amount": parse_money(match.group("cb")),
                "qb_amount": None if match.group("qb") == "-" else parse_money(match.group("qb")),
                "difference": parse_money(match.group("diff")),
                "note": normalize_text(match.group("note")),
                "is_total_row": customer.lower().endswith("total"),
                "source_page": line.page,
                "source_line": line.line_number,
                "source_text": line.text,
            }
        )
    return rows, notes


def validate_extraction(
    monthly_summary_rows: list[dict[str, Any]],
    customer_detail_rows: list[dict[str, Any]],
    reconciliation_rows: list[dict[str, Any]],
    config: ChargebackPDFConfig,
) -> list[str]:
    warnings: list[str] = []
    if config.target_year and config.target_month_name:
        target_summary = [
            row
            for row in monthly_summary_rows
            if row["year"] == config.target_year
            and row["month_name"].lower() == config.target_month_name.lower()
            and row["category"] == "grand_total"
        ]
        if not target_summary:
            warnings.append("Target month grand total not found in monthly summary")

    if not customer_detail_rows:
        warnings.append("No customer detail rows parsed")
    if not reconciliation_rows:
        warnings.append("No reconciliation rows parsed")
    return warnings


def lines_between(
    lines: list[PDFLine],
    start_patterns: list[str],
    stop_patterns: list[str],
) -> list[PDFLine]:
    in_block = False
    block: list[PDFLine] = []
    for line in lines:
        if not in_block and matches_any(line.text, start_patterns):
            in_block = True
            continue
        if in_block and matches_any(line.text, stop_patterns):
            break
        if in_block:
            block.append(line)
    return block


MONTH_NAMES = {
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december",
}


def normalize_month_name(value: str) -> str:
    return normalize_text(value).strip().capitalize()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ")).strip()


def parse_money(value: str) -> float:
    text = value.strip().replace("$", "").replace(",", "")
    if text in {"", "-"}:
        return 0.0
    return float(text)


def parse_percent(value: str) -> float:
    return float(value.strip().replace("%", "")) / 100


def matches_any(value: str, patterns: list[str]) -> bool:
    return any(re.search(pattern, value, flags=re.IGNORECASE) for pattern in patterns)


def write_default_config(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(ChargebackPDFConfig().model_dump(), indent=2) + "\n",
        encoding="utf-8",
    )


def write_profile_artifacts(profile: ChargebackPDFProfile, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "chargeback_pdf_profile.json").write_text(
        json.dumps(profile.to_dict(), indent=2),
        encoding="utf-8",
    )
    (output_dir / "chargeback_pdf_raw_text.txt").write_text(
        "\n".join(
            f"p{line.page}:l{line.line_number}: {line.text}"
            for line in profile.lines
        )
        + "\n",
        encoding="utf-8",
    )
    (output_dir / "chargeback_pdf_suggested_config.json").write_text(
        json.dumps(profile.suggested_config.model_dump(), indent=2) + "\n",
        encoding="utf-8",
    )
    (output_dir / "chargeback_pdf_profile.md").write_text(
        render_profile_markdown(profile),
        encoding="utf-8",
    )


def render_profile_markdown(profile: ChargebackPDFProfile) -> str:
    lines = [
        "# Chargeback PDF Profile",
        "",
        f"File: `{profile.path}`",
        f"Pages: `{profile.page_count}`",
        f"Text lines: `{len(profile.lines)}`",
        "",
        "## Suggested Config",
        "",
        "```json",
        json.dumps(profile.suggested_config.model_dump(), indent=2),
        "```",
        "",
        "## Monthly Line Candidates",
        "",
    ]
    for candidate in profile.monthly_line_candidates[:40]:
        lines.append(
            f"- p{candidate['page']}:l{candidate['line']} "
            f"{candidate['year']} {candidate['month_name']} "
            f"tokens={candidate['token_count']}: `{candidate['text']}`"
        )
    lines.extend(["", "## Anchor Candidates", ""])
    for candidate in profile.anchor_candidates[:80]:
        lines.append(
            f"- p{candidate['page']}:l{candidate['line']} "
            f"hints={candidate['hints']}: `{candidate['text']}`"
        )
    lines.extend(["", "## Table Summaries", ""])
    for table in profile.table_summaries:
        lines.append(
            f"- page {table['page']} table {table['table_index']}: "
            f"{table['row_count']} rows x {table['column_count']} columns"
        )
        for row in table["preview"]:
            lines.append(f"  - `{' | '.join(row)}`")
    return "\n".join(lines) + "\n"
