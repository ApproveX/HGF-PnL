from __future__ import annotations

from dataclasses import dataclass, field
import json
import re
from pathlib import Path
from typing import Any

import polars as pl


FrameMap = dict[str, pl.DataFrame]


@dataclass
class CloseValuesResult:
    values: dict[str, Any]
    warnings: list[str] = field(default_factory=list)

    @property
    def populated_key_count(self) -> int:
        return len(flatten_values(self.values))

    def to_dict(self) -> dict[str, Any]:
        return {
            "values": self.values,
            "warnings": self.warnings,
            "populated_key_count": self.populated_key_count,
        }


def build_consolidated_values(
    *,
    pl_by_dept: pl.DataFrame | None = None,
    br_info: pl.DataFrame | None = None,
    monthly_revenue_summary: pl.DataFrame | None = None,
    monthly_revenue_sales: pl.DataFrame | None = None,
    monthly_revenue_refunds: pl.DataFrame | None = None,
    division_cogs_matrix: pl.DataFrame | None = None,
    th_revenue_summary: pl.DataFrame | None = None,
    payroll_employees: pl.DataFrame | None = None,
    payroll_allocation_summaries: pl.DataFrame | None = None,
    payroll_distribution: pl.DataFrame | None = None,
    chargeback_customer_detail: pl.DataFrame | None = None,
    year: int | None = None,
    month_num: int | None = None,
) -> CloseValuesResult:
    values: dict[str, Any] = {}
    warnings: list[str] = []

    br = filter_period(br_info, year=year, month_num=month_num)
    pl_df = ensure_frame(pl_by_dept)
    cogs = filter_period(division_cogs_matrix, year=year, month_num=month_num)

    if not pl_df.is_empty():
        add_pl_by_dept_values(values, warnings, pl_df)
    if not br.is_empty():
        add_br_info_values(values, warnings, br)
    if not ensure_frame(monthly_revenue_sales).is_empty() or not ensure_frame(
        monthly_revenue_refunds
    ).is_empty():
        add_monthly_revenue_values(
            values,
            warnings,
            ensure_frame(monthly_revenue_summary),
            ensure_frame(monthly_revenue_sales),
            ensure_frame(monthly_revenue_refunds),
        )
    if not ensure_frame(th_revenue_summary).is_empty():
        add_th_revenue_values(values, warnings, ensure_frame(th_revenue_summary))
    if not cogs.is_empty():
        add_division_cogs_values(values, warnings, cogs)
    if not ensure_frame(chargeback_customer_detail).is_empty():
        add_chargeback_values(values, warnings, ensure_frame(chargeback_customer_detail))
    if not ensure_frame(payroll_employees).is_empty() or not ensure_frame(
        payroll_allocation_summaries
    ).is_empty():
        add_payroll_values(
            values,
            warnings,
            ensure_frame(payroll_employees),
            ensure_frame(payroll_allocation_summaries),
            ensure_frame(payroll_distribution),
        )

    add_default_zero_values(values)
    return CloseValuesResult(values=values, warnings=warnings)


def add_pl_by_dept_values(values: dict[str, Any], warnings: list[str], df: pl.DataFrame) -> None:
    gl_specs = [
        ("raw_master.gl.warehouse_rent_curci", ("5020", "warehouse", "rent"), "Total"),
        ("raw_master.gl.consulting_expense", ("6060", "consulting"), "Total Z-COMPANY"),
        ("raw_master.gl.rent_office", ("6070", "rent", "office"), "Total"),
        ("raw_master.gl.rent_showrooms", ("6080", "rent", "showrooms"), "Total"),
        ("raw_master.gl.equipment_lease", ("6090", "equipment", "lease"), "Total"),
        ("raw_master.gl.repairs_maintenance", ("6100", "repairs"), "Total"),
        ("raw_master.gl.vehicle_expense", ("6150", "vehicle"), "Total"),
        ("raw_master.gl.office_supplies", ("6300", "office", "supplies"), "Total"),
        ("raw_master.gl.art_assets", ("6310", "art", "assets"), "Total"),
        ("raw_master.gl.employee_nurturing", ("6320", "employee", "nurturing"), "Total"),
        ("raw_master.gl.software_web_services", ("software", "web"), "Total Z-COMPANY"),
        ("raw_master.gl.bank_fees", ("6410", "bank", "fees"), "Total"),
        ("raw_master.gl.merchant_account_fees", ("merchant", "account", "fees"), "Total"),
        ("raw_master.gl.cleaning_janitorial", ("cleaning", "janitorial"), "Total"),
        ("raw_master.gl.hr_recruiting", ("hr", "recruiting"), "Operations Dept"),
        ("raw_master.gl.insurance", ("6450", "insurance"), "Total"),
        ("raw_master.gl.telephone_internet", ("telephone", "internet"), "Total"),
        ("raw_master.gl.utilities", ("6550", "utilities"), "Total"),
        ("raw_master.gl.licenses_taxes_permits", ("licenses", "taxes", "permits"), "Total"),
        ("raw_master.gl.dues_subscriptions", ("dues", "subscriptions"), "Total"),
        ("raw_master.gl.meals_entertainment", ("6700", "meals"), "Total Z-COMPANY"),
        ("raw_master.gl.professional_fees_accounting", ("professional", "accounting"), "Total"),
        ("raw_master.gl.professional_fees_legal", ("professional", "legal"), "Total"),
    ]
    for key, needles, department in gl_specs:
        set_if_found(values, warnings, key, pl_lookup(df, needles, department))

    split_specs = [
        (
            "raw_master.consulting",
            ("6060", "consulting"),
            {
                "corp": "Total Z-COMPANY",
                "trend_house": "Brick Mortar - China",
                "dtc": "OG-DTC",
                "online": "Online",
                "online_lux": "Z-COMPANY",
            },
        ),
        (
            "raw_master.software_web",
            ("software", "web"),
            {
                "corp": "Total Z-COMPANY",
                "dtc": "OG-DTC",
                "online": "Online",
                "online_lux": "Z-COMPANY",
            },
        ),
        (
            "raw_master.travel",
            ("6590", "travel"),
            {
                "corp": "Total Z-COMPANY",
                "dtc": "OG-DTC",
                "trend_house": "Brick Mortar - China",
            },
        ),
        (
            "raw_master.meals",
            ("6700", "meals"),
            {
                "corp": "Total Z-COMPANY",
                "dtc": "OG-DTC",
                "trend_house": "Brick Mortar - China",
            },
        ),
        (
            "raw_master.advertising",
            ("6200", "advertising"),
            {
                "dtc": "OG-DTC",
                "online": "Online",
                "online_lux": "Z-COMPANY",
                "trend_house": "Brick Mortar - China",
            },
        ),
    ]
    for namespace, needles, departments in split_specs:
        for child_key, department in departments.items():
            value = pl_lookup(df, needles, department)
            set_nested_value(values, f"{namespace}.{child_key}", value if value is not None else 0.0)


def add_br_info_values(values: dict[str, Any], warnings: list[str], df: pl.DataFrame) -> None:
    br = br_info_dict(df)
    direct = {
        "Online Sales": "raw_master.sales.online",
        "AllPopArt Sales": "raw_master.sales.apa",
        "AllPopArt Returns and Allowances": "raw_master.returns.apa",
        "Employee Benefits": "full_report.source_totals.employee_benefits",
        "License & Tax": "raw_master.gl.licenses_taxes_permits",
        "LOC Interest": "raw_master.gl.loc_interest",
    }
    for label, key in direct.items():
        if label in br:
            set_nested_value(values, key, br[label])

    replacements = {
        "Bank Fees": ("raw_master.gl.bank_fees", "raw_master.gl.bank_fees_adjustment"),
        "Merchant Account Fees": (
            "raw_master.gl.merchant_account_fees",
            "raw_master.gl.merchant_account_fees_adjustment",
        ),
        "Equipment Leasing": (
            "raw_master.gl.equipment_lease",
            "raw_master.gl.equipment_lease_adjustment",
        ),
    }
    for label, (base_key, adjustment_key) in replacements.items():
        if label in br:
            set_nested_value(values, base_key, br[label])
            set_nested_value(values, adjustment_key, 0.0)

    known = set(direct) | set(replacements)
    for label in sorted(set(br) - known):
        warnings.append(f"Unmapped BR Info row requires review: {label}")


def add_monthly_revenue_values(
    values: dict[str, Any],
    warnings: list[str],
    summary: pl.DataFrame,
    sales: pl.DataFrame,
    refunds: pl.DataFrame,
) -> None:
    if not sales.is_empty() and {"channel", "net_sales"} <= set(sales.columns):
        total_sales = numeric_sum(sales, "net_sales")
        set_nested_value(values, "raw_master.sales.dtc", total_sales)
    else:
        revenue_total = summary_amount(summary, "REVENUE", "Grand Total")
        set_if_found(values, warnings, "raw_master.sales.dtc", revenue_total)

    if not refunds.is_empty() and "amount" in refunds.columns:
        if "has_amount" in refunds.columns:
            refunds = refunds.filter(pl.col("has_amount").fill_null(False))
        set_nested_value(values, "raw_master.returns.dtc", -numeric_sum(refunds, "amount"))
    else:
        refund_total = summary_amount(summary, "REFUNDS", "Grand Total")
        if refund_total is not None:
            set_nested_value(values, "raw_master.returns.dtc", -refund_total)


def add_th_revenue_values(values: dict[str, Any], warnings: list[str], df: pl.DataFrame) -> None:
    total = df.filter(pl.col("is_total_row").fill_null(False))
    if total.is_empty():
        total = df
    set_if_found(values, warnings, "raw_cogs.trend_house.total.revenue", frame_sum(total, "revenue"))
    set_if_found(
        values,
        warnings,
        "raw_cogs.trend_house.total.production_cost",
        frame_sum(total, "production_cost"),
    )
    set_if_found(
        values,
        warnings,
        "raw_cogs.trend_house.total.shipping_cost",
        frame_sum(total, "shipping_cost"),
    )
    tariff = frame_sum(total, "tariff")
    if tariff:
        set_nested_value(values, "raw_cogs.tariffs", tariff)


def add_division_cogs_values(values: dict[str, Any], warnings: list[str], df: pl.DataFrame) -> None:
    mappings = [
        ("raw_cogs.current_month.cogs.bm_usa_samples", "COGS", ("sample",), ()),
        ("raw_cogs.material.bm_usa_samples", "Material Cogs", ("sample",), ()),
        ("raw_cogs.labor.bm_usa_samples", "Labor Cogs", ("sample",), ()),
        ("raw_cogs.current_month.cogs.online_textiles_mww", "COGS", ("online", "textiles"), ()),
        ("raw_cogs.current_month.cogs.og_dtc", "COGS", ("og", "dtc"), ("returns",)),
        (
            "raw_cogs.current_month.cogs.og_dtc_returns",
            "COGS",
            ("returns", "replacements"),
            (),
        ),
        ("raw_cogs.current_month.cogs.all_pop_art", "COGS", ("all", "pop", "art"), ()),
        ("raw_cogs.material.og_dtc", "Material Cogs", ("og", "dtc"), ("returns",)),
        ("raw_cogs.material.og_dtc_returns", "Material Cogs", ("returns", "replacements"), ()),
        ("raw_cogs.material.all_pop_art", "Material Cogs", ("all", "pop", "art"), ()),
        ("raw_cogs.labor.og_dtc", "Labor Cogs", ("og", "dtc"), ("returns",)),
        ("raw_cogs.labor.og_dtc_returns", "Labor Cogs", ("returns", "replacements"), ()),
        ("raw_cogs.labor.all_pop_art", "Labor Cogs", ("all", "pop", "art"), ()),
        ("raw_cogs.shipping_actual.online", "SHIPPING (LTL & SP) Actual", ("prev", "trade"), ()),
        (
            "raw_cogs.shipping_actual.online_usa",
            "SHIPPING (LTL & SP) Actual",
            ("online", "usa"),
            (),
        ),
        (
            "raw_cogs.shipping_actual.og_dtc",
            "SHIPPING (LTL & SP) Actual",
            ("og", "dtc"),
            ("returns",),
        ),
        (
            "raw_cogs.shipping_actual.og_dtc_returns",
            "SHIPPING (LTL & SP) Actual",
            ("returns", "replacements"),
            (),
        ),
        (
            "raw_cogs.shipping_actual.all_pop_art",
            "SHIPPING (LTL & SP) Actual",
            ("all", "pop", "art"),
            (),
        ),
        ("raw_cogs.fedex.online", "FEDEX", ("prev", "trade"), ()),
        ("raw_cogs.fedex.online_usa", "FEDEX", ("online", "usa"), ()),
        ("raw_cogs.fedex.og_dtc", "FEDEX", ("og", "dtc"), ("returns",)),
        ("raw_cogs.fedex.og_dtc_returns", "FEDEX", ("returns", "replacements"), ()),
        ("raw_cogs.fedex.all_pop_art", "FEDEX", ("all", "pop", "art"), ()),
        ("raw_cogs.fedex.corporate_shipping", "FEDEX", ("corporate", "shipping"), ()),
        ("raw_cogs.ups.online_usa", "UPS", ("online", "usa"), ()),
        (
            "raw_cogs.shipping_quoted_dtc.og_dtc",
            "SHIPPING (LTL & SP) Quoted D2C",
            ("og", "dtc"),
            ("returns",),
        ),
    ]
    for key, metric, channel_needles, excluded_needles in mappings:
        set_nested_value(
            values,
            key,
            cogs_lookup(df, metric, channel_needles, excluded_needles) or 0.0,
        )

    th_fedex = cogs_lookup(df, "FEDEX", ("brick", "mortar", "china"))
    if th_fedex is not None:
        set_nested_value(values, "raw_cogs.shipping_for_samples.current_month", th_fedex)

    for metric, writer_metric in [
        ("COGS", "current_month.cogs"),
        ("Material Cogs", "material"),
        ("Labor Cogs", "labor"),
    ]:
        online_standalone = cogs_lookup(df, metric, ("prev", "trade"))
        online_usa = cogs_lookup(df, metric, ("online", "usa"))
        if online_standalone is not None or online_usa is not None:
            set_nested_value(
                values,
                f"raw_cogs.{writer_metric}.online_usa",
                (online_usa or 0.0) + (online_standalone or 0.0),
            )
            set_nested_value(values, f"raw_cogs.{writer_metric}.online", 0.0)


def add_chargeback_values(values: dict[str, Any], warnings: list[str], df: pl.DataFrame) -> None:
    if {"department", "is_total_row", "amount"} <= set(df.columns):
        bm = df.filter(
            (pl.col("department").fill_null("") == "B&M") & pl.col("is_total_row").fill_null(False)
        )
        if not bm.is_empty():
            set_nested_value(values, "raw_master.returns.trend_house", numeric_sum(bm, "amount"))
        else:
            warnings.append("Chargeback B&M total row not found; TH returns not populated.")


def add_payroll_values(
    values: dict[str, Any],
    warnings: list[str],
    employees: pl.DataFrame,
    summaries: pl.DataFrame,
    distribution: pl.DataFrame,
) -> None:
    if not employees.is_empty() and {"section", "gross_pay"} <= set(employees.columns):
        section_map = {
            "Production": "raw_payroll.production",
            "Corp": "raw_payroll.corp.corp",
            "Art": "raw_payroll.corp.art",
            "IT": "raw_payroll.corp.it",
        }
        by_section = employees.group_by("section").agg(pl.col("gross_pay").sum().alias("amount"))
        for row in by_section.to_dicts():
            key = section_map.get(row["section"])
            if key:
                set_nested_value(values, key, float(row["amount"]))

    add_payroll_sales(values, summaries)
    add_payroll_allocation_breakdowns(values, summaries)
    add_lital_allocation(values, distribution)


def add_payroll_sales(values: dict[str, Any], summaries: pl.DataFrame) -> None:
    if summaries.is_empty():
        return
    mapping = {
        "Trend House": "raw_payroll.sales.trend_house",
        "DTC": "raw_payroll.sales.dtc",
        "Online": "raw_payroll.sales.online",
    }
    for label, key in mapping.items():
        value = payroll_summary_lookup(summaries, "Sales Dept", label)
        if value is not None:
            set_nested_value(values, key, value)
    set_nested_value(values, "raw_payroll.sales.og_specialty_usa", 0.0)
    set_nested_value(values, "raw_payroll.sales.online_lux", 0.0)


def add_payroll_allocation_breakdowns(values: dict[str, Any], summaries: pl.DataFrame) -> None:
    if summaries.is_empty():
        return
    mapping = {
        ("Art", "TH"): "raw_payroll.allocation_breakdowns.art.trend_house",
        ("Art", "B&M USA"): "raw_payroll.allocation_breakdowns.art.og_specialty_usa",
        ("Art", "Online Lux"): "raw_payroll.allocation_breakdowns.art.online_lux",
        ("Art", "Online"): "raw_payroll.allocation_breakdowns.art.online",
        ("Art", "OG DTC"): "raw_payroll.allocation_breakdowns.art.dtc",
        ("Art", "APA"): "raw_payroll.allocation_breakdowns.art.all_pop_art",
        ("Art", "Ink"): "raw_payroll.allocation_breakdowns.art.ink",
        ("Art", "General"): "raw_payroll.allocation_breakdowns.art.general",
        ("Art", "Total"): "raw_payroll.allocation_breakdowns.art.total",
        ("IT", "TH"): "raw_payroll.allocation_breakdowns.it.trend_house",
        ("IT", "B&M USA"): "raw_payroll.allocation_breakdowns.it.og_specialty_usa",
        ("IT", "Online Lux"): "raw_payroll.allocation_breakdowns.it.online_lux",
        ("IT", "Online"): "raw_payroll.allocation_breakdowns.it.online",
        ("IT", "OG DTC"): "raw_payroll.allocation_breakdowns.it.dtc",
        ("IT", "APA"): "raw_payroll.allocation_breakdowns.it.all_pop_art",
        ("IT", "Ink"): "raw_payroll.allocation_breakdowns.it.ink",
        ("IT", "General"): "raw_payroll.allocation_breakdowns.it.general",
        ("IT", "Total"): "raw_payroll.allocation_breakdowns.it.total",
    }
    for (department, category), key in mapping.items():
        set_nested_value(values, key, payroll_summary_lookup(summaries, department, category) or 0.0)


def add_lital_allocation(values: dict[str, Any], distribution: pl.DataFrame) -> None:
    if distribution.is_empty():
        return
    mapping = {
        "DTC": "raw_payroll.lital_allocation.dtc",
        "ONLINE": "raw_payroll.lital_allocation.online",
        "TH": "raw_payroll.lital_allocation.trend_house",
        "CORP": "raw_payroll.lital_allocation.corp",
    }
    rows = distribution.filter(pl.col("block").fill_null("") == "Lital Allocation in G&A Exp")
    for label, key in mapping.items():
        match = rows.filter(pl.col("label").fill_null("").str.to_uppercase() == label)
        if not match.is_empty():
            set_nested_value(values, key, float(match.select(pl.col("amount").sum()).item() or 0.0))


def pl_lookup(df: pl.DataFrame, line_needles: tuple[str, ...], department: str) -> float | None:
    if df.is_empty() or not {"line_item", "department", "amount"} <= set(df.columns):
        return None
    rows = df.filter(pl.col("department") == department)
    for needle in line_needles:
        rows = rows.filter(normalized_contains_expr("line_item", needle))
    if rows.is_empty():
        return None
    return float(rows.select(pl.col("amount").sum()).item() or 0.0)


def cogs_lookup(
    df: pl.DataFrame,
    metric: str,
    channel_needles: tuple[str, ...],
    excluded_needles: tuple[str, ...] = (),
) -> float | None:
    if df.is_empty() or not {"type", "channel", "amount"} <= set(df.columns):
        return None
    rows = df.filter(normalized_equals_expr("type", metric))
    for needle in channel_needles:
        rows = rows.filter(normalized_contains_expr("channel", needle))
    for needle in excluded_needles:
        rows = rows.filter(~normalized_contains_expr("channel", needle))
    if "formula" in rows.columns:
        rows = rows.filter(~pl.col("formula").fill_null("").str.contains("SUM", literal=False))
    if rows.is_empty():
        return None
    return float(rows.select(pl.col("amount").sum()).item() or 0.0)


def payroll_summary_lookup(df: pl.DataFrame, department: str, category: str) -> float | None:
    if df.is_empty() or not {"department", "allocation_category", "amount"} <= set(df.columns):
        return None
    rows = df.filter(
        (pl.col("department") == department) & (pl.col("allocation_category") == category)
    )
    if rows.is_empty():
        return None
    return float(rows.select(pl.col("amount").sum()).item() or 0.0)


def br_info_dict(df: pl.DataFrame) -> dict[str, float]:
    if df.is_empty():
        return {}
    result: dict[str, float] = {}
    for row in df.select(["override_name", "value"]).to_dicts():
        if row["value"] is not None:
            result[str(row["override_name"])] = float(row["value"])
    return result


def summary_amount(df: pl.DataFrame, section: str, label: str) -> float | None:
    if df.is_empty() or not {"section", "label", "amount"} <= set(df.columns):
        return None
    rows = df.filter((pl.col("section") == section) & (pl.col("label") == label))
    if rows.is_empty():
        return None
    return float(rows.select(pl.col("amount").sum()).item() or 0.0)


def frame_sum(df: pl.DataFrame, column: str) -> float | None:
    if df.is_empty() or column not in df.columns:
        return None
    return float(df.select(pl.col(column).fill_null(0).sum()).item() or 0.0)


def numeric_sum(df: pl.DataFrame, column: str) -> float:
    return float(df.select(pl.col(column).fill_null(0).sum()).item() or 0.0)


def set_if_found(
    values: dict[str, Any],
    warnings: list[str],
    key: str,
    value: float | None,
) -> None:
    if value is None:
        warnings.append(f"Value not found for {key}")
        return
    set_nested_value(values, key, value)


def set_nested_value(values: dict[str, Any], source_key: str, value: Any) -> None:
    current = values
    parts = source_key.split(".")
    for part in parts[:-1]:
        next_value = current.setdefault(part, {})
        if not isinstance(next_value, dict):
            raise ValueError(f"Cannot set nested value through scalar key: {source_key}")
        current = next_value
    current[parts[-1]] = value


def set_nested_default(values: dict[str, Any], source_key: str, value: Any) -> None:
    current = values
    parts = source_key.split(".")
    for part in parts[:-1]:
        next_value = current.setdefault(part, {})
        if not isinstance(next_value, dict):
            return
        current = next_value
    current.setdefault(parts[-1], value)


def add_default_zero_values(values: dict[str, Any]) -> None:
    zero_keys = [
        "raw_master.sales.ink",
        "raw_master.returns.og_specialty_usa",
        "raw_master.returns.og_specialty_trade",
        "raw_master.returns.ink",
        "raw_cogs.ups.online",
        "raw_cogs.ups.og_dtc",
        "raw_cogs.ups.og_dtc_returns",
        "raw_cogs.ups.all_pop_art",
    ]
    for key in zero_keys:
        set_nested_default(values, key, 0.0)


def flatten_values(values: dict[str, Any], prefix: str = "") -> dict[str, Any]:
    flat: dict[str, Any] = {}
    for key, value in values.items():
        full_key = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            flat.update(flatten_values(value, full_key))
        else:
            flat[full_key] = value
    return flat


def filter_period(
    df: pl.DataFrame | None,
    *,
    year: int | None,
    month_num: int | None,
) -> pl.DataFrame:
    frame = ensure_frame(df)
    if frame.is_empty():
        return frame
    if year is not None and "year" in frame.columns:
        frame = frame.filter(pl.col("year") == year)
    if month_num is not None and "month_num" in frame.columns:
        frame = frame.filter(pl.col("month_num") == month_num)
    return frame


def ensure_frame(df: pl.DataFrame | None) -> pl.DataFrame:
    return df if df is not None else pl.DataFrame()


def normalized_contains_expr(column: str, needle: str) -> pl.Expr:
    return pl.col(column).fill_null("").map_elements(
        lambda value: normalize_key(value).find(normalize_key(needle)) >= 0,
        return_dtype=pl.Boolean,
    )


def normalized_equals_expr(column: str, value: str) -> pl.Expr:
    normalized = normalize_key(value)
    return pl.col(column).fill_null("").map_elements(
        lambda cell: normalize_key(cell) == normalized,
        return_dtype=pl.Boolean,
    )


def normalize_key(value: Any) -> str:
    text = str(value or "").lower().replace("speciality", "specialty")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def read_frame(path: Path | None) -> pl.DataFrame | None:
    if path is None:
        return None
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pl.read_csv(path)
    if suffix == ".json":
        return pl.read_json(path)
    if suffix == ".parquet":
        return pl.read_parquet(path)
    raise ValueError(f"Unsupported table format for {path}")


def write_values_json(result: CloseValuesResult, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result.values, indent=2) + "\n", encoding="utf-8")
