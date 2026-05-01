# Consolidated P&L Writer

Module:
- `hgf_pnl.writers.consolidated_pnl`

CLI:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/HGF_CONSOLIDATED_MARCH_2026_GENERATED.xlsx \
  --values tmp/consolidated_values.json
```

Create an editable writer config:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --init-config configs/consolidated_pnl_writer.json
```

Generate an example values file from a completed workbook:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --values tmp/consolidated_values_from_final.json \
  --example-values-from "sample_files/P&L_S/FULL COMPANY P&L_s/HGF CONSOLIDATED_MARCH 2026_FINAL.xlsx"
```

## Role In The Flow

The writer expects extraction, review, and override resolution to happen before it runs. It does not inspect source workpapers directly.

Inputs:
- template workbook
- writer config
- approved values JSON

Output:
- generated consolidated P&L workbook
- write metadata
- validation metadata

## Write Surface

The writer targets the hidden raw-data tabs:

- `RAW DATA_Master File`
- `RAW DATA_COGS & Freight`
- `RAW DATA_Payroll`

It preserves the visible `MARCH 2026 FULL ` formulas and styles. One known exception is `MARCH 2026 FULL !EB48`, the hidden source-total cell for Employee Benefits, because the template does not stage that BR Info override in a raw-data tab.

The writer also sets workbook calculation flags so Excel/LibreOffice should recalculate formulas when opened or saved.

## Config Shape

Cell mappings use semantic `source_key` values:

```json
{
  "sheet_name": "RAW DATA_Master File",
  "cell": "B72",
  "source_key": "raw_master.sales.dtc",
  "required": false,
  "value_type": "number"
}
```

The values JSON can be nested:

```json
{
  "raw_master": {
    "sales": {
      "dtc": 163408.05
    }
  },
  "raw_payroll": {
    "production": 64635.26
  }
}
```

or flat:

```json
{
  "raw_master.sales.dtc": 163408.05,
  "raw_payroll.production": 64635.26
}
```

## Formula Safety

By default, the writer will not overwrite an existing formula cell. This protects raw-tab subtotal formulas like:

- `RAW DATA_Master File!B50`
- `RAW DATA_Master File!B57`
- `RAW DATA_Master File!B63`
- `RAW DATA_Master File!B75`
- `RAW DATA_Master File!B84`
- `RAW DATA_Master File!B112`
- `RAW DATA_Payroll!B26`

If a reviewed config intentionally needs to replace a formula, set `overwrite_formula` on that specific cell write.

The March 2026 payroll allocation layout intentionally replaces the template's old Lital allocation total row:

- `RAW DATA_Payroll!A24` is set to `CORP`
- `RAW DATA_Payroll!B24` receives `raw_payroll.lital_allocation.corp`
- `RAW DATA_Payroll!B25` becomes `=SUM(B21:B24)`
- `RAW DATA_Payroll!B26` is cleared

The writer can also refresh the visible `Payroll - Art` and `Payroll- IT` actual formulas on `MARCH 2026 FULL `. These formulas intentionally overwrite the stale template constants when `raw_payroll.allocation_breakdowns` values are present. If no breakdown values are present, those cells are skipped and the template formulas remain unchanged.

Expected value keys:

```json
{
  "raw_payroll": {
    "allocation_breakdowns": {
      "art": {
        "trend_house": 13007.702,
        "og_specialty_usa": 0,
        "online_lux": 0,
        "online": 5857.694,
        "dtc": 8173.08,
        "all_pop_art": 0,
        "ink": 0,
        "general": 5923.084
      },
      "it": {
        "trend_house": 0,
        "og_specialty_usa": 0,
        "online_lux": 0,
        "online": 1230.768,
        "dtc": 1230.768,
        "all_pop_art": 0,
        "ink": 0,
        "general": 13384.614
      }
    }
  }
}
```

## Validation

Config can include validations that compare written cells to the approved values before recalculation:

```json
{
  "name": "raw master DTC sales",
  "sheet_name": "RAW DATA_Master File",
  "cell": "B72",
  "expected_source_key": "raw_master.sales.dtc",
  "tolerance": 0.01
}
```

Later output validation should add formula-result checks after LibreOffice or Excel recalculation, especially visible report totals and hidden check columns.
