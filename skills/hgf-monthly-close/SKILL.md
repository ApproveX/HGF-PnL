---
name: hgf-monthly-close
description: Use for HGF monthly close packages: discover files, configure and run extractors, review audit warnings and overrides, generate consolidated P&L workbooks, and validate output artifacts.
---

# HGF Monthly Close

This skill operates the HGF P&L close tooling that ships with this plugin. Use it whenever the user asks to process, inspect, validate, or generate an HGF monthly P&L package.

The Python implementation lives alongside this skill:

```text
skills/hgf-monthly-close/
├── SKILL.md            ← this file
├── hgf_pnl/            ← Python package (extractors, writers, pipeline, formulas)
└── scripts/            ← CLI entry points
```

The plugin source itself is read-only on most installs, so the Python venv lives at a separate, user-writable location. The bootstrap command sets this up; the skill expects:

| What | Path |
|---|---|
| Python interpreter (`$HGF_PY`) | `$HOME/.local/share/hgf-pnl/venv/bin/python` |
| Plugin source root (`$HGF_ROOT`) | contents of `$HOME/.local/share/hgf-pnl/plugin_root` |
| Run output directory | `tmp/runs/<period-slug>/` under the user's current working directory (writable) |

In every example below, treat `$HGF_PY` and `$HGF_ROOT` as placeholders. Because shell variables do not persist across separate Bash tool calls, define both at the top of each call, or substitute the resolved literal paths:

```bash
HGF_PY="$HOME/.local/share/hgf-pnl/venv/bin/python"
HGF_ROOT="$(cat $HOME/.local/share/hgf-pnl/plugin_root)"
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/discover_package.py" --help
```

## Core Rule

Treat source workpapers as immutable. Do not edit files under the client-provided package unless the user explicitly asks. Generated configs, manifests, extracted tables, values JSON, and output workbooks go into a run directory:

```text
tmp/runs/<period-slug>/
```

For client-facing generated workbooks, use clear names such as:

```text
tmp/runs/march-2026/HGF_CONSOLIDATED_MARCH_2026_GENERATED.xlsx
```

## Narrate Progress While Working

The pipeline takes minutes end-to-end and produces signals the user needs to see in real time. Do not go silent. As you work, send short updates — one or two sentences — that surface:

- the step you are about to run, in plain language
- what you found that mattered (totals, missing inputs, low-confidence classifications, color-coded review rows)
- anything surprising, including totals that disagree, unsupported formulas, or files outside the expected layout
- judgement calls you are about to make, and why
- what is next

Updates should describe findings, not mechanics. Prefer:

> "Payroll extraction finished. Employee gross pay matches the Distribution tab at $241,161.25. The IT allocation summary differs by $0.04 — looks like a rounding artifact, but flagging it."

over:

> "Ran extract_payroll_journal.py. Got CSVs."

If you hit a warning that requires the user's judgement (override approval, unknown-charge magenta row, mismatched declared total), pause and ask before continuing.

## Setup

If either of these is missing, run `/hgf-pnl-bootstrap` before doing anything else:

- `$HOME/.local/share/hgf-pnl/venv/bin/python`
- `$HOME/.local/share/hgf-pnl/plugin_root`

A quick check:

```bash
test -x "$HOME/.local/share/hgf-pnl/venv/bin/python" \
  && test -f "$HOME/.local/share/hgf-pnl/plugin_root" \
  && "$HOME/.local/share/hgf-pnl/venv/bin/python" -c "import hgf_pnl" \
  && echo OK
```

If anything fails, run `/hgf-pnl-bootstrap`. That command creates the external venv, installs `hgf_pnl` as a wheel (the plugin source is read-only, so editable installs do not work), records the plugin source path, smoke-tests a script, and reports whether LibreOffice is available for workbook recalculation. Do not invent a manual setup path; let the bootstrap command handle it so failures are surfaced consistently.

When the plugin updates, re-run `/hgf-pnl-bootstrap` to pick up the new `hgf_pnl` code in the venv.

## Workflow

1. Discover the package.
2. Review file classifications and missing expected inputs.
3. Configure extractors as needed.
4. Run extractors and write normalized outputs.
5. Validate extraction totals, formulas, warnings, and manual-review items.
6. Resolve overrides into an approved values JSON.
7. Run the consolidated P&L writer.
8. Validate the generated workbook.
9. Update the run manifest after each major step.

Do not skip the manifest. It is the audit trail for what the agent believed, used, ignored, changed, and produced.

## Step 1 — Discover The Package

Always run discovery first. It scans the close-package folder, classifies every file, and produces an initial manifest.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/discover_package.py" \
  "sample_files/Workpapers MARCH" \
  --manifest-output tmp/runs/march-2026/run_manifest.json
```

Use `--inspect-workbooks` when classification is ambiguous or when sheet names will help downstream config:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/discover_package.py" \
  "sample_files/Workpapers MARCH" \
  --inspect-workbooks \
  --discovery-output tmp/runs/march-2026/discovery.json \
  --manifest-output tmp/runs/march-2026/run_manifest.json
```

Review:
- `source_input` files with extractor matches.
- `instruction` files, especially addback PDFs and email-thread PDFs.
- `supporting_input` files that may need manual interpretation.
- `deliverable_or_prior_output` files.
- missing expected inputs, e.g. BR Info or Payroll when they live outside the selected folder.

Treat discovery classifications as suggestions, not facts. For low-confidence or ambiguous files, inspect the file and either update the manifest classification or mark the input unselected with a reason.

## Step 2 — Run Extractors

Use the configured extractors instead of writing one-off parsing code. Most Excel extractors support:

- `--config <path>`
- `--init-config <path>`
- `--no-calculate-formulas`

Formula evaluation is on by default for Excel extractors. Unsupported formulas are preserved as warnings/status values, not silently coerced. The supported subset and known unsupported areas are documented in `$HGF_ROOT/skills/hgf-monthly-close/docs/extractors.md`.

When a file is natural-language-heavy or accountant-reviewed, inspect it before running strict extraction. For chargeback PDFs, inspect text/table structure before choosing extraction settings. For addbacks, read the PDF/email instructions first, then configure the reviewed GL workbook extractor. For yellow/red/magenta reviewed GL rows, preserve both source color and semantic comments.

When changing extractor config, write the config to the run directory and record the config path in the manifest.

### `extract_pl_by_dept.py` — Profit and Loss By Department

Matrix P&L unpivot for `Workpapers <MONTH>/DATA/Profit and Loss By Dept.xlsx`. Header row 5; data rows roughly 6 through 52.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_pl_by_dept.py" \
  "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx" \
  --output tmp/runs/march-2026/pl_by_dept.csv \
  --format csv
```

Use `--no-totals` to drop rollup columns. Configurable keys: `sheet_name`, `header_keywords`, `line_item_keywords`, `include_total_columns`, `total_column_patterns`, `skip_line_patterns`, `section_patterns`, `preserve_zero_amounts`. Output is long format with `line_item`, `section`, `department`, `amount`, plus formula/calculation status fields.

### `extract_th_revenue.py` — Trend House Revenue Report

Three sheets — `Summary`, `Details`, `USA Stock`. The March sample's `Details` sheet already includes the USA Stock rows; treat `usa_stock` as a supporting subset, not additive to `po_details` revenue.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_th_revenue.py" \
  "sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx" \
  --output-dir tmp/runs/march-2026/th_revenue \
  --format csv
```

Smoke-check expected for March 2026: summary and details non-total revenue both `1,242,393.40`; non-total total cost `750,574.50`.

### `extract_payroll_journal.py` — Payroll Journal

Two sheets: `Payroll` (source of truth) and `Payroll Distribution` (intermediary copy that may have stale formulas). By default, distribution output is derived from the `Payroll` sheet. Only use `--use-distribution-sheet` when you have reviewed the workbook and intentionally want the intermediary parsed.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_payroll_journal.py" \
  "sample_files/PAYROLL/Payroll Journal_March 2026.xlsx" \
  --output-dir tmp/runs/march-2026/payroll_journal \
  --format csv
```

Output tables: `employees`, `allocations`, `allocation_summaries`, `distribution`. The `allocation_summaries` table is the preferred source for `Payroll - Art` and `Payroll- IT` actuals on `MARCH 2026 FULL `. March 2026 spot checks: 43 employee rows, gross pay total `241,161.25`, Art `32,961.56`, IT `15,846.15`, Production `64,635.26`, Corp `64,655.38`.

Critical writer mapping from `payroll_allocation_summaries.csv`:

| summary filter | writer key |
|---|---|
| `department == "Art"` and `allocation_category == "TH"` | `raw_payroll.allocation_breakdowns.art.trend_house` |
| `department == "Art"` and `allocation_category == "B&M USA"` | `raw_payroll.allocation_breakdowns.art.og_specialty_usa` |
| `department == "Art"` and `allocation_category == "Online Lux"` | `raw_payroll.allocation_breakdowns.art.online_lux` |
| `department == "Art"` and `allocation_category == "Online"` | `raw_payroll.allocation_breakdowns.art.online` |
| `department == "Art"` and `allocation_category == "OG DTC"` | `raw_payroll.allocation_breakdowns.art.dtc` |
| `department == "Art"` and `allocation_category == "APA"` | `raw_payroll.allocation_breakdowns.art.all_pop_art` |
| `department == "Art"` and `allocation_category == "General"` | `raw_payroll.allocation_breakdowns.art.general` |
| `department == "Art"` and `allocation_category == "Total"` | `raw_payroll.allocation_breakdowns.art.total` |
| `department == "IT"` and `allocation_category == "Online"` | `raw_payroll.allocation_breakdowns.it.online` |
| `department == "IT"` and `allocation_category == "OG DTC"` | `raw_payroll.allocation_breakdowns.it.dtc` |
| `department == "IT"` and `allocation_category == "General"` | `raw_payroll.allocation_breakdowns.it.general` |
| `department == "IT"` and `allocation_category == "Total"` | `raw_payroll.allocation_breakdowns.it.total` |

If a direct Art/IT category is absent from the summary table for a month, set that writer key to `0` rather than omitting the whole `allocation_breakdowns` object. The writer intentionally skips all Art/IT formula writes when none of the formula source values are present.

The March 2026 `Lital Allocation in G&A Exp` block on `Payroll Distribution` has formulas pointing at the wrong source rows for TH and CORP. The source rows on `Payroll!M57` and `Payroll!M59` match the cached values.

### `extract_br_info.py` — Manual Override Table

`BR Info.xlsx` is accountant-entered overrides keyed by month. Rows go to long format: `year`, `month_num`, `month_name`, `override_name`, `value`. The current March 2026 sample has 9 override rows totaling `578,002.00` on the `2026` sheet, including Online Sales. Do not rely on older notes that say 8 rows or `35,197.00`.

Resolve March BR Info rows into writer values using these known labels:

| BR Info override label | writer key |
|---|---|
| `Online Sales` | `raw_master.sales.online` |
| `AllPopArt Sales` | `raw_master.sales.apa` |
| `AllPopArt Returns and Allowances` | `raw_master.returns.apa` |
| `Employee Benefits` | `full_report.source_totals.employee_benefits` |
| `Equipment Leasing` | `raw_master.gl.equipment_lease_adjustment` |
| `Bank Fees` | `raw_master.gl.bank_fees_adjustment` |
| `Merchant Account Fees` | `raw_master.gl.merchant_account_fees_adjustment` |
| `License & Tax` | `raw_master.gl.licenses_taxes_permits` |
| `LOC Interest` | `raw_master.gl.loc_interest` |

Review every BR Info row before writing. If a label is new or ambiguous, add it to the manifest as `needs_review` instead of guessing.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_br_info.py" \
  "sample_files/BR Info.xlsx" \
  --output tmp/runs/march-2026/br_info.csv \
  --format csv
```

### `extract_monthly_revenue.py` — DTC & WS Monthly Revenue

Four sheets — summary, Shopify orders, refunds, coupons.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_monthly_revenue.py" \
  "sample_files/Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx" \
  --output-dir tmp/runs/march-2026/monthly_revenue \
  --format csv
```

March 2026 spot checks: Shopify net sales `163,408.05` (DTC `146,599.00` + WS `16,809.05`), refunds `10,114.56`, coupons `17,212.05`.

### `extract_division_cogs.py` — Division COGS

Year matrix tabs (`2018`–`2026`) and partner detail tabs (`YYYY Partner Details`). Year matrices are forward-filled by month because many rows use merged month cells. Partner detail tabs unpivot from repeated month groups of `COGS`, `Material Cost`, `Labor Cost`. `VLOOKUP` formulas with workbook-index/whole-column references are reported unsupported rather than evaluated.

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_division_cogs.py" \
  "sample_files/Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx" \
  --output-dir tmp/runs/march-2026/division_cogs \
  --format csv
```

March 2026 spot checks: 2026 March COGS total column `213,273.23`; Online - USA COGS `181,629.67`; D2C partner detail COGS `19,658.51` (material `15,596.90`, labor `4,061.61`).

### `extract_addbacks_gl.py` — Reviewed GL With Addbacks

This is the agentic step. Read the email/PDF instructions first (`March Addbacks_$23,195.pdf` for the March sample), then configure the reviewed GL extractor. Default row groups encode the March email semantics: `addbacks` (rows where `Comments` is `Addback`), `red_addback_color_rows` (red/pink fill, validation only), `unknown_charges` (magenta), `account_department_edits` (yellow), `other_review_rows` (blue, undescribed).

Run with the declared total parsed from the email:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_addbacks_gl.py" \
  "sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx" \
  --declared-addbacks-total 23195 \
  --output-dir tmp/runs/march-2026/addbacks_gl \
  --format csv
```

For the March sample, the comment-based addback total `23,195.16` is the authoritative match to the email total. The red/pink row total `23,248.67` is `53.51` higher because one red/pink row is not marked `Addback`; keep that as supporting evidence rather than authority. Configurable: `sheet_name`, `header_aliases`, `row_group_rules`, `declared_totals`.

### `extract_chargeback_pdf.py` — Chargeback Report PDF

The PDF is an email export. `pdfplumber` table extraction loses some monthly summary values, so the extractor parses line text and preserves source page/line references. Always profile first:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/profile_chargeback_pdf.py" \
  "sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf" \
  --output-dir tmp/runs/march-2026/chargeback_pdf_profile
```

Inspect `chargeback_pdf_profile.md` and `chargeback_pdf_raw_text.txt`, edit `chargeback_pdf_suggested_config.json` if anchors/category order/month/year need adjustment, then run:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/extract_chargeback_pdf.py" \
  "sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf" \
  --config tmp/runs/march-2026/chargeback_pdf_profile/chargeback_pdf_suggested_config.json \
  --output-dir tmp/runs/march-2026/chargeback_pdf \
  --format csv
```

March 2026 numbers: monthly grand total `-104,205`; allowance `-65,092`; penalty `-7,821`; Amazon holdback provision `-15,661`; return `-15,596`; software fees `-34`; customer-detail grand total `-88,543`. Customer-detail lines are rounded to whole dollars; the sum of individual rows `-88,541` differs from the provided `Grand Total` row `-88,543` by rounding. Treat provided totals as authoritative for the PDF block and retain source lines for audit.

## Step 3 — Review Items

Before writing the consolidated workbook, summarize for the user:

- unsupported formulas
- mismatched totals
- missing expected files
- duplicate candidates for one extractor
- addback rows and totals
- yellow-row expected account/department edits
- magenta unknown-charge rows
- BR Info manual overrides
- chargeback totals and source-page evidence
- writer-cell values that require manual judgement

For each approved override, record:
- target key or cell
- original value
- override value
- reason
- source
- approver, if known

## Step 4 — Consolidated Values JSON

The writer reads an approved values JSON. Keys can be nested or flat:

```json
{
  "raw_master": { "sales": { "dtc": 163408.05 } },
  "raw_payroll": { "production": 64635.26 }
}
```

```json
{
  "raw_master.sales.dtc": 163408.05,
  "raw_payroll.production": 64635.26
}
```

Key namespaces include:
- `raw_master.*` — Master File raw-data tab
- `raw_cogs.*` — COGS & Freight raw-data tab
- `raw_payroll.*` — Payroll raw-data tab, plus `raw_payroll.allocation_breakdowns.{art,it}.*` for the visible Art/IT actual rows on `MARCH 2026 FULL `

The writer does not build this JSON from extractor outputs. The agent must explicitly transform reviewed extractor rows into these keys. In particular, do not stop after writing `payroll_allocation_summaries.csv` or `br_info.csv`; map those rows into `consolidated_values.json` before running the writer.

## Step 5 — Run The Writer

The writer targets the hidden raw-data tabs and preserves the visible `MARCH 2026 FULL ` formulas and styles:

```text
RAW DATA_Master File
RAW DATA_COGS & Freight
RAW DATA_Payroll
```

Known exception: `MARCH 2026 FULL !EB48` is the hidden source-total cell for Employee Benefits. The March template does not stage this BR Info override in a raw-data tab, so the writer config may write that hidden source cell directly.

By default the writer will not overwrite an existing formula cell. Protected raw-tab subtotal formulas include:
- `RAW DATA_Master File!B50, B57, B63, B75, B84, B112`
- `RAW DATA_Payroll!B26`

If a reviewed config intentionally needs to replace a formula, set `overwrite_formula` on that specific cell write. The March 2026 payroll allocation layout deliberately replaces the template's old Lital allocation total row: `RAW DATA_Payroll!A24` becomes `CORP`, `B24` receives `raw_payroll.lital_allocation.corp`, `B25` becomes `=SUM(B21:B24)`, and `B26` is cleared.

The writer can refresh the visible `Payroll - Art` and `Payroll- IT` actual formulas on `MARCH 2026 FULL ` when `raw_payroll.allocation_breakdowns` values are present. If no breakdown values are present, those cells are skipped and the template formulas remain unchanged.

Generate an editable writer config:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/write_consolidated_pnl.py" \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --init-config tmp/runs/march-2026/consolidated_writer_config.json
```

Write the workbook from approved values:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/write_consolidated_pnl.py" \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/runs/march-2026/HGF_CONSOLIDATED_MARCH_2026_GENERATED.xlsx \
  --values tmp/runs/march-2026/consolidated_values.json \
  --config tmp/runs/march-2026/consolidated_writer_config.json
```

For sample round-trip testing only, values can be extracted from a completed workbook:

```bash
"$HGF_PY" "$HGF_ROOT/skills/hgf-monthly-close/scripts/write_consolidated_pnl.py" \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --values tmp/runs/march-2026/consolidated_values_from_final.json \
  --example-values-from "sample_files/P&L_S/FULL COMPANY P&L_s/HGF CONSOLIDATED_MARCH 2026_FINAL.xlsx"
```

Do not use a completed workbook as source values for production unless the user explicitly wants a template/writer validation exercise.

## Step 6 — Recalculate

The writer sets workbook recalculation flags. If LibreOffice is available, use it to recalculate cached formula values and re-save the workbook. If not, clearly tell the user that Excel/LibreOffice should recalculate on open. The bootstrap command reports availability; if you skip bootstrap, check directly:

```bash
which libreoffice
which soffice
```

## Step 7 — Output Validation

Before saying an output workbook is ready:

1. Confirm the writer completed with no validation failures.
2. Compare mapped raw-tab cells to approved values.
3. Check skipped cells and explain why they were skipped.
4. Check workbook recalculation flags.
5. If recalculation is available, inspect visible totals and formula errors after recalculation.
6. Record validation results in the manifest.

Known workbook-specific concerns:
- `MARCH 2026 FULL ` is formula-heavy and depends on the raw-data tabs.
- The `Payroll` tab is the payroll source of truth. The `Payroll Distribution` tab is an intermediary copy and may contain stale formulas, especially in the March 2026 `Lital Allocation in G&A Exp` block.
- Some hidden legacy columns reference `RAW DATA_Master File!B134:B151`, while the template raw tab currently ends at row 133. Flag these as stale hidden references unless the client says otherwise.

## Manifest Discipline

Update the manifest after each of:

- extractor config chosen
- extractor run completed
- warning reviewed
- override approved
- writer run completed
- output validation completed

Status transitions: `discovered` → `configured` → `extracted` → `reviewed` → `written` → `validated`. Use `needs_review` or `failed` when applicable.

## Communication Template

When reporting status, separate:

- extraction/configuration facts
- validation warnings
- assumptions
- files produced
- actions still requiring human review

Never present a generated workbook as client-ready if formula recalculation or review warnings remain unresolved.

## Reference: Scripts → Modules

| Script | Module | Notes |
|---|---|---|
| `discover_package.py` | `hgf_pnl.pipeline.discovery`, `hgf_pnl.pipeline.manifest` | Always run first |
| `extract_pl_by_dept.py` | `hgf_pnl.extractors.pl_by_dept` | Matrix P&L unpivot |
| `extract_th_revenue.py` | `hgf_pnl.extractors.th_revenue` | Treat `USA Stock` as subset |
| `extract_payroll_journal.py` | `hgf_pnl.extractors.payroll_journal` | `Payroll` is source of truth |
| `extract_br_info.py` | `hgf_pnl.extractors.br_info` | Wide-to-long override table |
| `extract_monthly_revenue.py` | `hgf_pnl.extractors.monthly_revenue` | DTC/WS Shopify, refunds, coupons |
| `extract_division_cogs.py` | `hgf_pnl.extractors.division_cogs` | Year matrix + partner details |
| `extract_addbacks_gl.py` | `hgf_pnl.extractors.addbacks_gl` | Read email/PDF first |
| `profile_chargeback_pdf.py` | `hgf_pnl.extractors.chargeback_pdf` | Profile before extracting |
| `extract_chargeback_pdf.py` | `hgf_pnl.extractors.chargeback_pdf` | Run with profiled config |
| `write_consolidated_pnl.py` | `hgf_pnl.writers.consolidated_pnl` | Writes hidden raw-data tabs only |

Detailed module documentation lives in `$HGF_ROOT/skills/hgf-monthly-close/docs/pipeline.md`, `$HGF_ROOT/skills/hgf-monthly-close/docs/extractors.md`, `$HGF_ROOT/skills/hgf-monthly-close/docs/consolidated_writer.md`, and `$HGF_ROOT/skills/hgf-monthly-close/docs/priority_file_exploration.md`.
