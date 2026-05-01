# Extractors

## Profit and Loss By Department

Module:
- `hgf_pnl.extractors.pl_by_dept`

CLI:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx"
```

By default, Excel extractors evaluate supported formulas in memory and use calculated values for `amount`/numeric output fields. Use `--no-calculate-formulas` to inspect workbook cached values instead.

Write CSV:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py \
  "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx" \
  --output tmp/pl_by_dept.csv \
  --format csv
```

Exclude rollup/total columns:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py \
  "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx" \
  --no-totals
```

Disable formula evaluation:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py \
  "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx" \
  --no-totals \
  --no-calculate-formulas
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py --init-config configs/pl_by_dept.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_pl_by_dept.py \
  "sample_files/Workpapers MARCH/DATA/Profit and Loss By Dept.xlsx" \
  --config configs/pl_by_dept.json
```

### Heuristic Behavior

The extractor does not require exact row numbers or exact department names.

It detects:
- the target sheet from sheet-name keywords
- the department header row from dense text/header keyword scoring
- the line-item column from line-item keyword scoring
- department columns from non-empty headers with numeric/formula-like values below
- report title and period from rows above the header

### Agent-Tunable Config

The JSON config can adjust:
- `sheet_name` or `sheet_name_keywords`
- `header_keywords`
- `line_item_keywords`
- `max_header_scan_rows`
- `include_total_columns`
- `total_column_patterns`
- `skip_line_patterns`
- `section_patterns`
- `stop_after_blank_rows`
- `preserve_zero_amounts`
- `calculate_formulas`
- `use_calculated_formula_values`

### Output Shape

Rows are normalized to long format:

- `source_file`
- `sheet`
- `row`
- `line_item`
- `section`
- `department`
- `amount`
- `is_total_column`
- `is_section_row`
- `line_item_cell`
- `amount_cell`
- `formula`
- `fill_color`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

The sample workbook has stale/cached formula values that are mostly zero, so the extractor preserves cached values and formulas. Formula evaluation is enabled by default, and `amount` uses the calculated result when the formula is supported.

## Formula Evaluation

Module:
- `hgf_pnl.formulas`

Current supported formula subset:
- numeric literals
- string literals
- unary `+` and `-`
- arithmetic `+`, `-`, `*`, `/`, `^`
- percent suffix, such as `10%`
- parentheses
- same-sheet cell references, such as `B10`
- quoted or unquoted same-workbook sheet references, such as `'RAW DATA_Master File'!B107`
- workbook-index-prefixed same-workbook sheet references, such as `[1]Payroll!M56`
- ranges, such as `B7:B9`
- aggregate functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`

Unsupported formulas return a `FormulaSentinel` with status `unsupported`, instead of being silently coerced. This is intentional so ingestion can proceed while preserving the formulas that need new parser support.

Known unsupported areas:
- `SUMIF`, `VLOOKUP`, `XLOOKUP`
- `IF` and comparison expressions
- external workbook references
- array formulas and pivot/table formulas
- Excel-specific error semantics beyond simple division by zero

Formula engine future plans:
- Add support for common conditional functions, especially `IF`.
- Add criteria functions such as `SUMIF`, `SUMIFS`, `COUNTIF`, and `COUNTIFS`.
- Add lookup support for `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, and `INDEX`/`MATCH`.
- Add comparison operators and boolean coercion compatible with Excel behavior.
- Add better Excel error propagation for values such as `#DIV/0!`, `#N/A`, and `#VALUE!`.
- Add a formula audit report that summarizes unsupported functions and references by workbook/sheet/cell.

## P&L By Department Future Plans

- Add reconciliation checks, such as verifying department totals equal row totals.
- Add a richer confidence report for detected header rows, line-item columns, and department columns.
- Decide whether total/rollup columns should be excluded by default for downstream calculations.
- Test against another month or another client export once one is available.

## Payroll Journal

Module:
- `hgf_pnl.extractors.payroll_journal`

CLI:

```bash
.venv/bin/python scripts/extract_payroll_journal.py \
  "sample_files/PAYROLL/Payroll Journal_March 2026.xlsx"
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_payroll_journal.py \
  "sample_files/PAYROLL/Payroll Journal_March 2026.xlsx" \
  --output-dir tmp/payroll_journal \
  --format csv
```

Disable formula evaluation:

```bash
.venv/bin/python scripts/extract_payroll_journal.py \
  "sample_files/PAYROLL/Payroll Journal_March 2026.xlsx" \
  --no-calculate-formulas
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_payroll_journal.py --init-config configs/payroll_journal.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_payroll_journal.py \
  "sample_files/PAYROLL/Payroll Journal_March 2026.xlsx" \
  --config configs/payroll_journal.json
```

### Heuristic Behavior

The extractor detects:
- the `Payroll` sheet by sheet-name keyword, preferring an exact `Payroll` match.
- the `Payroll Distribution` sheet by distribution keywords.
- employee rows from the configured code/name/gross-pay columns.
- department sections from the section label/total columns on the last employee row in each section.
- allocation tables from nearby header rows containing allocation labels and a `Total` column.
- allocation total rows immediately below those tables, preserving department/category/source-cell evidence.
- payroll distribution blocks from one-column block anchors such as `Payroll Sales`, `Payroll Corp`, and `Lital Allocation in G&A Exp`.

### Output Tables

The extractor exposes:
- `employees`: one row per employee gross-pay line.
- `allocations`: one row per non-zero employee allocation target.
- `allocation_summaries`: one row per non-zero department/category allocation total row cell.
- `distribution`: one row per payroll distribution summary line.

### March 2026 Sample Validation

For the sample workbook:
- Employee rows: `43`
- Employee gross-pay total: `241,161.25`
- Production: `64,635.26`
- Sales Dept: `44,909.04`
- IT: `15,846.15`
- Art: `32,961.56`
- Corp: `64,655.38`
- Ops: `18,153.86`
- Non-zero employee allocation rows: `24`
- Non-zero allocation summary rows: `13`

The `allocation_summaries` table is the preferred source for report rows that need department-level allocation detail, such as `Payroll - Art` and `Payroll- IT` on `MARCH 2026 FULL `. For March 2026:

- Art comes from `Payroll!G47:N47`: `OG DTC`, `Online`, `TH`, `General`, and `Total`.
- IT comes from `Payroll!G38:J38`: `OG DTC`, `Online`, `General`, and `Total`.
- `General` should be allocated to the P&L actual columns using the visible revenue-share cells for the target column.

When preparing writer values, map `allocation_summaries` rows into `raw_payroll.allocation_breakdowns`:

| department | allocation_category | writer key |
|---|---|---|
| `Art` | `TH` | `raw_payroll.allocation_breakdowns.art.trend_house` |
| `Art` | `B&M USA` | `raw_payroll.allocation_breakdowns.art.og_specialty_usa` |
| `Art` | `Online Lux` | `raw_payroll.allocation_breakdowns.art.online_lux` |
| `Art` | `Online` | `raw_payroll.allocation_breakdowns.art.online` |
| `Art` | `OG DTC` | `raw_payroll.allocation_breakdowns.art.dtc` |
| `Art` | `APA` | `raw_payroll.allocation_breakdowns.art.all_pop_art` |
| `Art` | `General` | `raw_payroll.allocation_breakdowns.art.general` |
| `Art` | `Total` | `raw_payroll.allocation_breakdowns.art.total` |
| `IT` | `Online` | `raw_payroll.allocation_breakdowns.it.online` |
| `IT` | `OG DTC` | `raw_payroll.allocation_breakdowns.it.dtc` |
| `IT` | `General` | `raw_payroll.allocation_breakdowns.it.general` |
| `IT` | `Total` | `raw_payroll.allocation_breakdowns.it.total` |

Use `0` for absent direct categories so the writer can still refresh every Art/IT actual formula.

The `Payroll Distribution` sheet is treated as an intermediary copy format. By default, the extractor derives distribution output directly from the `Payroll` sheet, which is the source of truth. This matters for the March 2026 `Lital Allocation in G&A Exp` block: the intermediary sheet has formulas pointing at the wrong source rows for TH and CORP, while the source formulas on `Payroll!M57` and `Payroll!M59` match the cached values.

Use `--use-distribution-sheet` only when an agent has reviewed the workbook and intentionally wants the intermediary sheet parsed.

### Output Shape

Employee rows:
- `source_file`
- `sheet`
- `row`
- `employee_code`
- `employee_name`
- `gross_pay`
- `section`
- `section_total_amount`
- `section_total_source_row`
- `allocated_total`
- `allocation_difference`
- `formula`
- `cached_gross_pay`
- `calculated_gross_pay`
- `calculation_status`
- `calculation_detail`

Allocation rows:
- `source_file`
- `sheet`
- `header_row`
- `row`
- `employee_code`
- `employee_name`
- `section`
- `allocation_category`
- `amount`
- `gross_pay`
- `percent_of_gross`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

Distribution rows:
- `source_file`
- `sheet`
- `block`
- `row`
- `label`
- `amount`
- `is_total_row`
- `is_check_row`
- `is_difference_row`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

## BR Info

Module:
- `hgf_pnl.extractors.br_info`

CLI:

```bash
.venv/bin/python scripts/extract_br_info.py "sample_files/BR Info.xlsx"
```

Write CSV:

```bash
.venv/bin/python scripts/extract_br_info.py \
  "sample_files/BR Info.xlsx" \
  --output tmp/br_info.csv \
  --format csv
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_br_info.py --init-config configs/br_info.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_br_info.py \
  "sample_files/BR Info.xlsx" \
  --config configs/br_info.json
```

### Heuristic Behavior

This workbook is treated as accountant-entered manual overrides. The extractor detects:
- the target sheet, defaulting to the first sheet unless configured.
- the year from the sheet name or cells above the month header.
- the month header row by scanning for month names.
- override labels from the configured label column.
- populated month/override cells as long-format rows.

Formula evaluation is still available by default for consistency with other Excel extractors, but the March 2026 sample contains no formulas.

### March 2026 Sample Validation

For the sample workbook:
- Sheet: `2026`
- Header row: `2`
- Parsed override rows: `9`
- Populated month: `March`
- Value total: `578,002.00`

Known March labels and writer targets:

| override_name | writer key |
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

### Output Shape

Rows are normalized to:
- `source_file`
- `sheet`
- `year`
- `month_num`
- `month_name`
- `override_name`
- `value`
- `raw_value`
- `source_cell`
- `formula`
- `cached_value`
- `calculated_value`
- `calculation_status`
- `calculation_detail`

## Addbacks Reviewed GL

Module:
- `hgf_pnl.extractors.addbacks_gl`

CLI:

```bash
.venv/bin/python scripts/extract_addbacks_gl.py \
  "sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx"
```

Run with the total parsed from the email/PDF instructions:

```bash
.venv/bin/python scripts/extract_addbacks_gl.py \
  "sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx" \
  --declared-addbacks-total 23195
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_addbacks_gl.py \
  "sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx" \
  --output-dir tmp/addbacks_gl \
  --format csv
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_addbacks_gl.py --init-config configs/addbacks_gl.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_addbacks_gl.py \
  "sample_files/Workpapers MARCH/HGF GL_March_Sent April 13_DONE.xlsx" \
  --config configs/addbacks_gl.json
```

### Agentic Configuration

This extractor is designed for a natural-language instruction step before workbook ingestion. The agent can read an email/PDF thread, then populate config fields such as:

- `sheet_name` or `sheet_name_keywords`
- `header_aliases`
- `row_group_rules`
- `declared_totals`
- `calculate_formulas`
- `use_calculated_formula_values`

The default row groups encode the March email semantics:

- `addbacks`: rows where `Comments` is `Addback`
- `red_addback_color_rows`: rows with red/pink fill, for validation
- `unknown_charges`: rows with magenta fill
- `account_department_edits`: rows with yellow fill
- `other_review_rows`: blue rows not described by the email

For the March sample, the comment-based addback total is the authoritative match to the email total. Red/pink fill is kept as supporting evidence because one red/pink row is not marked `Addback`.

Example config fragment from the March email:

```json
{
  "declared_totals": [
    {
      "group_name": "addbacks",
      "amount": 23195,
      "amount_column": "amount",
      "tolerance": 1
    }
  ],
  "row_group_rules": [
    {
      "name": "addbacks",
      "description": "Rows explicitly marked as addbacks in the accountant comments.",
      "match_mode": "any",
      "fill_colors": [],
      "comment_patterns": ["^addback$"],
      "comment_column": "comments",
      "nonblank_columns": [],
      "blank_columns": []
    },
    {
      "name": "account_department_edits",
      "description": "Yellow rows where columns J/K contain expected account or department edits.",
      "match_mode": "any",
      "fill_colors": ["FFFFFF00"],
      "comment_patterns": [],
      "comment_column": "comments",
      "nonblank_columns": [],
      "blank_columns": []
    }
  ]
}
```

### March 2026 Sample Validation

For the reviewed GL workbook:

- Sheet: `NEW MONTH`
- Header row: `5`
- Comment-based addback rows: `122`
- Comment-based addback total: `23,195.16`
- Red/pink rows: `123`
- Red/pink total: `23,248.67`
- Magenta unknown-charge rows: `1`
- Yellow account/department edit rows: `139`

The extractor emits a warning that the red/pink row total differs from the comment-based addback total by `53.51`.

### Output Tables

The extractor exposes:

- `ledger`: normalized reviewed GL rows
- `groups`: one row per matched configured row group
- `summaries`: row count and amount total by group

Important fields include:

- `source_file`
- `sheet`
- `row`
- `date`
- `transaction_type`
- `num`
- `name`
- `memo_description`
- `account_section`
- `account`
- `amount`
- `department`
- `expected_account`
- `expected_department`
- `comments`
- `dominant_fill_color`
- `row_fill_colors`
- `colored_cells`
- `group_name`
- `amount_cell`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

## DTC & WS Monthly Revenue

Module:
- `hgf_pnl.extractors.monthly_revenue`

CLI:

```bash
.venv/bin/python scripts/extract_monthly_revenue.py \
  "sample_files/Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx"
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_monthly_revenue.py \
  "sample_files/Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx" \
  --output-dir tmp/monthly_revenue \
  --format csv
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_monthly_revenue.py --init-config configs/monthly_revenue.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_monthly_revenue.py \
  "sample_files/Workpapers MARCH/DATA/DTC & WS Monthly Revenue - report (03.01-03.31) (1).xlsx" \
  --config configs/monthly_revenue.json
```

### Heuristic Behavior

The extractor detects sheet roles by sheet-name keywords:
- `summary`
- `shopify`
- `refunds`
- `coupons`

It detects headers on detail sheets by scanning the first rows for known column aliases plus fuzzy matching. The summary sheet is parsed as display rows grouped under section anchors such as `REVENUE` and `REFUNDS`.

Formula evaluation is enabled by default for consistency with other Excel extractors, although the March 2026 sample contains no formulas.

### Output Tables

The extractor exposes:
- `summary`: normalized rows from the workbook summary/pivot display.
- `sales`: Shopify order-level revenue rows.
- `refunds`: refund rows, including detail rows without an amount by default.
- `coupons`: coupon/order rows.

### March 2026 Sample Validation

For the sample workbook:
- Shopify sales rows: `159`
- Shopify net sales: `163,408.05`
- DTC net sales: `146,599.00`
- WS net sales: `16,809.05`
- Refund rows: `19`
- Refund rows with amount: `16`
- Refund amount total: `10,114.56`
- OG-DTC refunds: `9,940.81`
- OG-WS refunds: `173.75`
- Coupon rows: `19`
- Coupon total: `17,212.05`

### Output Shape

Summary rows:
- `source_file`
- `sheet`
- `section`
- `metric`
- `row`
- `label`
- `amount`
- `is_total_row`
- `source_cell`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

Sales rows:
- `source_file`
- `sheet`
- `row`
- `day`
- `order_name`
- `customer_name`
- `customer_email`
- `gross_sales`
- `orders`
- `quantity_ordered_per_order`
- `average_order_value`
- `quantity_returned`
- `net_sales`
- `channel`
- `method`
- `source_headers`
- `formula_status`
- `formula_detail`

Refund rows:
- `source_file`
- `sheet`
- `row`
- `date`
- `year`
- `month`
- `requested_by`
- `order_number`
- `division`
- `amount`
- `has_amount`
- `refund_category`
- `return_reason`
- `return_sku`
- `model_number`
- `image`
- `size`
- `acrylic`
- `embellishment`
- `notes`
- `saved`
- `payment_method`
- `pp_customer_email`
- `jason_approval`
- `jason_comments`
- `notified_date`
- `source_headers`
- `formula_status`
- `formula_detail`

Coupon rows:
- `source_file`
- `sheet`
- `row`
- `order`
- `date`
- `customer`
- `payment_status`
- `fulfillment_status`
- `items`
- `total`
- `channel`
- `delivery_status`
- `delivery_method`
- `source_headers`
- `formula_status`
- `formula_detail`

## Division COGS

Module:
- `hgf_pnl.extractors.division_cogs`

CLI:

```bash
.venv/bin/python scripts/extract_division_cogs.py \
  "sample_files/Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx"
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_division_cogs.py \
  "sample_files/Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx" \
  --output-dir tmp/division_cogs \
  --format csv
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_division_cogs.py --init-config configs/division_cogs.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_division_cogs.py \
  "sample_files/Workpapers MARCH/DATA/INTERNAL - Division COGS 2019 - Current (26).xlsx" \
  --config configs/division_cogs.json
```

### Heuristic Behavior

The extractor detects:
- year matrix tabs whose sheet names look like `2018`, `2019`, ..., `2026`.
- partner detail tabs whose names look like `YYYY Partner Details`.
- older 2018/2019 COGS matrices where the header row starts with `COGS`.
- 2020+ matrices where the header row starts with `Month` and `Type`.
- month groups in partner detail tabs, whether the month headers are on row 1 or row 2.

Year matrix rows are forward-filled by month, because many workbook rows use merged month cells. Partner detail tabs are unpivoted from repeated month groups of `COGS`, `Material Cost`, and `Labor Cost`.

Formula evaluation is enabled by default. Unsupported formulas retain cached values when present and preserve formula status. This workbook includes `VLOOKUP` formulas with workbook-index and whole-column references; those are currently reported as unsupported rather than evaluated.

### Output Tables

The extractor exposes:
- `matrix`: year-tab rows normalized to month/type/channel/amount.
- `partner_details`: partner-detail rows normalized to month/partner/measure/amount.

### March 2026 Sample Validation

For the sample workbook:
- Year matrix rows: `3,276`
- Partner detail rows: `6,020`
- 2026 March COGS total column: `213,273.23`
- 2026 March Online - USA COGS: `181,629.67`
- 2026 March D2C partner detail COGS: `19,658.51`
- 2026 March D2C partner detail material cost: `15,596.90`
- 2026 March D2C partner detail labor cost: `4,061.61`

### Output Shape

Matrix rows:
- `source_file`
- `sheet`
- `year`
- `month_num`
- `month_name`
- `month`
- `row`
- `type`
- `channel`
- `amount`
- `raw_value`
- `is_total_column`
- `source_cell`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

Partner detail rows:
- `source_file`
- `sheet`
- `year`
- `month_num`
- `month_name`
- `month`
- `row`
- `partner`
- `measure`
- `amount`
- `raw_value`
- `source_cell`
- `month_header_row`
- `measure_header_row`
- `formula`
- `cached_amount`
- `calculated_amount`
- `calculation_status`
- `calculation_detail`

## Trend House Revenue Report

Module:
- `hgf_pnl.extractors.th_revenue`

CLI:

```bash
.venv/bin/python scripts/extract_th_revenue.py \
  "sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx"
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_th_revenue.py \
  "sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx" \
  --output-dir tmp/th_revenue \
  --format csv
```

Disable formula evaluation:

```bash
.venv/bin/python scripts/extract_th_revenue.py \
  "sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx" \
  --no-calculate-formulas
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_th_revenue.py --init-config configs/th_revenue.json
```

Then run with:

```bash
.venv/bin/python scripts/extract_th_revenue.py \
  "sample_files/Workpapers MARCH/DATA/TH March 2026 Revenue Report.xlsx" \
  --config configs/th_revenue.json
```

### Heuristic Behavior

The extractor detects sheet roles by configurable sheet-name keywords:
- `summary`
- `details`
- `usa_stock`

It detects header rows by scanning the first rows of each sheet for expected revenue/cost/account headers. Header names are mapped to canonical names by aliases plus fuzzy matching, so small variations like `Shipping Cost` vs `Shipping cost` should still parse.

### Output Tables

The extractor exposes:
- `account_summary`: account-level rows from the `Summary` sheet.
- `po_details`: complete PO-level rows from the `Details` sheet.
- `usa_stock`: supporting USA Stock rows from the `USA Stock` sheet.
- `all_rows`: all extracted rows, including supporting tabs.

Important: in the March 2026 sample, the `Details` sheet already includes the USA Stock rows. Treat `usa_stock` as a supporting/subset table, not something to add to `po_details` revenue.

### Output Shape

Rows are normalized to:

- `source_file`
- `sheet`
- `role`
- `row`
- `is_total_row`
- `internal_po`
- `account`
- `revenue`
- `production_cost`
- `shipping_cost`
- `tariff`
- `total_cost`
- `gross_margin_pct`
- `gross_margin_amount`
- `computed_gross_margin_pct`
- `computed_gross_margin_amount`
- `validation_warnings`
- `formula_status`
- `formula_detail`
- `source_headers`

### March 2026 Sample Validation

For the sample workbook:
- Summary non-total revenue: `1,242,393.40`
- Details non-total revenue: `1,242,393.40`
- USA Stock non-total revenue: `31,473.30`
- Summary non-total total cost: `750,574.50`
- Details non-total total cost: `750,574.50`

## Chargeback Report PDF

Module:
- `hgf_pnl.extractors.chargeback_pdf`

Agentic profile step:

```bash
.venv/bin/python scripts/profile_chargeback_pdf.py \
  "sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf" \
  --output-dir tmp/chargeback_pdf_profile
```

This writes:
- `chargeback_pdf_profile.md`: compact review packet with line candidates, anchor candidates, and table previews.
- `chargeback_pdf_profile.json`: structured profile data.
- `chargeback_pdf_raw_text.txt`: all extracted text lines with page/line references.
- `chargeback_pdf_suggested_config.json`: starter config inferred from the PDF.

CLI:

```bash
.venv/bin/python scripts/extract_chargeback_pdf.py \
  "sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf"
```

Write CSV outputs:

```bash
.venv/bin/python scripts/extract_chargeback_pdf.py \
  "sample_files/Workpapers MARCH/DATA/- OG _ Chargeback Report - 03. March 2026.pdf" \
  --output-dir tmp/chargeback_pdf \
  --format csv
```

Create an editable config:

```bash
.venv/bin/python scripts/extract_chargeback_pdf.py --init-config configs/chargeback_pdf.json
```

Recommended agent workflow:
- Run `profile_chargeback_pdf.py`.
- Inspect `chargeback_pdf_profile.md` and `chargeback_pdf_raw_text.txt`.
- Edit `chargeback_pdf_suggested_config.json` if anchors/category order/month/year need adjustment.
- Run `extract_chargeback_pdf.py --config <edited-config>`.

### Parser Strategy

The PDF is an email export. `pdfplumber` table extraction loses some values in the monthly summary table, so the extractor primarily parses line text and preserves source page/line references. The profiler still captures table previews because they are useful context for an agent deciding how to configure anchors.

It extracts:
- `monthly_summary`: month/category trend rows from the large chargeback summary.
- `customer_detail`: March customer/reseller deduction rows by department.
- `reconciliation`: CB report vs QB differences and notes.
- `notes`: unstructured explanatory lines in the reconciliation block.

### Output Highlights

For the March 2026 sample:
- Monthly chargeback grand total: `-104,205.00`
- Allowance: `-65,092.00`
- Penalty: `-7,821.00`
- Amazon holdback provision: `-15,661.00`
- Return: `-15,596.00`
- Software fees: `-34.00`
- Customer-detail grand total: `-88,543.00`
- Reconciliation grand-total difference: `-4,814.72`

### Important Caveat

The customer-detail lines are rounded to whole dollars in the PDF text. The sum of individual non-total customer rows is `-88,541`, while the provided `March Total` and `Grand Total` rows are `-88,543`. Treat provided total rows as authoritative for the PDF block and retain source lines for audit.

### Output Shape

Monthly summary rows:
- `year`
- `month_num`
- `month_name`
- `category`
- `category_position`
- `amount`
- `percent_of_total`
- `grand_total`
- `source_page`
- `source_line`
- `source_text`
- `confidence`

Customer detail rows:
- `month_name`
- `department`
- `customer`
- `amount`
- `is_total_row`
- `source_page`
- `source_line`
- `source_text`

Reconciliation rows:
- `customer`
- `cb_report_amount`
- `qb_amount`
- `difference`
- `note`
- `is_total_row`
- `source_page`
- `source_line`
- `source_text`
