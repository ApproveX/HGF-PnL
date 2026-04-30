---
name: hgf-monthly-close
description: Use for HGF monthly close packages: discover files, configure and run extractors, review audit warnings and overrides, generate consolidated P&L workbooks, and validate output artifacts.
---

# HGF Monthly Close

This skill operates the HGF P&L close tooling in this repository. Use it when the user asks to process, inspect, validate, or generate an HGF monthly P&L package.

## Core Rule

Treat source workpapers as immutable. Do not edit files under the client-provided package unless the user explicitly asks. Generated configs, manifests, extracted tables, values JSON, and output workbooks should go into a run directory such as:

```text
tmp/runs/<period-slug>/
```

For client-facing generated workbooks, use clear names such as:

```text
tmp/runs/march-2026/HGF_CONSOLIDATED_MARCH_2026_GENERATED.xlsx
```

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

## Discovery And Manifest

Run discovery first:

```bash
.venv/bin/python scripts/discover_package.py \
  "sample_files/Workpapers MARCH" \
  --manifest-output tmp/runs/march-2026/run_manifest.json
```

Use `--inspect-workbooks` when classification is ambiguous or when sheet names are useful:

```bash
.venv/bin/python scripts/discover_package.py \
  "sample_files/Workpapers MARCH" \
  --inspect-workbooks \
  --discovery-output tmp/runs/march-2026/discovery.json \
  --manifest-output tmp/runs/march-2026/run_manifest.json
```

Review:
- `source_input` files with extractor matches.
- `instruction` files, especially addback PDFs and email-thread PDFs.
- `supporting_input` files that may require manual interpretation.
- `deliverable_or_prior_output` files.
- missing expected inputs, such as BR Info or Payroll when they live outside the selected folder.

## Extractors

Use the configured extractors rather than writing one-off parsing code.

Available scripts:

```text
scripts/extract_pl_by_dept.py
scripts/extract_th_revenue.py
scripts/extract_chargeback_pdf.py
scripts/extract_payroll_journal.py
scripts/extract_br_info.py
scripts/extract_monthly_revenue.py
scripts/extract_division_cogs.py
scripts/extract_addbacks_gl.py
```

Most Excel extractors support:
- `--config <path>`
- `--init-config <path>`
- `--no-calculate-formulas`

Formula evaluation is on by default for Excel extractors. Unsupported formulas should be preserved as warnings/status values, not silently coerced.

## Agentic Configuration

When a file is natural-language-heavy or accountant-reviewed, inspect it before running strict extraction.

Examples:
- For chargeback PDFs, inspect text/table structure before choosing PDF table extraction settings.
- For addbacks, read the PDF/email instructions first, then configure the reviewed GL workbook extractor.
- For yellow/red/magenta reviewed GL rows, preserve both source color and semantic comments.

When changing extractor config, write the config to the run directory and record the config path in the manifest.

## Review Items

Always summarize these before writing the final workbook:

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

For approved overrides, record:
- target key or cell
- original value
- override value
- reason
- source
- approver, if known

## Consolidated Writer

The writer should not write directly to the visible `MARCH 2026 FULL ` tab except for an explicitly approved template change. It writes to hidden raw-data tabs:

```text
RAW DATA_Master File
RAW DATA_COGS & Freight
RAW DATA_Payroll
```

Known exception: `MARCH 2026 FULL !EB48` is the hidden source-total cell for Employee Benefits. The March template does not stage this BR Info override in a raw-data tab, so the writer config may write that hidden source cell directly.

Generate an editable writer config:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --init-config tmp/runs/march-2026/consolidated_writer_config.json
```

Write the workbook from approved values:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/runs/march-2026/HGF_CONSOLIDATED_MARCH_2026_GENERATED.xlsx \
  --values tmp/runs/march-2026/consolidated_values.json \
  --config tmp/runs/march-2026/consolidated_writer_config.json
```

For sample round-trip testing only, values can be extracted from a completed workbook:

```bash
.venv/bin/python scripts/write_consolidated_pnl.py \
  "sample_files/HGF CONSOLIDATED_MARCH 2026 Template.xlsx" \
  tmp/unused.xlsx \
  --values tmp/runs/march-2026/consolidated_values_from_final.json \
  --example-values-from "sample_files/P&L_S/FULL COMPANY P&L_s/HGF CONSOLIDATED_MARCH 2026_FINAL.xlsx"
```

Do not use a completed workbook as source values for production unless the user explicitly wants a template/writer validation exercise.

## Recalculation

The writer sets workbook recalculation flags. If LibreOffice is available, use it later to recalculate cached formula values and re-save the workbook. If LibreOffice is not installed, clearly tell the user that Excel/LibreOffice should recalculate on open.

Check availability:

```bash
which libreoffice
which soffice
```

## Output Validation

Before saying an output workbook is ready:

1. Confirm the writer completed with no validation failures.
2. Compare mapped raw-tab cells to approved values.
3. Check skipped cells and explain why they were skipped.
4. Check workbook recalculation flags.
5. If recalculation is available, inspect visible totals and formula errors after recalculation.
6. Record validation results in the manifest.

Known workbook-specific validation concerns:
- `MARCH 2026 FULL ` is formula-heavy and relies on raw-data tabs.
- The `Payroll` tab is the payroll source of truth. The `Payroll Distribution` tab is an intermediary copy format and may contain stale formulas, especially in the March 2026 `Lital Allocation in G&A Exp` block.
- Some hidden legacy columns reference `RAW DATA_Master File!B134:B151`, while the template raw tab currently ends at row 133. Flag these as stale hidden references unless the client says otherwise.

## Documentation

Use these docs for implementation details:

```text
docs/pipeline.md
docs/extractors.md
docs/consolidated_writer.md
docs/priority_file_exploration.md
```

## Communication

When reporting status, separate:
- extraction/configuration facts
- validation warnings
- assumptions
- files produced
- actions still requiring human review

Never present a generated workbook as client-ready if formula recalculation or review warnings remain unresolved.
