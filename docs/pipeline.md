# Package Discovery And Run Manifest

Modules:
- `hgf_pnl.pipeline.discovery`
- `hgf_pnl.pipeline.manifest`

CLI:

```bash
.venv/bin/python scripts/discover_package.py \
  "sample_files/Workpapers MARCH" \
  --discovery-output tmp/workpapers_march.discovery.json \
  --manifest-output tmp/workpapers_march.manifest.json
```

Optionally inspect workbook sheet names:

```bash
.venv/bin/python scripts/discover_package.py \
  "sample_files/Workpapers MARCH" \
  --inspect-workbooks \
  --manifest-output tmp/workpapers_march.manifest.json
```

## Discovery

Discovery recursively scans a close package folder and emits one record per file.

It skips by default:
- `:Zone.Identifier` sidecar files
- temporary Office lock files beginning with `~$`
- hidden dot folders

Each discovered file records:
- absolute path
- relative path
- file name and suffix
- size and modified timestamp
- role
- document type
- matched extractor or writer
- confidence
- classification reasons
- optional metadata, such as workbook sheet names

Known extractor/writer matches currently include:
- `pl_by_dept`
- `th_revenue`
- `chargeback_pdf`
- `payroll_journal`
- `br_info`
- `monthly_revenue`
- `division_cogs`
- `addbacks_gl`
- `writer:consolidated_pnl`

## Manifest

The manifest is the durable run record. The initial manifest is generated from discovery before extractors run.

It includes:
- run id
- package root
- detected period
- discovered inputs
- selected/unselected state
- config paths
- output paths
- overrides
- pipeline events
- warnings
- final workbook/report paths

Initial status is `discovered`. Later orchestration can advance it through:
- `configured`
- `extracted`
- `reviewed`
- `written`
- `validated`
- `needs_review`
- `failed`

## Agent Use

The agent should treat discovery classifications as suggestions, not facts. For low-confidence or ambiguous files, it should inspect the file and either update the manifest classification or mark the input unselected with a reason.

The manifest should be updated after each major step:
- extractor config chosen
- extractor run completed
- warning reviewed
- override approved
- writer run completed
- output validation completed

This gives the final close package a durable audit trail of what was used, what was ignored, and why.
