---
description: Bootstrap the HGF P&L plugin — create the Python venv, install hgf_pnl, and verify the toolchain.
---

# hgf-pnl-bootstrap

The user wants to prepare this checkout to run the `hgf-monthly-close` skill. Walk through the steps below in order. Adapt freely if anything fails — surface the failure to the user, suggest a fix, and continue once they confirm.

The plugin root is the directory that contains `pyproject.toml`, `requirements.txt`, the `.claude-plugin/` folder, and `skills/hgf-monthly-close/`. Run all commands from there.

## 1. Verify Python

```bash
python3 --version
```

The project requires Python `>=3.12` (see `pyproject.toml`). If the system `python3` is older, look for `python3.12`, `python3.13`, or a `uv` install. Tell the user what you found and ask before downloading or installing a new interpreter.

## 2. Create the venv (if missing)

Check for `.venv/bin/python` at the plugin root. If it does not exist:

- Prefer `uv venv` if `uv` is on PATH, since it is faster and matches how the project was originally bootstrapped.
- Otherwise fall back to `python3 -m venv .venv`.

Tell the user which path you took.

## 3. Install dependencies and the editable package

`hgf_pnl` lives under `skills/hgf-monthly-close/hgf_pnl` and is wired into `pyproject.toml` via `[tool.setuptools.packages.find]`. An editable install puts it on the venv's import path so the scripts in `skills/hgf-monthly-close/scripts/` work without `PYTHONPATH` games.

If `uv` is available:

```bash
uv pip install -e .
```

Otherwise:

```bash
.venv/bin/python -m pip install -e .
```

If `pip` is missing from a `uv`-built venv, use `uv pip` rather than bootstrapping pip into it.

## 4. Verify the install

```bash
.venv/bin/python -c "import hgf_pnl; print(hgf_pnl.__file__)"
```

The path should end with `skills/hgf-monthly-close/hgf_pnl/__init__.py`. If it points anywhere else, the editable install is stale — reinstall.

Then smoke-test one script:

```bash
.venv/bin/python skills/hgf-monthly-close/scripts/discover_package.py --help
```

## 5. Check the recalculation toolchain

The consolidated writer sets workbook recalc flags but cannot recalculate cached formula values itself. LibreOffice headless is the supported recalculation path.

```bash
which libreoffice || which soffice
```

If neither is installed, tell the user that generated workbooks will need to be opened in Excel or LibreOffice for formulas to refresh, and ask whether they want help installing LibreOffice.

## 6. Report status

Summarize for the user, in this shape:

- Python interpreter used and version.
- Whether the venv was reused or freshly created.
- Whether `hgf_pnl` resolves to the in-skill path.
- Whether LibreOffice is available, and what that means for workbook recalculation.
- Any warnings or surprises encountered during the steps above.

Do not proceed to running close-package work in the same response unless the user explicitly asks. Bootstrap is the only goal of this command.
