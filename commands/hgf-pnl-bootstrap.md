---
description: Bootstrap the HGF P&L plugin — create an external Python venv, install hgf_pnl, and verify the toolchain.
---

# hgf-pnl-bootstrap

The user wants to prepare their machine to run the `hgf-monthly-close` skill. The plugin install directory is read-only on most setups, so the venv lives outside it. Walk through the steps below in order. Adapt freely if anything fails — surface the failure to the user, suggest a fix, and continue once they confirm.

## Locations

| What | Path |
|---|---|
| Plugin source (read-only) | `$CLAUDE_PLUGIN_ROOT` — the directory containing `pyproject.toml` and `.claude-plugin/`. If the env var is unset, find this directory by locating the `.claude-plugin/plugin.json` whose `name` is `hgf-pnl`. |
| Plugin home (writable) | `${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl` |
| Plugin venv | `<plugin home>/venv` |
| Plugin-root pointer | `<plugin home>/plugin_root` (text file containing the resolved plugin source path) |

Resolve `$CLAUDE_PLUGIN_ROOT` once at the start. Use the resolved value (an absolute path) in every command below; do not rely on the variable persisting across shell calls.

## 1. Verify Python

```bash
python3 --version
```

The project requires Python `>=3.10` (see `pyproject.toml`). If the system `python3` is older, look for `python3.10`, `python3.11`, `python3.12`, `python3.13`, or a `uv` install. Tell the user what you found and ask before installing a new interpreter.

## 2. Create the plugin home and venv

Make the plugin home directory if missing:

```bash
mkdir -p "${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl"
```

Create the venv inside it (skip if `venv/bin/python` already exists). Prefer `uv` since it is faster and matches how the project was originally built:

```bash
uv venv "${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv"
```

If `uv` is not on PATH, fall back to:

```bash
python3 -m venv "${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv"
```

Tell the user which path you took.

## 3. Install hgf_pnl as a wheel (NOT editable)

The plugin source is read-only. An editable install (`pip install -e`) writes `*.egg-info/` back into the source tree and will fail. Install as a regular wheel instead, which copies the package code into the venv's `site-packages/`:

If `uv` is available:

```bash
uv pip install --python "${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv/bin/python" "<resolved CLAUDE_PLUGIN_ROOT>"
```

Otherwise:

```bash
"${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv/bin/python" -m pip install "<resolved CLAUDE_PLUGIN_ROOT>"
```

If the venv was built by `uv` and lacks `pip`, use `uv pip` rather than bootstrapping pip into it.

Note that this means: when the plugin updates, the user must re-run `/hgf-pnl-bootstrap` to pick up the new code. Tell the user this so they know.

## 4. Record the plugin root

Write the resolved plugin root path to the pointer file so the skill can find the scripts later:

```bash
printf '%s\n' "<resolved CLAUDE_PLUGIN_ROOT>" \
  > "${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/plugin_root"
```

## 5. Verify the install

```bash
"${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv/bin/python" \
  -c "import hgf_pnl; print(hgf_pnl.__file__)"
```

The path should land inside the venv's `site-packages/hgf_pnl/`. If it does not, the install is wrong — investigate before continuing.

Then smoke-test one script from the plugin source:

```bash
"${XDG_DATA_HOME:-$HOME/.local/share}/hgf-pnl/venv/bin/python" \
  "<resolved CLAUDE_PLUGIN_ROOT>/skills/hgf-monthly-close/scripts/discover_package.py" --help
```

## 6. Check the recalculation toolchain

The consolidated writer sets workbook recalculation flags but cannot recalculate cached formula values itself. LibreOffice headless is the supported recalculation path.

```bash
which libreoffice || which soffice
```

If neither is installed, tell the user that generated workbooks will need to be opened in Excel or LibreOffice for formulas to refresh, and ask whether they want help installing LibreOffice.

## 7. Report status

Summarize for the user, in this shape:

- Python interpreter used and version.
- Whether the venv was reused or freshly created, and where it lives.
- Whether `hgf_pnl` resolves to the venv site-packages.
- Whether LibreOffice is available, and what that means for workbook recalculation.
- The plugin root path that was recorded, so the user can confirm.
- Any warnings or surprises encountered along the way.

Do not proceed to running close-package work in the same response unless the user explicitly asks. Bootstrap is the only goal of this command.
