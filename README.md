# AML Training Reconciliation — Offline

This project runs locally and does not use an API, cloud service, or AI model.

On managed work devices, local processing does not override browser/DLP policy.
See [Work-device upload policies](README_GUI.md#work-device-upload-policies) for
the browser-free workflow, loopback-only GUI controls, and the security review
checklist.

## Output

It generates `AML_Training_Full_Reconciliation.xlsx` with:

1. Executive Summary
2. Course Reconciliation
3. Year Movement
4. Participant History
5. All Records
6. Exceptions
7. Methodology

## Setup in VS Code

Use Python 3.10+.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

Put all raw `.xlsx`, `.xlsm`, `.xls`, or `.csv` LMS exports inside `input/`.

Then run:

```bash
python aml_reconcile.py --input ./input --output ./output/AML_Training_Full_Reconciliation.xlsx --config ./config.json
```

On macOS, if `python` is unavailable:

```bash
python3 aml_reconcile.py --input ./input --output ./output/AML_Training_Full_Reconciliation.xlsx --config ./config.json
```

## Core controls

- User ID is the primary participant key.
- Username/email is used only when User ID is unavailable.
- No fuzzy-name merging is used.
- Multiple courses for the same person are retained.
- Multi-year participation is retained.
- Suspected exact repeats are flagged, not deleted.
- Missing critical fields are listed in Exceptions.
- Late completions still count as Completed and are also flagged as Completed Late.
- Assignment Year prefers assignment-begin year, then assignment-end year, filename year, then completion year.

## Organisation-specific settings

Edit `config.json` to change the internal email domain, internal staff role markers, or critical fields.

The script already knows many common LMS heading aliases. If your source headings are different, add them to `column_aliases` in `aml_reconcile.py`.

## Validation

A good run ends with:

```text
Validation: PASS
```

If it says `REVIEW REQUIRED`, inspect the validation notes in the console and workbook.

## Independently verify a workbook

`recon_verifier.py` is a separate, configuration-driven audit tool. It reads the
raw exports again and checks row counts, mapped fields and dates, derived values,
rollup tables, and individual dashboard cells against a completed workbook. It
does not modify either the sources or the workbook.

```bash
python recon_verifier.py --config recon_config_aml_training.json --report recon_report.md
```

The file locations can also be supplied without editing a shared config:

```bash
python recon_verifier.py --config recon_config_aml_training.json \
  --raw-glob '/data/Training_Progress_Report_*.xlsx' \
  --recon-path /data/AML_Training_Reconciliation.xlsx
```

Use `--data-dir /data` to relocate paths that are relative in the config file.
Absolute paths remain absolute; `--raw-glob` and `--recon-path` always take
precedence.

Every configured check prints a `PASS` or `FAIL` line. The optional Markdown
report includes up to 20 mismatch samples per check. Exit status `0` means all
checks passed, `1` means at least one reconciliation check failed, and `2` means
the files or configuration could not be read.

All paths in the JSON file are resolved relative to the config file. Copy the
included config for another reconciliation and adjust:

- `raw` and `recon` for workbook locations and sheet layouts. Raw sheets can use
  a flat header or a two-row grouped header (`Group::Field`).
- `join` for a positional match or, preferably, stable raw/master business keys.
- `field_map` and `date_field_map` for direct row-level comparisons.
- `derived_checks` for `role_evidence`, `value_in_set`, `value_equals`,
  `year_of_date`, or `duration_bucket` rules.
- `summary_checks` for grouped or grand-total tables, and `cell_checks` for
  dashboard cells. Supported aggregations are `count`, `nunique`, `sum`, and
  `mean`.

For stacked tables, set `data_end_row` on a summary check so rows belonging to a
second table are not interpreted as part of the first one. Positional joins
require source order to be preserved; key joins should be used whenever a stable
business key exists.
