# AML Training Reconciliation — Offline

This project runs locally and does not use an API, cloud service, or AI model.

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
