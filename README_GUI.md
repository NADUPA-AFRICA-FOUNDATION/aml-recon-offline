# AML Training Reconciliation Studio — Offline GUI

A polished local GUI for the AML/CFT training reconciliation engine.

## What changed

You no longer need to interact with the command line for normal reconciliation work.

The GUI provides:

- Drag-and-drop Excel/CSV uploads
- One-click reconciliation
- Offline/local processing
- Internal/external classification controls
- Live management KPIs
- Exceptions preview
- Validation status
- One-click Excel export
- The same deterministic reconciliation logic behind the interface

## First-time installation

You need Python 3.10+.

### macOS / Linux

```bash
cd aml_recon_offline
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows PowerShell

```powershell
cd aml_recon_offline
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Once dependencies are installed, the app can run without internet access.

## Launch the GUI

### macOS / Linux

Double-click/run `run_gui_mac_linux.sh`, or:

```bash
python3 start_gui.py
```

### Windows

Double-click:

```text
run_gui_windows.bat
```

The app opens locally in your browser at:

```text
http://127.0.0.1:8501
```

This is a local address. Your uploaded files are processed on the computer running the application.

## Normal workflow

1. Launch the app.
2. Drag all AML training Excel/CSV exports into the upload area.
3. Review the internal email-domain and role-marker settings.
4. Click **Run full reconciliation**.
5. Review KPIs, population split, exceptions, and validation.
6. Click **Download reconciled workbook**.

## Privacy

The reconciliation engine itself does not use ChatGPT, an API, cloud storage, or an external web service. Uploaded files are written only to a temporary local processing folder during a run and the temporary workspace is deleted afterwards.

## Workbook output

The GUI generates:

`AML_Training_Full_Reconciliation.xlsx`

with:

1. Executive Summary
2. Course Reconciliation
3. Year Movement
4. Participant History
5. All Records
6. Exceptions
7. Methodology

## Important matching controls

- User ID is the primary learner key.
- Username/email is the fallback identifier.
- Names are not fuzzy-matched.
- Multiple courses are legitimate participation, not duplicates.
- Multi-year participation is legitimate, not duplication.
- Suspicious exact assignment repeats are flagged for analyst review instead of deleted.

## Interface improvements

- A focused workspace with a three-step workflow and grouped workbook contents.
- Consistent teal accents, readable metric cards, keyboard focus indicators and narrow-screen styles.
- Searchable exceptions and clear empty, success and error states.
- Results are cleared when exports or classification rules change, preventing stale downloads.
- Separate uploads with identical filenames are preserved with numbered source prefixes.
- The GUI and command line share one reconciliation engine.

## Verification

```bash
python -m unittest discover -s tests -v
```

The tests cover startup, reconciliation, seven-sheet export, duplicate filenames,
changed rules, exception search and processing errors using synthetic records.
