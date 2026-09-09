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

### Work-device upload policies

The browser interface looks like a website, but the supplied launcher binds it
only to `127.0.0.1`. That loopback address is the same computer—not an internet
site—and is not reachable from another device. Browser telemetry is disabled,
and Streamlit's CORS and XSRF protections are explicitly enabled. The page uses
plain HTTP on the local loopback connection; it should not be exposed on a LAN or
public interface.

An organisation's browser, DLP agent, or endpoint policy may still classify a
file selection on `http://127.0.0.1:8501` as an upload. Do not bypass that
control. Use one of these approved deployment patterns:

1. **Preferred when browser uploads are restricted:** use the command-line
   workflow. Put approved exports in `input/` and run `aml_reconcile.py`; for
   independent verification, place the files beside the verifier config or use
   `--data-dir`. No browser file picker is involved.
2. **Local GUI after security approval:** ask IT/security to assess and, if
   appropriate, allow the exact loopback origin `http://127.0.0.1:8501`. Do not
   request a broad web-domain exception.
3. **Organisation-hosted deployment:** treat it as a normal internal application:
   place it behind the organisation's SSO/reverse proxy and TLS, restrict network
   access, define retention and logging, patch dependencies, and complete the
   organisation's privacy/security review. The bundled launcher is not an
   authenticated multi-user server and must not be used for this pattern.

Before approval, provide the security team with:

- the source repository and locked/approved dependency versions;
- the data classification of the LMS exports and generated workbook;
- confirmation that processing is local and temporary files are deleted after
  each run;
- the loopback-only launch settings in `start_gui.py`;
- the required input/output folders and their operating-system permissions;
- evidence from endpoint monitoring that the process makes no outbound
  connections during a representative run.

Approval ultimately depends on your organisation's policy and security team;
local processing alone does not automatically make the tool compliant.

### Browser-free commands

```bash
python aml_reconcile.py --input ./input \
  --output ./output/AML_Training_Full_Reconciliation.xlsx \
  --config ./config.json

python recon_verifier.py --config recon_config_aml_training.json \
  --data-dir /approved/local/folder --report ./output/recon_report.md
```

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

## When an export is not recognised

The importer detects known column headings within the first 50 non-empty rows,
so report titles and blank rows above the table are supported. CSV imports accept
comma, semicolon, tab and pipe separators, including Excel `sep=` declarations.
If no training records are recognised, the error lists the detected headings.
Check these against the expected names (such as User ID, Email, Full Name,
Assignment Title and Completion Status). Custom headings can be mapped through
`column_aliases` in a CLI configuration file; for the GUI, extend
`DEFAULT_CONFIG["column_aliases"]` in `aml_reconcile.py`.
