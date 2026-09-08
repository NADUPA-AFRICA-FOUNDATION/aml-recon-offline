#!/usr/bin/env python3
"""Launch AML Training Reconciliation Studio in the default browser."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"

cmd = [
    sys.executable,
    "-m",
    "streamlit",
    "run",
    str(APP),
    "--server.address=127.0.0.1",
    "--server.port=8501",
    "--server.headless=false",
    "--browser.gatherUsageStats=false",
]
raise SystemExit(subprocess.call(cmd, cwd=str(ROOT)))
