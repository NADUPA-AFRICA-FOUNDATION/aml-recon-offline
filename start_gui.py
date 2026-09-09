#!/usr/bin/env python3
"""Launch AML Training Reconciliation Studio in the default browser."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"


def build_command() -> list[str]:
    """Return the security-conscious, loopback-only Streamlit command."""
    return [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(APP),
        "--server.address=127.0.0.1",
        "--server.port=8501",
        "--server.headless=false",
        "--server.enableXsrfProtection=true",
        "--server.enableCORS=true",
        "--server.fileWatcherType=none",
        "--server.runOnSave=false",
        "--browser.gatherUsageStats=false",
    ]


def main() -> int:
    return subprocess.call(build_command(), cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
