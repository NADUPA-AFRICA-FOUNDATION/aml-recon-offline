#!/usr/bin/env python3
"""Offline AML/CFT training reconciliation.

Usage:
  python aml_reconcile.py --input ./input \
      --output ./output/AML_Training_Full_Reconciliation.xlsx \
      --config ./config.json

Reads .xlsx/.xlsm/.xls/.csv files recursively. No network/API/AI required.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

DEFAULT_CONFIG: Dict[str, Any] = {
    "internal_email_domains": ["safaricom.co.ke"],
    "internal_role_markers": ["safaricom", "employee", "staff", "internal"],
    "critical_fields": [
        "User ID", "Username", "Full Name", "Assignment ID",
        "Assignment Title", "Completion Status"
    ],
    "column_aliases": {
        "User ID": ["user id", "userid", "user_id", "employee id", "learner id", "person id"],
        "Full Name": ["full name", "fullname", "name", "learner name", "participant name", "employee name"],
        "Username": ["username", "email", "email address", "user email", "login", "login id"],
        "Roles": ["roles", "role", "user roles", "job role", "learner role"],
        "Staff Category": ["staff category", "staff type", "employee category", "person type"],
        "Department": ["department", "dept", "business unit", "function"],
        "Region": ["region", "area", "zone"],
        "Territory": ["territory", "territory name"],
        "Shop/Team": ["shop/team", "shop", "team", "branch", "location"],
        "Learning Item ID": ["learning item id", "learning id", "item id", "course id"],
        "Learning Title": ["learning title", "learning item", "course title", "course name", "item title"],
        "Training Category": ["training category", "category", "learning category"],
        "Assignment Type": ["assignment type", "type"],
        "Assignment ID": ["assignment id", "assignment_id", "training assignment id", "enrollment id", "enrolment id"],
        "Assignment Title": ["assignment title", "training title", "assigned learning", "assignment name"],
        "Assignment Status": ["assignment status", "status", "enrollment status", "enrolment status"],
        "Assignment Begin": ["assignment begin", "assignment start", "start date", "assigned date", "begin date"],
        "Assignment End": ["assignment end", "assignment due", "end date", "due date", "deadline"],
        "Completion Status": ["completion status", "completion_status", "learning status", "training status"],
        "Result Status": ["result status", "result", "pass status", "outcome"],
        "Training Hours": ["training hours", "hours", "duration hours", "learning hours"],
        "Completion Date": ["completion date", "completed date", "date completed", "completion datetime"]
    },
    "status_rules": {
        "completed": [r"^\s*completed\b", r"\bcomplete[d]?\b"],
        "late": [r"out\s*of\s*deadline", r"outof\s*deadline", r"\blate\b", r"after\s*deadline"],
        "started": [r"started.*not.*completed", r"in\s*progress", r"\bstarted\b"],
        "not_started": [r"not\s*started", r"not\s*commenced"],
        "passed": [r"^\s*passed\s*$", r"\bpass(ed)?\b"]
    }
}

OUTPUT_COLUMNS = [
    "Source File", "Participant Key", "User ID", "Full Name", "Username",
    "Roles", "Staff Category", "Department", "Region", "Territory", "Shop/Team",
    "Learning Item ID", "Learning Title", "Training Category", "Assignment Type",
    "Assignment ID", "Assignment Title", "Assignment Status", "Assignment Begin",
    "Assignment End", "Completion Status", "Result Status", "Training Hours",
    "Population", "Assignment Year", "Completed Flag", "Passed Flag",
    "Started Not Completed Flag", "Not Started Flag", "Completed Late Flag",
    "Missing Critical Fields", "Exact Repeat Review"
]


def clean(v: Any) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""
    return re.sub(r"\s+", " ", s)


def header_key(v: Any) -> str:
    s = clean(v).lower().replace("_", " ").replace("-", " ")
    s = re.sub(r"[^a-z0-9/ ]+", "", s)
    return re.sub(r"\s+", " ", s).strip()


def deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            deep_merge(base[k], v)
        else:
            base[k] = v
    return base


def load_config(path: Optional[Path]) -> Dict[str, Any]:
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if path:
        with path.open("r", encoding="utf-8") as f:
            deep_merge(cfg, json.load(f))
    return cfg


def source_files(folder: Path) -> List[Path]:
    allowed = {".xlsx", ".xlsm", ".xls", ".csv"}
    return sorted(
        p for p in folder.rglob("*")
        if p.is_file() and p.suffix.lower() in allowed and not p.name.startswith("~$")
        and p.name != "AML_Training_Full_Reconciliation.xlsx"
    )


def read_file(path: Path) -> List[Tuple[str, pd.DataFrame]]:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return [("CSV", pd.read_csv(path, dtype=object, encoding=enc))]
            except UnicodeDecodeError:
                continue
        raise RuntimeError(f"Could not decode CSV: {path}")
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    return [(name, df) for name, df in sheets.items() if not df.empty]


def column_map(df: pd.DataFrame, aliases: Dict[str, Sequence[str]]) -> Dict[str, str]:
    available = {header_key(c): str(c) for c in df.columns}
    out = {}
    for target, vals in aliases.items():
        for alias in [target, *vals]:
            if header_key(alias) in available:
                out[target] = available[header_key(alias)]
                break
    return out


def matches(value: Any, patterns: Sequence[str]) -> bool:
    s = clean(value)
    return bool(s) and any(re.search(p, s, re.I) for p in patterns)


def make_participant_key(uid: Any, username: Any, full_name: Any, unique_row: int) -> str:
    if clean(uid):
        return f"ID:{clean(uid)}"
    if clean(username):
        return f"USER:{clean(username).lower()}"
    # Deliberately never merge on name alone.
    safe = re.sub(r"[^a-z0-9]+", "_", clean(full_name).lower()).strip("_") or "unknown"
    return f"UNRESOLVED:{safe}:{unique_row}"


def population(row: pd.Series, cfg: Dict[str, Any]) -> str:
    username = clean(row.get("Username")).lower()
    for domain in cfg["internal_email_domains"]:
        d = clean(domain).lower().lstrip("@")
        if username.endswith("@" + d):
            return "Internal staff"
    role_text = f"{clean(row.get('Roles'))} {clean(row.get('Staff Category'))}".lower()
    if any(clean(m).lower() in role_text for m in cfg["internal_role_markers"] if clean(m)):
        return "Internal staff"
    return "External partner"


def year_from(begin: Any, end: Any, completion: Any, filename: str) -> Any:
    for v in (begin, end):
        dt = pd.to_datetime(v, errors="coerce")
        if pd.notna(dt):
            return int(dt.year)
    found = re.findall(r"\b(20\d{2})\b", filename)
    if found:
        return int(found[0])
    dt = pd.to_datetime(completion, errors="coerce")
    return int(dt.year) if pd.notna(dt) else ""


def normalize(df: pd.DataFrame, path: Path, sheet: str, cfg: Dict[str, Any], offset: int) -> pd.DataFrame:
    cmap = column_map(df, cfg["column_aliases"])
    raw_fields = list(cfg["column_aliases"].keys())
    out = pd.DataFrame(index=df.index)
    for field in raw_fields:
        out[field] = df[cmap[field]] if field in cmap else ""

    out["Source File"] = path.name if sheet == "CSV" else f"{path.name} | {sheet}"

    text_fields = [x for x in raw_fields if x not in {"Assignment Begin", "Assignment End", "Completion Date", "Training Hours"}]
    for c in text_fields:
        out[c] = out[c].map(clean)

    for c in ("Assignment Begin", "Assignment End", "Completion Date"):
        out[c] = pd.to_datetime(out[c], errors="coerce")
    out["Training Hours"] = pd.to_numeric(out["Training Hours"], errors="coerce")

    out["Participant Key"] = [
        make_participant_key(u, e, n, offset + i + 1)
        for i, (u, e, n) in enumerate(zip(out["User ID"], out["Username"], out["Full Name"]))
    ]
    out["Population"] = out.apply(lambda r: population(r, cfg), axis=1)

    rules = cfg["status_rules"]
    out["Completed Flag"] = out["Completion Status"].map(lambda x: int(matches(x, rules["completed"])))
    out["Passed Flag"] = out["Result Status"].map(lambda x: int(matches(x, rules["passed"])))
    out["Completed Late Flag"] = out["Completion Status"].map(lambda x: int(matches(x, rules["late"])))

    started, not_started = [], []
    for cs, ast, done in zip(out["Completion Status"], out["Assignment Status"], out["Completed Flag"]):
        combined = f"{clean(cs)} {clean(ast)}"
        s = int(not done and matches(combined, rules["started"]))
        n = int(not done and matches(combined, rules["not_started"]))
        if s and n:
            # Explicit 'not started' wins over generic started tokens.
            s = 0
        started.append(s)
        not_started.append(n)
    out["Started Not Completed Flag"] = started
    out["Not Started Flag"] = not_started

    out["Assignment Year"] = [
        year_from(b, e, c, src)
        for b, e, c, src in zip(out["Assignment Begin"], out["Assignment End"], out["Completion Date"], out["Source File"])
    ]

    missing = []
    for _, r in out.iterrows():
        absent = [f for f in cfg["critical_fields"] if clean(r.get(f)) == ""]
        missing.append(", ".join(absent))
    out["Missing Critical Fields"] = missing
    out["Exact Repeat Review"] = ""

    for c in OUTPUT_COLUMNS:
        if c not in out:
            out[c] = ""
    return out[OUTPUT_COLUMNS + ["Completion Date"]]


def flag_repeats(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols = ["Participant Key", "Assignment ID", "Assignment Title", "Assignment Begin", "Assignment End", "Completion Status", "Result Status"]
    canon = pd.DataFrame(index=out.index)
    for c in cols:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            canon[c] = out[c].dt.strftime("%Y-%m-%d").fillna("")
        else:
            canon[c] = out[c].fillna("").map(lambda x: clean(x).lower())
    has_identity = out["Assignment ID"].map(clean).ne("") | out["Assignment Title"].map(clean).ne("")
    resolved = ~out["Participant Key"].str.startswith("UNRESOLVED:", na=False)
    dup = canon.duplicated(cols, keep=False) & has_identity & resolved
    out.loc[dup, "Exact Repeat Review"] = "Potential exact repeated assignment record - verify against source"
    return out


def metrics(g: pd.DataFrame, include_pass_rate: bool = False) -> pd.Series:
    rows = len(g)
    data = {
        "Assignment Rows": rows,
        "Unique Participants": g["Participant Key"].nunique(),
        "Completed": int(g["Completed Flag"].sum()),
        "Completion Rate": float(g["Completed Flag"].sum() / rows) if rows else 0.0,
        "Passed": int(g["Passed Flag"].sum()),
        "Started Not Completed": int(g["Started Not Completed Flag"].sum()),
        "Not Started": int(g["Not Started Flag"].sum()),
        "Completed Late": int(g["Completed Late Flag"].sum()),
    }
    if include_pass_rate:
        data["Pass Rate"] = float(g["Passed Flag"].sum() / rows) if rows else 0.0
    return pd.Series(data)


def course_summary(df: pd.DataFrame) -> pd.DataFrame:
    result = df.groupby(["Assignment ID", "Assignment Title", "Population"], dropna=False).apply(lambda g: metrics(g, True)).reset_index()
    ordered = ["Assignment ID", "Assignment Title", "Population", "Assignment Rows", "Unique Participants", "Completed", "Completion Rate", "Passed", "Pass Rate", "Started Not Completed", "Not Started", "Completed Late"]
    return result[ordered]


def year_summary(df: pd.DataFrame) -> pd.DataFrame:
    result = df.groupby(["Assignment Year", "Population"], dropna=False).apply(metrics).reset_index()
    return result[["Assignment Year", "Population", "Assignment Rows", "Unique Participants", "Completed", "Completion Rate", "Passed", "Started Not Completed", "Not Started", "Completed Late"]]


def join_unique(vals: Iterable[Any]) -> str:
    seen, result = set(), []
    for v in vals:
        s = clean(v)
        if s and s not in seen:
            seen.add(s); result.append(s)
    return "; ".join(result)


def history(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for key, g in df.groupby("Participant Key", dropna=False):
        years = sorted({int(float(y)) for y in g["Assignment Year"] if clean(y) and re.fullmatch(r"\d{4}(?:\.0)?", str(y))})
        movement = "Single assignment year" if len(years) <= 1 else " → ".join(map(str, years))
        dates = pd.concat([
            pd.to_datetime(g["Completion Date"], errors="coerce"),
            pd.to_datetime(g["Assignment Begin"], errors="coerce"),
            pd.to_datetime(g["Assignment End"], errors="coerce")
        ]).dropna()
        identity = g["Assignment ID"].map(clean).where(g["Assignment ID"].map(clean).ne(""), g["Assignment Title"].map(clean))
        rows.append({
            "Participant Key": key,
            "User ID(s)": join_unique(g["User ID"]),
            "Full Name(s)": join_unique(g["Full Name"]),
            "Username(s)": join_unique(g["Username"]),
            "Population": join_unique(g["Population"]),
            "Assignment Rows": len(g),
            "Distinct Courses": identity[identity.ne("")].nunique(),
            "Completed": int(g["Completed Flag"].sum()),
            "Passed": int(g["Passed Flag"].sum()),
            "Started Not Completed": int(g["Started Not Completed Flag"].sum()),
            "Not Started": int(g["Not Started Flag"].sum()),
            "Completed Late": int(g["Completed Late Flag"].sum()),
            "Assignment Years": ", ".join(map(str, years)),
            "Movement Classification": movement,
            "First Activity": dates.min() if not dates.empty else pd.NaT,
            "Latest Activity": dates.max() if not dates.empty else pd.NaT,
        })
    return pd.DataFrame(rows)


def exception_rows(df: pd.DataFrame) -> pd.DataFrame:
    mask = df["Missing Critical Fields"].map(clean).ne("") | df["Exact Repeat Review"].map(clean).ne("") | df["Participant Key"].str.startswith("UNRESOLVED:", na=False)
    cols = ["Source File", "Participant Key", "User ID", "Full Name", "Username", "Assignment ID", "Assignment Title", "Completion Status", "Result Status", "Missing Critical Fields", "Exact Repeat Review"]
    return df.loc[mask, cols].copy()


def methodology(cfg: Dict[str, Any]) -> pd.DataFrame:
    rows = [
        ["Unit of analysis", "One participant × one training assignment row.", "Preserve multiple legitimate trainings.", "Do not deduplicate legitimate repeat learners.", "Normalized records", "Assignments and unique people are reported separately."],
        ["Participant key", "User ID first; username/email only if User ID is absent.", "Consolidate individual history.", "Exact identifier matching only.", "User ID / Username", "No fuzzy-name merging."],
        ["Population", "Internal email domains and configured role/category markers identify internal staff.", "Separate staff from partners.", "Others default to External partner.", "Username / Roles / Staff Category", "Review config.json for your organisation."],
        ["Completed", "Completion Status matching configured Completed patterns.", "Measure completion.", "Late completion still counts as completed.", "Completion Status", ""],
        ["Passed", "Result Status matching configured Passed patterns.", "Separate pass from completion.", "Binary Passed Flag.", "Result Status", ""],
        ["Started Not Completed", "Started/in-progress and not completed.", "Identify partial completion.", "Separate binary flag.", "Completion/Assignment Status", ""],
        ["Not Started", "Explicit not-started/not-commenced status and not completed.", "Identify no commencement.", "Separate binary flag.", "Completion/Assignment Status", ""],
        ["Completed Late", "Late/out-of-deadline status pattern.", "Identify deadline breach.", "Still counted as completed.", "Completion Status", ""],
        ["Duplicate review", "Same participant, assignment identity, dates, completion status and result.", "Identify suspicious exact repeats.", "Flag only; never auto-delete.", "Assignment-level fields", "Multi-course and multi-year participation is preserved."],
        ["Missing critical fields", "Configured key fields checked for blanks.", "Surface data-quality risks.", "Retain record and list in Exceptions.", "Normalized records", "Configured in config.json."],
        ["Assignment year", "Begin year, then end year, then year in filename, then completion year.", "Protect cohort-year logic.", "One year per assignment row.", "Dates / filename", "Adjust if your LMS defines cohorts differently."]
    ]
    return pd.DataFrame(rows, columns=["Control", "Rule", "Purpose", "Treatment", "Source", "Notes"])


def validations(all_records: pd.DataFrame, courses: pd.DataFrame, years: pd.DataFrame, hist: pd.DataFrame) -> List[str]:
    errors = []
    total = len(all_records)
    if int(courses["Assignment Rows"].sum()) != total:
        errors.append("Course Reconciliation assignment rows do not equal All Records.")
    if int(years["Assignment Rows"].sum()) != total:
        errors.append("Year Movement assignment rows do not equal All Records.")
    if len(hist) != all_records["Participant Key"].nunique():
        errors.append("Participant History rows do not equal unique Participant Keys.")
    states = all_records["Completed Flag"] + all_records["Started Not Completed Flag"] + all_records["Not Started Flag"]
    if (states > 1).any():
        errors.append("Some rows have conflicting completion-state flags.")
    return errors


def executive_population(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for p, g in df.groupby("Population", dropna=False):
        total = len(g); completed = int(g["Completed Flag"].sum())
        rows.append([p, total, g["Participant Key"].nunique(), completed, completed / total if total else 0])
    return pd.DataFrame(rows, columns=["Population", "Assignments", "Participants", "Completed", "Completion Rate"])


def format_workbook(path: Path, validation_errors: List[str]) -> None:
    wb = load_workbook(path)
    dark = PatternFill("solid", fgColor="1F4E3D")
    light = PatternFill("solid", fgColor="D9EAD3")
    warning = PatternFill("solid", fgColor="FCE4D6")
    white_bold = Font(color="FFFFFF", bold=True)

    ws = wb["Executive Summary"]
    ws["A1"] = "AML Training Reconciliation"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = dark
    ws.merge_cells("A1:E1")
    ws["A2"] = "All participants included; internal staff and external partners separated; legitimate repeated training participation retained."
    ws.merge_cells("A2:E2")
    ws["A2"].alignment = Alignment(wrap_text=True)

    for s in ["Course Reconciliation", "Year Movement", "Participant History", "All Records", "Exceptions", "Methodology"]:
        sh = wb[s]
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions
        for c in sh[1]:
            c.fill = dark; c.font = white_bold; c.alignment = Alignment(wrap_text=True, vertical="center")
        if sh.max_row >= 2:
            try:
                table = Table(displayName=re.sub(r"[^A-Za-z0-9]", "", s)[:20] + "Tbl", ref=f"A1:{get_column_letter(sh.max_column)}{sh.max_row}")
                table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True, showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
                sh.add_table(table)
            except Exception:
                pass
        headers = {c.value: c.column for c in sh[1]}
        for h in ("Completion Rate", "Pass Rate"):
            if h in headers:
                for r in range(2, sh.max_row + 1): sh.cell(r, headers[h]).number_format = "0.0%"
        for h in ("Assignment Begin", "Assignment End", "First Activity", "Latest Activity"):
            if h in headers:
                for r in range(2, sh.max_row + 1): sh.cell(r, headers[h]).number_format = "yyyy-mm-dd"
        for col in range(1, sh.max_column + 1):
            h = clean(sh.cell(1, col).value)
            width = 38 if h in {"Source File", "Missing Critical Fields", "Exact Repeat Review", "Rule", "Purpose", "Treatment", "Notes"} else 30 if h in {"Full Name", "Assignment Title", "Learning Title", "Username", "Participant Key"} else min(max(len(h)+3, 12), 22)
            sh.column_dimensions[get_column_letter(col)].width = width
        for row in sh.iter_rows():
            for c in row: c.alignment = Alignment(vertical="top", wrap_text=True)

    # Executive summary headers and rates.
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value in {"Metric", "Population"}:
            for c in ws[r]:
                if c.value is not None: c.fill = light; c.font = Font(bold=True)
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Completion rate": ws.cell(r, 2).number_format = "0.0%"
        if ws.cell(r, 1).value in {"Internal staff", "External partner"}: ws.cell(r, 5).number_format = "0.0%"
    if validation_errors:
        ws["G1"] = "Validation issues"; ws["G1"].fill = warning; ws["G1"].font = Font(bold=True)
        for i, err in enumerate(validation_errors, 2): ws.cell(i, 7, err)
        ws.column_dimensions["G"].width = 60

    wb.save(path)


def process(input_dir: Path, output_path: Path, cfg: Dict[str, Any]) -> tuple[pd.DataFrame, List[str]]:
    """Generate the workbook and return records and validation notes for the GUI."""
    files = source_files(input_dir)
    if not files:
        raise FileNotFoundError(f"No Excel/CSV files found in {input_dir.resolve()}")

    frames, offset = [], 0
    print("Reading source files...")
    for path in files:
        file_count = 0
        for sheet, df in read_file(path):
            n = normalize(df, path, sheet, cfg, offset)
            meaningful = n[["User ID", "Username", "Full Name", "Assignment ID", "Assignment Title", "Completion Status"]].fillna("").astype(str).apply(lambda c: c.str.strip()).ne("").any(axis=1)
            n = n.loc[meaningful].copy()
            if not n.empty:
                frames.append(n); offset += len(n); file_count += len(n)
        print(f"  {path.name}: {file_count:,} records")

    if not frames:
        raise ValueError("Files were readable but no meaningful training records were detected.")

    all_records = flag_repeats(pd.concat(frames, ignore_index=True))
    courses = course_summary(all_records)
    years = year_summary(all_records)
    hist = history(all_records)
    exc = exception_rows(all_records)
    meth = methodology(cfg)
    errs = validations(all_records, courses, years, hist)
    pop = executive_population(all_records)

    total = len(all_records); completed = int(all_records["Completed Flag"].sum())
    kpis = pd.DataFrame([
        ["Assignment rows", total],
        ["Unique participants", all_records["Participant Key"].nunique()],
        ["Internal staff", all_records.loc[all_records["Population"] == "Internal staff", "Participant Key"].nunique()],
        ["External partners", all_records.loc[all_records["Population"] == "External partner", "Participant Key"].nunique()],
        ["Completed", completed],
        ["Completion rate", completed / total if total else 0],
        ["Passed", int(all_records["Passed Flag"].sum())],
        ["Started but not completed", int(all_records["Started Not Completed Flag"].sum())],
        ["Not started", int(all_records["Not Started Flag"].sum())],
        ["Completed late", int(all_records["Completed Late Flag"].sum())],
        ["Exception records", len(exc)],
        ["Validation status", "PASS" if not errs else "REVIEW REQUIRED"]
    ], columns=["Metric", "Value"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        kpis.to_excel(writer, sheet_name="Executive Summary", index=False, startrow=3)
        pop.to_excel(writer, sheet_name="Executive Summary", index=False, startrow=3 + len(kpis) + 3)
        courses.to_excel(writer, sheet_name="Course Reconciliation", index=False)
        years.to_excel(writer, sheet_name="Year Movement", index=False)
        hist.to_excel(writer, sheet_name="Participant History", index=False)
        all_records[OUTPUT_COLUMNS].to_excel(writer, sheet_name="All Records", index=False)
        exc.to_excel(writer, sheet_name="Exceptions", index=False)
        meth.to_excel(writer, sheet_name="Methodology", index=False)

    format_workbook(output_path, errs)

    print("\nResults")
    print(f"  Assignment rows:     {len(all_records):,}")
    print(f"  Unique participants: {all_records['Participant Key'].nunique():,}")
    print(f"  Completed:           {int(all_records['Completed Flag'].sum()):,}")
    print(f"  Passed:              {int(all_records['Passed Flag'].sum()):,}")
    print(f"  Exceptions:          {len(exc):,}")
    print(f"  Output:              {output_path.resolve()}")
    print("  Validation:          " + ("PASS" if not errs else "REVIEW REQUIRED"))
    if errs:
        for e in errs: print("    - " + e)
    return all_records, errs


def build(input_dir: Path, output_path: Path, cfg: Dict[str, Any]) -> List[str]:
    """Keep the command-line entry point's validation-only return contract."""
    _, errs = process(input_dir, output_path, cfg)
    return errs


def main() -> int:
    p = argparse.ArgumentParser(description="Offline AML/CFT training reconciliation")
    p.add_argument("--input", default="./input", help="Folder containing raw Excel/CSV exports")
    p.add_argument("--output", default="./output/AML_Training_Full_Reconciliation.xlsx", help="Output workbook path")
    p.add_argument("--config", default=None, help="Optional JSON config")
    args = p.parse_args()
    try:
        errs = build(Path(args.input), Path(args.output), load_config(Path(args.config) if args.config else None))
        return 1 if errs else 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
