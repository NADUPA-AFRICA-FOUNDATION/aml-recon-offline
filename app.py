from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from aml_reconcile import DEFAULT_CONFIG, process

APP_TITLE = "AML Training Reconciliation Studio"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    "<style>" + (Path(__file__).parent / "assets" / "studio.css").read_text() + "</style>",
    unsafe_allow_html=True,
)


def init_state() -> None:
    defaults = {
        "result_bytes": None,
        "result_name": "AML_Training_Full_Reconciliation.xlsx",
        "summary": None,
        "exceptions_df": None,
        "validation_errors": None,
        "result_signature": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()

workflow = st.empty()
workflow.markdown(
    """
<div class="hero">
  <div><div class="eyebrow">Compliance workspace / AML &amp; CFT</div>
  <h1>Reconcile.<br>Prove it.</h1>
  <p>Drop the exports. Challenge every row. Leave with one loud, traceable source of truth.</p></div>
  <div class="local-badge">● 100% local</div>
</div>
<div class="ticker" aria-hidden="true"><span>RAW FILES ✦ CONTROL TOTALS ✦ EXCEPTIONS ✦ AUDIT TRAIL ✦ NO CLOUD ✦ RAW FILES ✦ CONTROL TOTALS ✦ EXCEPTIONS ✦ AUDIT TRAIL ✦ NO CLOUD ✦&nbsp;</span></div>
<div class="workflow" aria-label="Reconciliation workflow">
  <div class="step"><b>01</b> Add exports</div>
  <div class="step"><b>02</b> Review results</div>
  <div class="step"><b>03</b> Export workbook</div>
</div>
""", unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="brand"><span class="brand-mark">AML</span><div>Reconcile Studio<small>TRAINING OPERATIONS</small></div></div>', unsafe_allow_html=True)
    st.subheader("Classification rules")
    st.caption("Set how learners are grouped in your report.")

    internal_domains_raw = st.text_area(
        "Internal email domains",
        value="safaricom.co.ke",
        height=76,
        help="One domain per line, e.g. safaricom.co.ke",
    )
    internal_markers_raw = st.text_area(
        "Internal role markers",
        value="safaricom\nemployee\nstaff\ninternal",
        height=110,
        help="If Roles or Staff Category contains one of these markers, the learner is classified as internal.",
    )

    st.divider()
    st.markdown("**Matching policy**")
    st.caption("1. User ID\n\n2. Username/email fallback\n\n3. Never fuzzy-match by name")

    st.markdown("**Duplicate policy**")
    st.caption("Multiple courses and multi-year participation are preserved. Only suspicious exact assignment repeats are flagged.")

    st.divider()
    st.caption("Review the Exceptions tab after each run before sharing your workbook.")

upload_col, guide_col = st.columns([1.55, 1], gap="large")

with upload_col, st.container(border=True):
    st.subheader("01 / Drop the evidence")
    st.caption("Multiple courses. Multiple years. One uncompromised audit trail.")
    uploaded_files = st.file_uploader(
        "Drop your LMS exports here",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    st.markdown(
        '<div class="small-note">Supported: XLSX, XLSM, XLS and CSV. You can add multiple years, courses and populations in one run.</div>',
        unsafe_allow_html=True,
    )

with guide_col, st.container(border=True):
    st.subheader("02 / Know the output")
    st.caption("Seven connected worksheets. Every number has somewhere to answer to.")
    st.markdown("""
<div class="report-row"><span class="report-num">01</span><div><strong>Management overview</strong><p>Executive Summary, Course Reconciliation and Year Movement</p></div></div>
<div class="report-row"><span class="report-num">02</span><div><strong>Every learner, every record</strong><p>Participant History and All Records</p></div></div>
<div class="report-row"><span class="report-num">03</span><div><strong>Controls you can inspect</strong><p>Exceptions and Methodology</p></div></div>
""", unsafe_allow_html=True)

st.divider()

run_left, run_right = st.columns([1.2, 2.8])
with run_left:
    run_clicked = st.button(
        "Run the controls →",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files,
    )
with run_right:
    if uploaded_files:
        total_size = sum(getattr(f, "size", 0) for f in uploaded_files)
        st.info(f"Ready: **{len(uploaded_files)} file(s)** • {total_size / (1024*1024):.1f} MB")
    else:
        st.info("Add at least one training export to enable the reconciliation engine.")

current_signature = (
    tuple((file.name, hashlib.sha256(file.getbuffer()).hexdigest()) for file in (uploaded_files or [])),
    internal_domains_raw,
    internal_markers_raw,
)

if st.session_state.summary and st.session_state.result_signature != current_signature:
    st.session_state.result_bytes = None
    st.session_state.summary = None
    st.session_state.exceptions_df = None
    st.session_state.validation_errors = None
    st.info("Inputs or classification rules changed. Run reconciliation again to refresh your results.")

if run_clicked:
    st.session_state.result_bytes = None
    st.session_state.summary = None
    st.session_state.exceptions_df = None
    st.session_state.validation_errors = None

    status = st.status("Preparing local reconciliation…", expanded=True)
    try:
        with tempfile.TemporaryDirectory(prefix="aml_recon_") as tmp:
            tmpdir = Path(tmp)
            input_dir = tmpdir / "input"
            output_dir = tmpdir / "output"
            input_dir.mkdir()
            output_dir.mkdir()

            status.write("Saving uploaded exports to a temporary local workspace…")
            for index, file in enumerate(uploaded_files):
                dest = input_dir / f"{index + 1:03d}_{Path(file.name).name}"
                dest.write_bytes(file.getbuffer())

            cfg = json.loads(json.dumps(DEFAULT_CONFIG))
            cfg["internal_email_domains"] = [
                x.strip().lower().lstrip("@")
                for x in internal_domains_raw.splitlines()
                if x.strip()
            ]
            cfg["internal_role_markers"] = [
                x.strip().lower()
                for x in internal_markers_raw.splitlines()
                if x.strip()
            ]

            cfg_path = tmpdir / "config.json"
            cfg_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
            output_path = output_dir / st.session_state.result_name

            status.write("Normalizing source columns and building participant keys…")
            all_records, validation_errors = process(
                input_dir=input_dir,
                output_path=output_path,
                cfg=cfg,
            )

            status.write("Calculating completion, pass, late, movement and exception controls…")
            exception_mask = (
                all_records["Missing Critical Fields"].fillna("").astype(str).str.strip().ne("")
                | all_records["Exact Repeat Review"].fillna("").astype(str).str.strip().ne("")
                | all_records["Participant Key"].fillna("").astype(str).str.startswith("UNRESOLVED:")
            )
            exceptions = all_records.loc[exception_mask].copy()

            summary = {
                "assignments": len(all_records),
                "participants": int(all_records["Participant Key"].nunique()),
                "internal": int(all_records.loc[all_records["Population"] == "Internal staff", "Participant Key"].nunique()),
                "external": int(all_records.loc[all_records["Population"] == "External partner", "Participant Key"].nunique()),
                "completed": int(all_records["Completed Flag"].sum()),
                "passed": int(all_records["Passed Flag"].sum()),
                "started_not_completed": int(all_records["Started Not Completed Flag"].sum()),
                "not_started": int(all_records["Not Started Flag"].sum()),
                "late": int(all_records["Completed Late Flag"].sum()),
                "exceptions": int(exception_mask.sum()),
                "completion_rate": float(all_records["Completed Flag"].mean()) if len(all_records) else 0.0,
            }

            st.session_state.result_signature = current_signature
            st.session_state.result_bytes = output_path.read_bytes()
            st.session_state.summary = summary
            st.session_state.exceptions_df = exceptions
            st.session_state.validation_errors = validation_errors

            if validation_errors:
                status.update(label="Reconciliation completed — review required", state="complete", expanded=False)
            else:
                status.update(label="Reconciliation completed successfully", state="complete", expanded=False)

    except Exception as exc:
        status.update(label="Reconciliation could not be completed", state="error", expanded=True)
        st.error(str(exc))

if st.session_state.summary:
    s = st.session_state.summary
    st.divider()
    st.subheader("03 / Control room")

    if st.session_state.validation_errors:
        st.warning("Validation needs review. Inspect the Validation and Exceptions tabs before sharing the workbook.")
    else:
        st.success("Validation passed — control totals reconcile across the workbook.")

    st.write("")
    metric_cols = st.columns(3)
    metrics = [
        ("Assignment rows", f"{s['assignments']:,}"),
        ("Unique participants", f"{s['participants']:,}"),
        ("Completed", f"{s['completed']:,}"),
        ("Completion rate", f"{s['completion_rate']:.1%}"),
        ("Passed", f"{s['passed']:,}"),
        ("Exceptions", f"{s['exceptions']:,}"),
    ]
    for index, (label, value) in enumerate(metrics):
        metric_cols[index % 3].metric(label, value)

    tab_overview, tab_population, tab_exceptions, tab_validation = st.tabs(
        ["Overview", "Population", "Exceptions", "Validation"]
    )

    with tab_overview:
        c1, c2, c3 = st.columns(3)
        c1.metric("Started, not completed", f"{s['started_not_completed']:,}")
        c2.metric("Not started", f"{s['not_started']:,}")
        c3.metric("Completed late", f"{s['late']:,}")
        st.caption("Late completions remain counted as completed and are separately flagged for analysis.")

    with tab_population:
        pop_df = pd.DataFrame(
            {
                "Population": ["Internal staff", "External partner"],
                "Unique participants": [s["internal"], s["external"]],
            }
        )
        st.dataframe(pop_df, use_container_width=True, hide_index=True)

    with tab_exceptions:
        exc_df = st.session_state.exceptions_df
        if exc_df is None or exc_df.empty:
            st.success("No reconciliation exceptions were detected under the configured rules.")
        else:
            display_cols = [
                "Source File", "Participant Key", "Full Name", "Username",
                "Assignment Title", "Completion Status", "Result Status",
                "Missing Critical Fields", "Exact Repeat Review",
            ]
            existing = [c for c in display_cols if c in exc_df.columns]
            query = st.text_input("Search exceptions", placeholder="Participant, course, source file or issue")
            filtered = exc_df[existing]
            if query.strip():
                matches = filtered.fillna("").astype(str).apply(
                    lambda column: column.str.contains(query.strip(), case=False, regex=False)
                ).any(axis=1)
                filtered = filtered.loc[matches]
            st.caption(f"Showing {len(filtered):,} of {len(exc_df):,} exception records")
            st.dataframe(filtered, use_container_width=True, hide_index=True, height=420)
            st.caption("Exceptions are retained in the final workbook. Nothing is silently deleted.")

    with tab_validation:
        if st.session_state.validation_errors:
            for item in st.session_state.validation_errors:
                st.warning(item)
        else:
            st.success("Course totals = All Records totals")
            st.success("Year totals = All Records totals")
            st.success("Population totals = All Records totals")
            st.success("Participant History = unique Participant Keys")
            st.success("Completion-state logic is internally consistent")

    st.divider()
    st.subheader("04 / Take the evidence")
    d1, d2 = st.columns([1.2, 2.8])
    with d1:
        st.download_button(
            "Download reconciled workbook",
            data=st.session_state.result_bytes,
            file_name=st.session_state.result_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with d2:
        st.success("The output contains the full seven-sheet AML training reconciliation and audit trail.")

else:
    st.markdown('<div class="empty-state"><h3>Your results will appear here</h3><p>Add your exports, check the classification rules, then run reconciliation to see your training overview.</p></div>', unsafe_allow_html=True)

st.divider()
with st.expander("Methodology & matching controls"):
    st.markdown(
        """
The interface runs the same deterministic offline reconciliation rules every time:

1. **Normalize** equivalent LMS columns across every uploaded file.
2. Build a **Participant Key** from User ID, then username/email as fallback.
3. Classify **Internal staff vs External partner** using the configured domain/role rules.
4. Calculate **Completed, Passed, Started Not Completed, Not Started and Completed Late** flags.
5. Preserve legitimate **multiple courses and multiple years** for the same learner.
6. Flag only suspicious **exact repeated assignment records** for review; nothing is automatically deleted.
7. Build course, year and participant-level reconciliation tables.
8. Run cross-sheet **control-total validation** before producing the workbook.

No generative AI is required during processing, which means the same source files and same configuration produce the same reconciliation logic offline.
"""
    )

st.markdown('<div class="footer-note">AML Reconcile Studio · Processed locally · Source records preserved · No cloud upload</div>', unsafe_allow_html=True)
