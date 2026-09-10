from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from aml_reconcile import DEFAULT_CONFIG, process
import studio_ui as ui
from studio_inputs import ISSUE_TYPES, filter_exceptions, parse_domains, parse_markers

APP_TITLE = "AML Training Reconciliation Studio"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=":material/verified_user:",
    layout="wide",
    initial_sidebar_state="auto",
)

st.markdown(
    "<style>" + "\n".join(
        (Path(__file__).parent / "assets" / name).read_text()
        for name in ("tokens.css", "studio.css")
    ) + "\n@media (prefers-contrast: more) {"
    + (Path(__file__).parent / "assets" / "high-contrast.css").read_text()
    + "}\n</style>",
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
        "export_requested": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()

st.markdown(ui.header(), unsafe_allow_html=True)
workflow_slot = st.empty()
workflow_slot.markdown(ui.workflow("upload"), unsafe_allow_html=True)

with st.sidebar:
    st.markdown(ui.brand(), unsafe_allow_html=True)
    st.header("Workspace preferences")
    increased_contrast = st.toggle(
        "Increase contrast", key="increased_contrast",
        help="Stronger text and boundaries with less shadow. Does not change your reconciliation rules.",
    )
    st.caption("Your preference applies to this session.")
    st.divider()
    st.markdown("**Matching policy**")
    st.caption("User ID first, then username or email. Names are never fuzzy-matched.")
    st.markdown("**Source records stay intact**")
    st.caption("Multiple courses and years are preserved. Possible repeats are flagged for review.")

if increased_contrast:
    st.markdown(
        "<style>" + (Path(__file__).parent / "assets" / "high-contrast.css").read_text() + "</style>",
        unsafe_allow_html=True,
    )

upload_col, guide_col = st.columns([1.55, 1], gap="large")

with upload_col, st.container(key="upload-card"):
    st.markdown('<span id="workspace" tabindex="-1"></span>', unsafe_allow_html=True)
    st.header("Add training exports")
    st.caption("Combine multiple courses and years in one report.")
    uploaded_files = st.file_uploader(
        "Choose Excel or CSV files",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        label_visibility="visible",
    )
    st.caption("XLSX, XLSM, XLS or CSV. Report titles above column headings are supported.")

    with st.container(key="classification-rules"):
        st.subheader("Classification rules")
        st.caption("A matching email domain or role marks a learner as internal. Other learners are classified as external partners.")
        internal_domains_raw = st.text_area(
            "Internal email domains", key="internal_domains",
            value="\n".join(DEFAULT_CONFIG["internal_email_domains"]), height=104,
            help="Enter domains such as safaricom.co.ke, separated by new lines, commas or semicolons. Do not enter full email addresses or URLs.",
        )
        st.caption("Email domains only, for example safaricom.co.ke. Separate entries with new lines or commas.")
        internal_markers_raw = st.text_area(
            "Internal role markers", key="internal_markers",
            value="\n".join(DEFAULT_CONFIG["internal_role_markers"]), height=144,
            help="One marker per line. A match anywhere in Roles or Staff Category classifies the learner as internal.",
        )
        st.caption("Role keywords only, one per line, for example employee or staff.")
        domains, domain_errors = parse_domains(internal_domains_raw)
        markers = parse_markers(internal_markers_raw)
        if domain_errors:
            st.error("Some domains are invalid. Enter domain names only, without email addresses, spaces or https://.", icon=":material/error:")
            st.text("Check: " + ", ".join(domain_errors))
        elif not domains and not markers:
            st.warning("With both rules empty, every learner will be classified as an external partner.", icon=":material/warning:")

    st.caption(f"Current rules · Email domains: {len(domains)} · Role markers: {len(markers)}")
    current_signature = (
        tuple((file.name, hashlib.sha256(file.getbuffer()).hexdigest()) for file in (uploaded_files or [])),
        tuple(domains), tuple(markers), tuple(domain_errors),
    )
    if st.session_state.summary and st.session_state.result_signature != current_signature:
        st.session_state.export_requested = False
        st.session_state.result_bytes = None
        st.session_state.summary = None
        st.session_state.exceptions_df = None
        st.session_state.validation_errors = None
        st.info("Inputs or classification rules changed. Run reconciliation again to refresh your results.", icon=":material/info:")

    if domain_errors:
        st.caption("Fix the email domains in Classification rules to continue.")
    elif uploaded_files:
        total_size = sum(getattr(file, "size", 0) for file in uploaded_files)
        st.caption(f"Ready to reconcile · Files: {len(uploaded_files)} · {total_size / (1024 * 1024):.1f} MB")
    else:
        st.caption("Add at least one export to enable reconciliation.")
    run_clicked = st.button(
        "Run full reconciliation",
        type="secondary" if st.session_state.summary else "primary",
        width="stretch",
        disabled=not uploaded_files or bool(domain_errors),
    )

with guide_col, st.container(key="report-card"):
    st.header("Your report includes")
    st.markdown(ui.report_guide(), unsafe_allow_html=True)
    with st.expander("Help preparing your exports"):
        st.markdown(
            "Use the original LMS export with its column headings. Common headings include "
            "**User ID**, **Email**, **Assignment Title** and **Completion Status**. "
            "If an import fails, open **Import details** to see the detected headings."
        )
        st.caption("Files are processed on the computer or server hosting this app. Temporary processing files are deleted after each run.")


if run_clicked:
    st.session_state.export_requested = False
    st.session_state.result_bytes = None
    st.session_state.summary = None
    st.session_state.exceptions_df = None
    st.session_state.validation_errors = None

    status = st.status("Reconciling your training records…", expanded=True)
    try:
        with tempfile.TemporaryDirectory(prefix="aml_recon_") as tmp:
            tmpdir = Path(tmp)
            input_dir = tmpdir / "input"
            output_dir = tmpdir / "output"
            input_dir.mkdir()
            output_dir.mkdir()

            status.write("Preparing your uploaded exports…")
            for index, file in enumerate(uploaded_files):
                dest = input_dir / f"{index + 1:03d}_{Path(file.name).name}"
                dest.write_bytes(file.getbuffer())

            cfg = json.loads(json.dumps(DEFAULT_CONFIG))
            cfg["internal_email_domains"] = domains
            cfg["internal_role_markers"] = markers

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
                | all_records["Status Review"].fillna("").astype(str).str.strip().ne("")
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

        st.rerun()

    except Exception as exc:
        status.update(label="Reconciliation could not be completed", state="error", expanded=True)
        st.error("These exports could not be reconciled.", icon=":material/error:")
        st.write("Check that each file contains training records and column headings, then replace the affected file and try again.")
        with st.expander("Import details", expanded=True):
            st.text(str(exc))

if st.session_state.summary:
    s = st.session_state.summary
    workflow_slot.markdown(ui.workflow("export" if st.session_state.export_requested else "review"), unsafe_allow_html=True)
    st.divider()
    st.header("Reconciliation results")

    if st.session_state.validation_errors or s["exceptions"]:
        st.warning(
            f"Review required · {s['exceptions']:,} exception records · "
            f"{len(st.session_state.validation_errors or []):,} validation issues. See Exceptions and Validation before sharing.",
            icon=":material/warning:",
        )
    else:
        st.success("Reconciliation complete. Control totals passed and no exceptions were detected.", icon=":material/check_circle:")

    st.markdown(ui.result_summary(s, bool(st.session_state.validation_errors or s["exceptions"])), unsafe_allow_html=True)
    metrics = [
        ("Assignment rows", f"{s['assignments']:,}"),
        ("Unique participants", f"{s['participants']:,}"),
        ("Completed", f"{s['completed']:,}"),
        ("Source files", f"{len(uploaded_files):,}"),
        ("Passed", f"{s['passed']:,}"),
        ("Exceptions", f"{s['exceptions']:,}"),
    ]
    for row_start in range(0, len(metrics), 3):
        for column, (label, value) in zip(st.columns(3), metrics[row_start:row_start + 3]):
            column.metric(label, value)

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
        st.table(pop_df)
        st.caption("Participants may appear in both populations if their source records use different email addresses or roles.")

    with tab_exceptions:
        exc_df = st.session_state.exceptions_df
        if exc_df is None or exc_df.empty:
            st.success("No reconciliation exceptions were detected under the configured rules.", icon=":material/check_circle:")
        else:
            display_cols = [
                "Source File", "Participant Key", "Full Name", "Username",
                "Assignment Title", "Completion Status", "Result Status",
                "Missing Critical Fields", "Exact Repeat Review", "Status Review",
            ]
            existing = [c for c in display_cols if c in exc_df.columns]
            query = st.text_input(
                "Search exceptions", placeholder="Participant, course, source file or issue", key="exception_query",
            )
            issue = st.selectbox("Issue type", ISSUE_TYPES, key="exception_issue")
            filtered = filter_exceptions(exc_df[existing], query, issue)
            st.caption(f"Showing {len(filtered):,} of {len(exc_df):,} exception records · {issue}")
            if query or issue != ISSUE_TYPES[0]:
                def clear_filters():
                    st.session_state.exception_query = ""
                    st.session_state.exception_issue = ISSUE_TYPES[0]
                st.button("Clear filters", on_click=clear_filters)
            if filtered.empty:
                st.info("No exceptions match these filters. Clear the filters or try a different participant or course.", icon=":material/info:")
            else:
                st.dataframe(filtered, use_container_width=True, hide_index=True, height=420)
                with st.expander("Read individual exception details"):
                    record = st.selectbox(
                        "Exception record", range(len(filtered)),
                        format_func=lambda index: f"Record {index + 1} of {len(filtered)}",
                    )
                    for field, value in filtered.iloc[record].items():
                        st.markdown(f"**{field}**")
                        st.text("Not provided" if pd.isna(value) or str(value).strip() == "" else str(value))
            st.caption("Filters only change this preview. Every exception remains in the final workbook.")

    with tab_validation:
        if st.session_state.validation_errors:
            for item in st.session_state.validation_errors:
                st.warning(item, icon=":material/warning:")
        else:
            st.markdown("**Control checks passed**")
            st.markdown(
                "- Course totals match all records.\n"
                "- Year totals match all records.\n"
                "- Participant history matches unique participant keys.\n"
                "- Completion and pass states are internally consistent."
            )

    st.divider()
    with st.container(key="export-card"):
        st.header("Export your report")
        d1, d2 = st.columns([1.2, 2.8])
        with d1:
            st.download_button(
                "Download reconciled workbook",
                data=st.session_state.result_bytes,
                file_name=st.session_state.result_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
                on_click=lambda: st.session_state.update(export_requested=True),
            )
        with d2:
            st.caption("Excel workbook · 7 worksheets · Includes all records and exceptions")
            st.caption("Filters in the Exceptions tab do not change this download.")

else:
    st.markdown(ui.empty_state(), unsafe_allow_html=True)

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

st.markdown('<div class="footer-note">Reconcile Studio · Source records preserved · Rule-based matching</div>', unsafe_allow_html=True)

st.caption("Processing happens on the computer or server running this app. Temporary upload files are deleted after each run.")
