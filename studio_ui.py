"""Small presentation components for the native Streamlit workspace.

Keep HTML display-only. Forms, actions, tables and status messages remain native
Streamlit widgets. Escape dynamic text at the HTML boundary.
"""
from html import escape


_ICONS = {
    "shield": '<path d="M12 3 4 6v6c0 5 8 9 8 9s8-4 8-9V6l-8-3Z"/><path d="m8 12 3 3 5-6"/>',
    "report": '<path d="M14 3H5v18h14V8l-5-5Z"/><path d="M14 3v5h5M8 12h8M8 16h5"/>',
    "people": '<circle cx="9" cy="8" r="3"/><path d="M3 21v-3a6 6 0 0 1 12 0v3M16 5a3 3 0 0 1 0 6M17 15a5 5 0 0 1 4 5"/>',
    "check": '<path d="m5 12 4 4L19 6"/>',
    "upload": '<path d="M12 16V3m-5 5 5-5 5 5M4 15v6h16v-6"/>',
}


def icon(name: str) -> str:
    return ('<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" '
            'aria-hidden="true" focusable="false">' + _ICONS[name] + '</svg>')


def brand() -> str:
    return (f'<div class="brand"><span class="brand-mark">{icon("shield")}</span>'
            '<div>Reconcile Studio<small>AML TRAINING OPERATIONS</small></div></div>')


def header() -> str:
    return f'''<a class="skip-link" href="#workspace">Skip to workspace</a>
<div class="hero">
<div><div class="eyebrow">Clarity in every record</div>
<h1>Training reconciliation</h1>
<p>Combine your exports, review completion and resolve exceptions.</p></div>
<div class="trust-note"><span class="icon-well">{icon("shield")}</span>
<div><strong>Your records, preserved.</strong>Clear matching rules.<br>Every exception visible.</div></div>
</div>'''


def workflow(stage: str) -> str:
    steps = [('upload', 'Add exports'), ('review', 'Review results'), ('export', 'Export workbook')]
    items = []
    for index, (key, label) in enumerate(steps, 1):
        active = key == stage
        current = ' aria-current="step"' if active else ''
        detail = '<small>Current step</small>' if active else ''
        items.append(f'<li class="step{" active" if active else ""}"{current}>'
                     f'<b aria-hidden="true">{index:02d}</b><span>{label}{detail}</span></li>')
    return '<ol class="workflow" aria-label="Reconciliation workflow">' + ''.join(items) + '</ol>'


def report_guide() -> str:
    rows = [
        ('report', 'The management picture', 'Executive Summary, Course Reconciliation and Year Movement'),
        ('people', 'Every learner, every record', 'Participant History and All Records'),
        ('shield', 'A clear trail of evidence', 'Exceptions and Methodology'),
    ]
    return ''.join(
        f'<div class="report-row"><span class="icon-well">{icon(symbol)}</span>'
        f'<div><strong>{title}</strong><p>{description}</p></div></div>'
        for symbol, title, description in rows
    ) + '<div class="report-foot">7 worksheets. One connected report.</div>'


def result_summary(summary: dict, needs_review: bool) -> str:
    title = 'Review before sharing' if needs_review else 'Control totals reconciled'
    note = ('Check the Exceptions and Validation tabs before exporting.' if needs_review
            else 'Review the results below, then download your workbook.')
    return f'''<section class="result-hero" aria-label="Training completion summary">
<div><p>Training completion</p><div class="amount">{summary['completion_rate']:.1%}</div>
<p>{summary['completed']:,} of {summary['assignments']:,} assignments completed</p></div>
<div class="summary-detail">{icon('report' if needs_review else 'check')}
<strong>{escape(title)}</strong><p>{escape(note)}</p></div></section>'''


def empty_state() -> str:
    return (f'<div class="empty-state"><span class="icon-well">{icon("upload")}</span>'
            '<h2>Your results will appear here</h2><p>Add your exports, check the classification rules, '
            'then run reconciliation to see your training overview.</p></div>')
