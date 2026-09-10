# Design Review: Reconcile Studio

Date: 9 September 2026. Baseline: `89105ed`.
Method: [Apple Design Skill](https://github.com/dickwu/apple-design-skill/blob/main/SKILL.md),
using its accessibility, colour, layout, typography, data entry, feedback and
search references. This is a source-code and application-state audit of the
Streamlit desktop/mobile web app, followed by implemented fixes. It is not a
visual browser audit or an accessibility certification.

## Summary

**Assessment: Good after remediation; visual validation still needed.** The green
token system and native controls provided a sound foundation. The most important
issues were task ordering on mobile, unvalidated classification settings, and
recovery from empty search results and failed imports. Those paths now have
clearer controls and feedback.

## Critical issues

No critical issue was established by the checks available in this session.
Keyboard-only, screen-reader, zoom and device-specific failures cannot be ruled
out: no connected browser was available. These are outstanding verification
items, not passing checks.

## Improvements implemented

### High · Validate classification settings before a run

**What:** A full email address, URL or malformed domain could be accepted and fail
to match internal staff. The resulting population split could look plausible.

**Why:** Errors should be caught where users enter data, before they commit to an
action. [Entering Data — Best practices](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/entering-data.md).

**Fix:** `studio_inputs.py` accepts pasted newline/comma/semicolon domain lists,
normalises case and leading @, removes duplicates, and reports invalid entries.
The run action is disabled until invalid domains are corrected. Empty rules
explicitly explain that all learners will be classified as external. Tests check
invalid input, recovery, and the resulting internal population.

### High · Keep mobile actions beside their inputs

**What:** Classification settings lived in an automatically hidden sidebar. The
run action followed both upload and explanatory columns, so secondary content
came before the action on narrow screens.

**Why:** Related controls should be grouped, and the main task should precede
supporting content. [Layout — Visual hierarchy](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/layout.md).

**Fix:** Settings, a rules summary, readiness text and the run button now share
the upload card. Report guidance follows that card in DOM order. Sidebar content
is limited to preferences and policy information. The heading names the task
directly, and viewport safe-area insets are respected.

### High · Offer readable alternatives and stronger contrast

**What:** Exceptions were exposed only through a wide interactive data grid;
there was no explicit increased-contrast preference. Section headings skipped
from h1 to h3, and the skip link targeted the introduction rather than inputs.

**Why:** Content should remain perceivable with different input and reading
methods. [Accessibility — Vision and Mobility](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/accessibility.md),
[Color — Best practices](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/color.md).

**Fix:** Added selectable individual exception details rendered as labelled text;
population figures use a native static table. Section headings now use h2 and the
skip link targets the upload area. A session toggle and `prefers-contrast: more`
share stronger text/boundary tokens and reduced shadows. Native keyboard
controls, focus rings and reduced-motion rules remain intact. Actual assistive
technology behaviour still requires browser testing.

### Medium · Make search failure recoverable

**What:** A search with no matches produced an empty grid without a recovery
control. Users could not narrow exceptions by issue type.

**Why:** Search scope should be apparent and easy to change.
[Searching — Best practices](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/searching.md).

**Fix:** Added Missing fields, Possible repeats and Unresolved identity filters;
a scoped result count; explicit no-match guidance; and Clear filters. The preview
states that filtering does not alter the workbook. Tests verify literal search,
combined filters, recovery and unchanged workbook bytes.

### Medium · Match feedback to its significance

**What:** A run could show passing totals alongside a warning, followed by five
more success alerts and another success-styled download description. Failed
imports displayed raw exception text without a stable recovery message.

**Why:** Excessive alerts reduce their usefulness; failures need actionable next
steps. [Feedback — Best practices](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/feedback.md).

**Fix:** One overall success/review message establishes the run outcome. Detailed
passing checks appear as a quiet checklist. Import failure includes a next step
and an expanded Import details section. Readiness and export metadata use
unobtrusive supporting text.

### Low · Reduce decorative competition

**What:** Large promotional copy and equally strong shadows made supporting
content compete with the main task. Metrics were inserted column-first, causing
a different sequence on narrow screens.

**Why:** Size, weight and placement should reflect information priority.
[Typography — Conveying hierarchy](https://github.com/dickwu/apple-design-skill/blob/main/references/hig/typography.md).

**Fix:** A shorter functional title, system font stack, lighter shadow tokens and
row-wise metrics improve consistency. Text can wrap at large sizes. Supporting
report content loses elevation on small screens; primary controls retain depth.

## Positive notes

- Original green identity, semantic tokens and readable numerical summaries.
- Native Streamlit controls preserve framework interaction behaviour.
- Invalidated results cannot be downloaded after source/rule changes.
- Imported records, flagged repeats and full workbook output are preserved.
- No remote fonts, proprietary brand assets or new runtime dependencies.

## Platform-specific notes and remaining checks

The skill's native-app recommendations are translated to web conventions: native
file selection, disclosure controls and tabs are retained. No artificial bottom
navigation or payment controls were introduced. The existing light appearance is
preserved; a complete dark appearance remains outside this change and is a known
gap against the skill's broader appearance guidance.

Before claiming full accessibility or production visual approval, verify:

- 375px, 768px and desktop layouts, portrait/landscape, and 200% text/zoom.
- Keyboard traversal through upload, rules, filters, tabs and download.
- Screen-reader announcements of errors, details, current step and table content.
- Focus visibility in collapsed/expanded sections and high/forced contrast modes.
- Reduced motion and system increased contrast in actual browsers.

Automated coverage: application flows, classification validation/recovery,
exception filtering, import formats, verifier, launcher, and both token palettes.
Contrast tests cover solid token pairs; they do not substitute for rendered-page
contrast measurement. The native Streamlit DOM selectors also need regression
review after framework upgrades.
