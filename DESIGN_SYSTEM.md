# Reconcile Studio visual system

An original East African mobile-money-inspired visual language applied to AML
training reconciliation. No M-PESA logos, assets, payment flows or affiliation
claims are used.

## Architecture

- `assets/tokens.css`: semantic colours, type, spacing, radii, elevation, focus,
  motion and breakpoint reference values.
- `assets/studio.css`: native Streamlit widget styling and layout rules.
- `studio_ui.py`: display-only brand, header, workflow, report guide, completion
  summary and empty-state components. Dynamic HTML text is escaped.
- `.streamlit/config.toml`: native widget theme. Its primary/background/text
  values mirror the tokens because Streamlit cannot consume CSS variables here.
- `app.py`: native interactive controls and the existing reconciliation flow.

Do not add remote fonts or icon dependencies: Inter is preferred when installed,
with Streamlit's existing Source Sans and system fonts as fallbacks. SVG icons
are original, decorative and hidden from assistive technology.

## Material and hierarchy

The base is #E8F0EB. Raised cards use paired light/dark shadows; fields and the
empty-results well use inset shadows. High-priority upload content uses #F8FAF9.
Deep green is reserved for the primary action and completion summary. Bright
#00A651 supplies brand energy rather than covering entire sections.

Normal-size white text on #00A651 does not reach AA. Use #007A3D for white-labelled
buttons and #006332 on hover. Success, pending/review, error and information must
always have explicit wording; preserve Streamlit's native status semantics.
Do not present passing workbook totals as proof that all exceptions are resolved.

Buttons are 52px tall; secondary controls have a 44px minimum. Form boundaries
and keyboard focus rings remain visible even where shadows are subtle. Use
borders when needed for control identification and forced-colour modes.

## Responsive and motion rules

Start with single-column composition. At 48rem, introduce horizontal header and
workflow arrangements; at 64rem, increase card padding. Below 48rem, native
Streamlit columns stack and result tabs scroll horizontally. Sidebar visibility
uses Streamlit's automatic viewport behaviour. Breakpoint literals in media
queries mirror the token reference values (CSS custom properties cannot be
used in media query conditions).

Hover-capable devices lift buttons by 1px. Pressed buttons inset by 1px. Standard
transitions use 300ms ease-out; reduced-motion removes transforms/transitions.

## Behaviour and validation

The workflow identifies the current step in text and with `aria-current`.
After a completed run, download becomes the primary action. A download request
advances the guide to Export; it does not claim that the browser saved the file.
Changing inputs clears prior results. Native upload, search, tabs, status and
Excel download controls remain responsible for keyboard interaction.

Run `python -m unittest discover -s tests -v` to check application/import flows.
The September 2026 redesign passed these tests, plus a colour-contrast check.
A connected browser was unavailable during implementation: visual QA at 375px,
768px and desktop widths, keyboard traversal, zoom, and screen-reader checks
remain required before asserting complete WCAG conformance. Streamlit upgrades
may change the native `data-testid` selectors used by the stylesheet.
