# dev-tests — payroll reconciliation test harnesses

Node scripts that replay the real inline JS from `payroll-reconciliation.html` /
`payroll-history.html` outside the browser (DOM, jsPDF and XLSX are stubbed) so
report and LSL-engine changes can be verified against real saved data without
opening the app. Not referenced by any page — safe to ignore in deployment.

Run from this folder with plain Node (no dependencies):

- `node test-audit.js` — history-page PDF export fed Zummo's April 2026 snapshot
  from the OneDrive history JSON; dumps the LSL pages' text.
- `node test-recon-audit.js` — recon-page "load past period" restore, then PDF export.
- `node test-recon-excel.js` — same restore, then Excel export; dumps LSL sheets.
- `node test-live-classification.js` — runs `calculateLSLProvision()` live with
  fabricated remuneration/contacts inputs and asserts the AASB 101 classification
  (vested or vesting ≤ 12 months → current), period-end reference date, and
  state-override/defaulted-state behaviour. Exits non-zero on failure.

Notes for editing these: the app's globals are `let`-bound inside the eval'd
script scope, so any test setup that assigns them must be appended into the
string passed to `eval`, not run afterwards. Snapshots are read from
`C:\Users\f869f\OneDrive - LF Advisory Pty Ltd\LF Advisory Workpapers\lf-advisory-payroll-history.json`.
