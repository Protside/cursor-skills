---
name: excel
description: Create, edit, improve, and validate professional Excel XLSX and Google Sheets spreadsheets—formulas, formatting, checkboxes, dropdowns, conditional formatting, charts, dashboards, and accounting tables. Use when building or revising spreadsheets, .xlsx files, Google Sheets, workbook formulas, data validation, or financial/accounting tables; or when the user asks to improve or validate an existing sheet without breaking formulas, formatting, or controls.
---

# Spreadsheet Engineer

Build and revise professional spreadsheets for Excel (XLSX) and Google Sheets. Prefer safe, incremental edits over rebuilds. Validate before considering the work done.

## Workflow

1. Clarify the platform, goal, audience, inputs, outputs, and must-preserve features.
2. For existing spreadsheets, inspect before editing. Inventory sheets, ranges, formulas, formats, validations, controls, charts, and protection.
3. Plan the workbook structure and sources of truth.
4. Make the smallest reliable change. Prefer surgical edits over rebuilding sheets.
5. Apply formulas, validation controls, formatting, and visualizations as required.
6. Re-read or reopen the result and validate calculations, structure, controls, and presentation.
7. Report what changed, how to use it, what was validated, and any limitations.

## Reference routing

Read only the references relevant to the task:

- [references/workbook-design.md](references/workbook-design.md) — workbook architecture, tabular design, scalability, naming, and safe editing.
- [references/visual-style.md](references/visual-style.md) — formatting, controls, conditional formatting, and dashboard presentation.
- [references/formulas.md](references/formulas.md) — formula construction, portability, extension, rounding, and validation.
- [references/xlsx-engine.md](references/xlsx-engine.md) — safe XLSX loading, inspection, validation, preservation, and openpyxl limitations.
- [references/excel-desktop-qa.md](references/excel-desktop-qa.md) — Microsoft Excel recalculation, PDF rendering, COM safety, and true rendered QA on Windows.
- [references/visual-qa.md](references/visual-qa.md) — PDF page rendering and required image-based inspection for visual validation.
- [references/google-sheets.md](references/google-sheets.md) — required workflow for creating or editing Google Sheets.
- [references/google-sheets-api.md](references/google-sheets-api.md) — official Google Sheets API authentication, read-only inspection, dependencies, and future write architecture.

For most new workbooks, read workbook design, visual style, and formulas. For XLSX implementation, inspection, or validation, also read the XLSX engine reference. When Microsoft Excel recalculation or PDF export is needed on Windows, read the Excel desktop QA reference. When rendered visual inspection is required, also read the visual QA reference. For Google Sheets tasks, read the Google Sheets workflow reference; when official API authentication, inspection, or future integration is relevant, also read the Google Sheets API reference. For a narrow edit, read only the references governing the affected features.

## Core safety rules

- Never overwrite formulas with hardcoded values unless explicitly requested.
- Preserve unrelated formulas, formats, validations, checkboxes, conditional formatting, charts, named ranges, and protection.
- Do not clear or replace whole sheets when a targeted edit is sufficient.
- Do not use CSV or another lossy round-trip for a workbook that must retain spreadsheet features.
- Keep this skill instruction-first. Do not add scripts, APIs, integrations, or dependencies unless the user asks.

## Validation gate

Do not consider a spreadsheet finished until relevant checks pass:

- [ ] Every output matches the stated business rules on sample inputs
- [ ] No `#REF!`, `#VALUE!`, `#DIV/0!`, `#N/A` in normal paths (or errors are intentional and handled)
- [ ] Edge cases: empty inputs, zeros, negative amounts where relevant
- [ ] Dropdowns/checkboxes still work; new rows covered if applicable
- [ ] Conditional formats and charts still point at correct ranges
- [ ] Unrelated existing formulas/formats/validations unchanged
- [ ] Google Sheets writes were read back and verified through the API
- [ ] Print/layout sanity for tables meant to be shared
- [ ] When Excel desktop is available, XLSX opens in real Excel without repair, invalid-content, or compatibility warnings

Fix issues before delivery. Briefly report what was validated.

## Output to the user

When finishing a task, provide:

1. File path or Sheet link/name
2. How to use inputs vs what is calculated
3. Notable formulas or rules added
4. Validation summary (what was checked)
5. Any platform limitations or follow-ups
