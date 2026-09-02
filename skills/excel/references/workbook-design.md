# Workbook Design

Use these standards when planning a new workbook or changing workbook structure.

## Screen-first design

- Design spreadsheet UX for normal on-screen use unless the user explicitly requests a print-first workbook.
- Treat print and PDF output as an important QA view, not as the default design target.
- Do not force a wide worksheet onto one printed page when the resulting text becomes unreadably small.
- Operational sheets may leave unused printable whitespace when their interactive layout is appropriate.
- Do not add rows, stretch tables, or distort working layouts merely to fill a printed page.

## Separate workbook roles

Use distinct areas or sheets for:

- **Inputs** — values entered, imported, or selected by users. Make required fields and allowed values explicit.
- **Calculations** — formulas, transformations, mappings, allocations, and reconciliations. Keep manual entry out of calculated ranges.
- **Outputs / Dashboard** — summaries, reports, charts, KPIs, and print-ready views. Derive these from inputs and calculations.
- **Reference / Setup** — categories, rates, thresholds, account lists, mapping tables, and configuration used by formulas or validations.

The separation may be physical sheets or clearly labeled sections in a small workbook. Do not split sheets merely to make the workbook look sophisticated.

## Source-of-truth rules

- Store each authoritative value in one place.
- Reference the source value from formulas and reports instead of copying it into multiple sheets.
- Distinguish entered data from calculated data.
- Keep lookup lists and configurable assumptions in a defined Setup area.
- Arrange related Setup and reference tables into readable logical groups. Avoid one excessively wide horizontal strip when a compact multi-row arrangement is clearer on screen.
- Design dropdown sources to expand when their Setup lists grow. Do not pair an expanding Setup table with a fixed named range that silently excludes new rows.
- Expose a configuration parameter only when changing it actually changes workbook behavior. Otherwise make the value informational or remove it.
- Do not duplicate totals as hardcoded values.
- Document units, sign conventions, date basis, currency, and important assumptions near their source.
- If imported data is authoritative, preserve it as a raw layer and derive cleaned or mapped values elsewhere.

## Tabular data

- Model one row as one entity, transaction, event, observation, or other clearly stated record.
- Give each column one stable meaning and one data type.
- Use one header row and unique, descriptive column names.
- Avoid blank separator rows, subtotals, merged cells, and decorative headings inside machine-readable data ranges.
- Use stable identifiers when records may otherwise be ambiguous.
- Store dates as real dates, numbers as numbers, and Boolean states as true/false or checkboxes.
- Do not encode multiple facts in one cell when separate fields are needed for filtering, formulas, or validation.

## Scalable structure and formulas

- Build tables and ranges so new rows inherit formulas, formats, validations, and conditional formatting.
- In Excel, prefer Tables for operational datasets when their automatic expansion, calculated columns, filters, and structured references improve reliability.
- Use a Table's built-in filter for its dataset; do not layer a duplicate worksheet AutoFilter over the same range.
- Start operational Tables with only a small practical number of entry rows. Do not pre-format hundreds of unused blank rows merely to simulate scalability.
- Formula-driven columns and dependent reports must expand with the dataset; do not stop at an arbitrary row such as 501.
- In Google Sheets, use bounded open-ended ranges or array formulas carefully; avoid full-column calculations when they create material performance cost.
- Keep summary formulas independent of fixed last-row assumptions.
- Put calculated columns beside their source data when that keeps row logic clear.
- Avoid manual subtotal rows within source data; calculate summaries outside the data table.
- Test row insertion and append behavior before delivery.

## One sheet versus multiple sheets

Use one sheet when:

- the dataset is small and has one purpose;
- users need a simple entry-and-review workflow;
- calculations and outputs are limited and can remain clearly separated;
- splitting would add navigation cost without improving control.

Use multiple sheets when:

- inputs, calculations, and outputs have different audiences or edit permissions;
- raw imports must remain untouched;
- reference lists drive validations or mappings;
- multiple related tables have different row grains;
- a dashboard or print-ready report should not be mixed with operational data;
- calculation complexity would make one sheet difficult to audit.

Never place unrelated datasets in one table. Never create separate monthly sheets when a single transaction table with a date or period field will scale better, unless the business process specifically requires period-specific sheets.

## Naming conventions

- Use short, descriptive sheet names such as `Transactions`, `Budget`, `Calculations`, `Dashboard`, and `Setup`.
- Use singular or plural terms consistently.
- Use descriptive headers such as `Transaction Date`, `Account`, and `Net Amount`; avoid unexplained abbreviations.
- Give named ranges and tables stable, meaningful names. Follow platform naming restrictions.
- Include units in headers when ambiguity exists, for example `Rate (%)` or `Amount (USD)`.
- Avoid version labels such as `Final`, `Final 2`, or dates in internal sheet names unless versioning is part of the required workflow.

## Safe editing of existing spreadsheets

Before editing:

1. Inventory exact sheet names and used ranges.
2. Inspect formulas, named ranges, data validations, checkboxes or controls, conditional formatting, charts, tables, merged cells, hidden rows or columns, and protection relevant to the change.
3. Identify the source of truth and dependencies of the target cells.
4. Decide the smallest range-level or cell-level change that satisfies the request.

While editing:

- Preserve current workbook behavior unless the requested change requires modifying it.
- Preserve formulas and workbook features outside the target.
- Do not delete and recreate a sheet to make a narrow change.
- Do not replace a formula with a value unless explicitly requested.
- Extend formulas, formats, and validations deliberately when adding rows or columns.
- Avoid moving cells referenced by formulas, charts, validations, or named ranges unless all dependencies are updated.
- Use a copy or a new staging sheet first when fidelity is uncertain.

## Data-entry controls

- Enforce practical business rules with data validation, not only with instructions or comments.
- Use validation to reject invalid signs, dates, ranges, states, and required controlled values when the rule is unambiguous.
- Keep formula cells protected from accidental edits when practical while leaving intended input cells editable.
- Excel may block Table row expansion while worksheet protection is active even when row insertion is allowed. For an expandable protected Table, define and test an explicit workflow such as temporarily unprotecting, adding rows, and re-protecting; do not assume `AllowInsertingRows` alone makes Table expansion work.
- Avoid unnecessary frozen panes, merged cells, styled spacer cells, and large preformatted empty regions.
- After removing a frozen or split pane, clear stale pane-specific selections and confirm that the worksheet view remains internally consistent.

After editing:

- Reopen or re-read the workbook.
- Compare affected formulas and features with the pre-edit inventory.
- Confirm append behavior and source-to-output reconciliation.
- Verify that unrelated workbook content remains unchanged.
