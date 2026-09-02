# Google Sheets Workflow

Use this workflow for any Google Sheets creation or editing task:

**READ / CREATE → PLAN → WRITE → READ BACK → VERIFY**

Do not write until the target spreadsheet, exact sheet names, affected ranges, and required preservation constraints are understood.

## 1. READ / CREATE

Before editing an existing Google Sheet:

- Inspect spreadsheet metadata.
- Confirm the spreadsheet title, locale, time zone, and relevant properties.
- Get exact sheet names and sheet identifiers.
- Inspect the target ranges and nearby headers or formulas.
- Inspect formulas, including array formulas and cross-sheet references.
- Inspect data validation rules and their source ranges.
- Inspect checkboxes and their checked/unchecked values.
- Inspect conditional formatting rules where available.
- Note protected ranges, named ranges, filters, frozen rows or columns, merged cells, charts, and hidden sheets when relevant.
- Identify the source of truth and downstream dependencies.

Never invent credentials, spreadsheet IDs, sheet names, range addresses, or existing content.

## 2. PLAN

Define:

- the smallest possible change that meets the request;
- exact target sheets and ranges;
- values, formulas, formatting, validation, or controls to preserve;
- whether rows or columns must be inserted or whether existing cells can be updated;
- how new rows will inherit formulas and validation;
- how the result will be verified.

Prefer targeted updates. Do not overwrite a whole sheet, broad range, or workbook when a narrow write is sufficient.

If credentials, access, or a spreadsheet ID are missing, stop and state what is required. Do not fabricate access details or work around authorization.

## 3. WRITE

- Write only the planned ranges.
- Preserve formulas and formatting outside the target.
- Do not replace formula cells with calculated values unless explicitly requested.
- Preserve data validations, checkboxes, conditional formatting, filters, and protected ranges.
- Use true Boolean checkbox values or the sheet’s configured custom checkbox values.
- Keep formulas compatible with Google Sheets and the spreadsheet locale.
- Apply formatting and validations only to the intended cells.
- Avoid clear-and-rewrite operations when updating a few cells is sufficient.
- Do not delete and recreate sheets for narrow edits.

For a new Sheet, establish the smallest viable structure first, then add formulas, controls, formatting, and outputs in deliberate stages.

## 4. READ BACK

Immediately after writing:

- Re-read every changed range.
- Read adjacent formula cells when the change may affect fill behavior.
- Confirm written values and formulas match the plan.
- Confirm formulas were not stored as plain text.
- Confirm validations and checkboxes remain attached to the intended cells.
- Re-inspect conditional formatting or other affected sheet properties where possible.

Do not rely only on a successful write response.

## 5. VERIFY

- Test the affected business rule with representative inputs.
- Check for formula errors and broken references.
- Confirm formulas and summaries include new rows.
- Confirm dropdowns reject or warn on invalid values as designed.
- Confirm checkboxes produce the expected Boolean or custom values.
- Confirm conditional formatting responds to expected states.
- Reconcile changed detail data to summaries, totals, or dashboards.
- Confirm unrelated ranges and features remain unchanged.
- Report what was changed and what was verified.

## Safety rules

- Never invent credentials or spreadsheet IDs.
- Never overwrite whole sheets unnecessarily.
- Never assume a tab name; read exact names first.
- Never assume an empty target; inspect it first.
- Never use a lossy export/import cycle to make an in-place edit.
- Never treat a write acknowledgement as final validation.
- If a feature cannot be inspected or preserved reliably, disclose the limitation before making a risky change.
- Keep read-only and read/write OAuth tokens separate, and request no Drive
  scope when the Sheets scope is sufficient.
- Treat `batchUpdate` as an atomic planned unit: correct a rejected request and
  resume the same spreadsheet rather than creating an accidental duplicate.
- In a `ru_RU` spreadsheet, validate locale-sensitive formula separators by
  reading formulas and calculated values back through the API.
- Bound native checkbox targets to a practical entry area because checkboxes
  may materialize `FALSE` values and expand the values-based used range.
