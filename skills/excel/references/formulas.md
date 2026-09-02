# Formula Engineering

Use these rules when designing, writing, extending, or validating spreadsheet formulas.

## Calculated values

- Use formulas for values derived from workbook data or assumptions.
- Do not hardcode subtotals, balances, variances, rates, or other calculated outputs.
- Keep assumptions in labeled input cells and reference them from formulas.
- Document sign conventions and units when they affect interpretation.
- Prefer formulas that are easy to audit over compact but opaque expressions.

## References

- Use relative references when a reference should move as the formula is filled.
- Use absolute references when a shared assumption, lookup range, threshold, or anchor must remain fixed.
- Use mixed references intentionally for row-by-column calculations.
- Prefer structured references in Excel Tables when they improve readability and automatic extension.
- Avoid fixed last-row references when the data is expected to grow.
- Do not reference entire columns in large or calculation-heavy workbooks unless the performance impact is acceptable.

## Function selection

Prefer standard, auditable functions:

- `SUMIFS` and `COUNTIFS` for multi-condition aggregation.
- `SUMIF` and `COUNTIF` for single-condition aggregation.
- `IF` for explicit business rules.
- `IFERROR` only when the fallback is meaningful; do not hide unexpected failures.
- `XLOOKUP` for clear exact-match lookups where supported.
- `INDEX` with `MATCH` when compatibility or two-dimensional lookup behavior requires it.
- `SUM`, `ROUND`, `MIN`, `MAX`, date functions, and other standard functions where they express the rule directly.

Prefer exact matches unless approximate matching is an explicit requirement. Keep lookup keys clean and type-consistent.

## Error handling

- Handle division by zero explicitly, for example with an `IF` denominator check or a meaningful `IFERROR` fallback.
- Decide whether an unavailable result should be blank, zero, `N/A`, or a labeled status; do not choose silently.
- Do not suppress all errors when an error indicates corrupt data or a broken reference.
- Validate empty, zero, negative, duplicate, and missing-lookup cases relevant to the model.

## Volatile formulas

Avoid unnecessary volatile or indirect formulas such as `OFFSET`, `INDIRECT`, and excessive use of frequently recalculated date/time functions.

Use them only when:

- the behavior cannot be expressed reliably with a non-volatile alternative;
- the expected workbook size makes the calculation cost acceptable;
- the portability impact is understood.

## Accounting and rounding

- Define whether source amounts, line calculations, subtotals, or final outputs are rounded.
- Use `ROUND` explicitly where accounting policy requires fixed precision.
- Do not rely only on displayed decimals; formatting does not change the stored value.
- Keep currency precision and rate precision separate.
- Use a consistent sign convention for income, expenses, debits, credits, assets, and liabilities.
- Reconcile rounded line totals to reported totals and document any required rounding adjustment.

## Formula extension

- Ensure formulas continue into newly added rows.
- In Excel, prefer calculated columns in Tables when appropriate so formula logic expands with the operational dataset.
- Make summaries reference scalable Tables or ranges rather than arbitrary last rows such as 501.
- Do not prepopulate hundreds of unused formulas solely to imitate expansion.
- In Google Sheets, use carefully bounded fill patterns or array formulas where they improve reliability.
- When inserting rows or columns, verify references, named ranges, chart ranges, conditional formatting, and validations.
- Test by appending at least one representative row and confirming formulas and summaries include it.

## Formula protection and business rules

- Protect formula cells from accidental edits when practical and leave intended input cells unlocked.
- Use data validation to enforce input rules that formulas depend on, such as positive-only amounts, valid dates, and controlled states.
- Do not rely on cell color or written instructions as the only safeguard for calculation integrity.

## Excel and Google Sheets compatibility

- Prefer formulas available on both platforms when cross-platform delivery matters.
- Identify Excel-only or Google-Sheets-only functions before using them.
- Do not assume the same array behavior, separator syntax, lookup support, or error behavior across platforms.
- Call out compatibility limitations in the delivery summary.
- If equivalent formulas differ, maintain a documented platform-specific version rather than pretending they are interchangeable.

Examples of features requiring compatibility review include dynamic arrays, `FILTER`, `QUERY`, `ARRAYFORMULA`, `LET`, `LAMBDA`, and platform-specific date or import functions.

## Formula validation

Before delivery:

1. Inspect representative formulas in every calculated region.
2. Confirm absolute, relative, mixed, structured, and cross-sheet references behave as intended.
3. Test normal inputs and relevant edge cases.
4. Search for `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, and unexpected `#N/A`.
5. Reconcile detail rows to subtotals, totals, and dashboard values.
6. Append a sample row and verify formula extension and range expansion.
7. Confirm formulas remain formulas after save and reopen or after write and re-read.

Do not report a spreadsheet as validated merely because the file opens.
