# Visual Style

Use these standards for professional spreadsheet presentation and data-entry usability.

## Business styling

- Use a restrained palette: one primary color, one accent color, and neutral fills.
- Prefer a readable business font such as Aptos, Calibri, Arial, or the organization’s standard font.
- Use the selected workbook font consistently in every visible title, instruction, input, formula, note, and label.
- Use size, weight, spacing, and limited fill color to establish hierarchy.
- Keep decorative elements subordinate to the data.
- Apply consistent styling across sheets with the same role.
- Keep implementation details such as programming-language, library, API, or internal tool names out of user-facing workbook instructions.

## Header hierarchy

- Use a clear workbook or report title where needed.
- Distinguish section headings from table headers.
- Use bold, high-contrast table headers.
- Freeze header rows and identifier columns when users will scroll through long tables.
- Enable filters on structured data ranges.
- Avoid repeated title blocks that consume working space.

## Layout and readability

- Set column widths to show common values without excessive whitespace.
- Wrap text for descriptions, notes, and long headers where it improves readability.
- Set row heights to display wrapped content without clipping.
- Instruction and help text must always be fully visible; widen its text area or increase row height rather than relying on overflow.
- Align text left, numeric values right, and headers consistently.
- Use whitespace and subtle fills instead of heavy borders to separate sections.
- Keep raw data in a dedicated table or sheet; do not mix it with dashboard presentation unless the workflow requires a compact single-sheet design.
- Avoid excessive merged cells. Never merge cells inside sortable or filterable data ranges.
- Do not style blank spacer cells merely to extend a header band or decoration.
- Freeze panes only when users will materially benefit from persistent headers or identifiers.

## Number formats

Apply formats consistently by meaning:

- Dates: use an unambiguous format appropriate to the audience, such as `dd-mmm-yyyy`.
- Months: use `mmm yyyy` while storing a real date.
- Currency: include the correct symbol or currency code and consistent decimals.
- Accounting: align currency symbols and display negative values consistently.
- Percentages: store decimal values and apply a percentage format; do not type `%` into text.
- Counts and whole numbers: avoid unnecessary decimals.
- Measurements: state the unit in the header when it is not obvious.
- Zeros: choose one consistent display (`0`, `-`, or blank) based on the reporting need.

Do not use formatting to disguise incorrect data types.

## Inputs, formulas, and controls

- Give user-input cells a subtle, consistent fill or border treatment.
- Keep formula and output cells visually neutral or use a restrained protected style.
- Reserve green, red, and amber primarily for semantic status, exceptions, thresholds, and conditional formatting—not generic formula identification.
- Include a small legend only when the distinction is not self-evident.
- Prefer checkboxes for true/false states such as `Paid`, `Approved`, or `Active`.
- Prefer dropdowns for multi-state fields such as status, category, account, or frequency.
- Keep dropdown lists concise, mutually understandable, and sourced from a Setup area where appropriate.
- Do not use free text when a controlled field is required for reliable filtering or formulas.

## Conditional formatting

Use conditional formatting only when it communicates an actionable state or material exception:

- thresholds and variances;
- overdue or incomplete records;
- duplicate or invalid values;
- budget or performance status;
- reconciliation failures.

Keep the rule set small, non-conflicting, and easy to explain. Use accessible colors and pair color with text, symbols, or values when users must distinguish critical states.

Do not apply decorative heatmaps or traffic-light colors without a defined business meaning.

## Dashboards

- Give each dashboard a clear purpose and reporting period.
- Place the most important KPIs first.
- Limit charts to those that answer a specific question.
- Use clear titles, direct labels where practical, and consistent scales.
- Charts must remain visually useful when the dataset is empty or sparse. Use a restrained user-facing empty-state message when it improves comprehension without fragile formulas.
- Prevent chart titles, axis titles, legends, labels, and plot areas from overlapping.
- For period-based charts, explicitly verify that rendered month or period labels are visible and readable with representative data; visible bars without category labels are not sufficient.
- Remove axis titles, legends, gridlines, or labels that do not add information.
- Use subtle gridlines and readable chart typography.
- Avoid 3D charts, unnecessary legends, dense gridlines, and excessive color.
- Keep filters or selectors together and visually distinct from outputs.
- Ensure chart source ranges expand correctly with new data.
- Keep operational raw data out of the dashboard unless drill-down is explicitly required.
- Check readability at normal zoom and on the expected screen or print size.

## Final visual review

- Give report and summary sheets a clear visible title and hierarchy.
- Confirm frozen panes and filters work as intended.
- Check for clipped headers, `#####` values, hidden content, and inconsistent alignment.
- Verify column widths, wrapped text, row heights, and page layout.
- Confirm input and formula cells remain distinguishable.
- Confirm conditional formats and charts reference the intended ranges.
- Review every visible sheet at normal zoom before delivery.
