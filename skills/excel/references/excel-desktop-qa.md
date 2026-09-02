# Microsoft Excel Desktop QA

Use Microsoft Excel desktop through Windows COM automation when an XLSX requires authoritative Excel recalculation or Excel-native PDF rendering.

On this Windows machine, Microsoft Excel is the preferred recalculation and rendering engine. openpyxl remains the structural inspection and editing layer; it is not a substitute for Excel calculation or rendering.

## Required workflow

**CREATE / EDIT XLSX**
→ `inspect_xlsx.py`
→ `validate_xlsx.py`
→ `recalc_with_excel.py`
→ `render_with_excel.py`
→ inspect the rendered PDF visually
→ fix the authored workbook
→ recalculate again
→ render again
→ final validation

Do not skip directly from successful export to a visual-QA claim. A PDF must actually be inspected.

## Compatibility gate

- openpyxl inspection and validation alone are not sufficient for a delivery-quality XLSX.
- When Microsoft Excel desktop is available, every generated workbook must open successfully in a real Excel COM instance before it is compatibility-validated.
- Use normal load first; do not silently fall back to repair mode.
- Any `Workbooks.Open` failure, repair mode, repair notice, invalid-content message, or compatibility warning is a FAIL.
- Treat Microsoft Excel as the authoritative XLSX compatibility check on this Windows system.
- Stop recalculation and rendering when the compatibility gate fails; diagnose the authored workbook first.

## Recalculate with Excel

```bash
python scripts/recalc_with_excel.py workbook.xlsx \
  --output-dir output/recalculated
```

The command:

- copies the source into a temporary directory;
- opens the temporary copy in an isolated, hidden Excel instance;
- requests a full Excel calculation rebuild;
- waits for Excel to report calculation complete;
- saves a separate `.xlsx` copy;
- verifies that the source file hash did not change;
- closes the workbook and quits the isolated Excel instance.

The default output name is `<source>-recalculated.xlsx`. Use `--output-name` to choose another `.xlsx` filename. Existing output artifacts are not replaced unless `--replace-existing` is supplied. The source path is always rejected as an output target.

Never claim formulas were recalculated unless this command completed successfully and reported success.

Generated workbooks intended for this workflow should use automatic calculation mode without `forceFullCalc=True`. The COM tool itself requests a full rebuild; a forced-full-calculation-on-load flag can leave Excel in an `xlPending` state and prevent the completion gate from succeeding.

## Render with Excel

Prefer rendering the recalculated copy:

```bash
python scripts/render_with_excel.py \
  output/recalculated/workbook-recalculated.xlsx \
  --output-dir output/rendered
```

The command opens a temporary workbook copy read-only and exports the workbook through Excel’s own fixed-format PDF engine. It does not save the XLSX. The default PDF name is `<source>-rendered.pdf`.

Use the Excel-generated PDF as the primary rendered representation for visual QA. Confirm that the reported PDF exists and is non-empty, then inspect it visually.

Excel PDF export is a QA representation, not automatically the workbook's primary design target. Prefer screen-first workbook usability unless the user requested print-first output. Do not shrink wide interactive sheets to unreadable text merely to improve printable-page utilization, and do not treat expected whitespace on a small operational dataset as a defect by itself.

## Visual inspection

Review every rendered sheet and page for:

- clipped or truncated text;
- unsuitable column widths and row heights;
- broken or ineffective wrapping;
- excessive whitespace and blank regions;
- page breaks and poor print scaling;
- dashboard hierarchy and KPI spacing;
- clipped charts, titles, labels, or legends;
- misplaced titles;
- inconsistent fonts;
- unreadable dates, numbers, currency, and percentages;
- unclear conditional-format colors or states;
- awkward alignment;
- visually dominant helper or technical areas.

Record concrete findings by sheet and page. Fix the authored XLSX, not merely the PDF, then repeat recalculation and rendering.

## Safety rules

- Never overwrite the original XLSX during QA unless the user explicitly requests it.
- Preserve the authored XLSX separately from temporary recalculated and rendered artifacts.
- Recalculate a copy, not the source.
- Render from a temporary copy and close the workbook with `SaveChanges=False`.
- Always close COM workbook objects and call `Excel.Application.Quit()` in cleanup paths.
- Use an isolated `DispatchEx` Excel instance so cleanup does not target a user’s existing Excel session.
- Disable alerts, events, link updates, and macros during automated QA.
- Do not kill arbitrary `Excel.exe` processes; that could terminate the user’s work. Prevent orphans with isolated instances, deterministic close/quit logic, and explicit error reporting.
- Never substitute openpyxl style inspection for true rendering.
- Never claim visual QA if no rendered PDF or image was inspected.

## Limitations

- Excel COM automation requires Windows, Microsoft Excel desktop, and `pywin32`.
- PDF export follows Excel print areas, page setup, hidden-sheet state, and installed font availability.
- Excel-generated PDFs are authoritative for this workflow but may still differ from another machine because of fonts, printer metrics, Excel build, locale, and page settings.
- COM cleanup is best-effort. The tools create isolated Excel instances and always attempt `Close()` and `Quit()`, but an Excel crash or operating-system failure can still leave a process behind.
- A successful render proves that Excel exported a PDF; it does not prove that a person inspected its visual quality.

Final delivery should use the original authored XLSX unless the recalculated copy is intentionally selected as the deliverable.
