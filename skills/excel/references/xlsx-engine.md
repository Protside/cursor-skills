# XLSX Engine

Use this reference for inspecting or validating Excel `.xlsx` workbooks with the reusable openpyxl tools in `scripts/`.

## Inspect before editing

Run:

```bash
python scripts/inspect_xlsx.py path/to/workbook.xlsx
```

Use `inspect_xlsx.py`:

- before changing an existing workbook;
- when sheet names, ranges, formulas, validations, tables, charts, or protection are unknown;
- to establish a pre-edit inventory;
- after a structural change when a human-readable feature report is useful.

The inspector opens the workbook with formulas preserved and also reads cached values to identify stored error values. It does not save or modify the workbook.

Detailed lists show up to 50 entries by default. Use `--max-items 0` to show all entries:

```bash
python scripts/inspect_xlsx.py path/to/workbook.xlsx --max-items 0
```

Never save an existing workbook merely to inspect it. Loading and saving can rewrite the OOXML package and may alter unsupported or partially supported features.

## Validate before delivery

Run:

```bash
python scripts/validate_xlsx.py path/to/workbook.xlsx
```

Use `validate_xlsx.py`:

- after creating an XLSX workbook;
- after editing an existing workbook;
- before delivery;
- when checking package integrity, sheet names, formulas, stored errors, validations, merged ranges, and basic structure.

The validator reports:

- `PASS` — no failed checks or warnings were found;
- `WARNING` — no failed checks were found, but a condition needs review;
- `FAIL` — one or more structural or stored-content checks failed.

Exit codes are `0` for PASS, `1` for WARNING, and `2` for FAIL.

Validation is not a substitute for business-rule testing or visual review. Compare edited workbooks with the pre-edit inspection when preservation matters.

## Safe loading rules

- Use `data_only=False` when formulas must remain formulas.
- Use `read_only=False` for a full feature inventory because optimized read-only mode does not expose every workbook feature.
- Do not call `save()` during inspection or validation.
- Close every loaded workbook, including cached-value views.
- Use `keep_links=True` when loading workbooks that may contain external links.
- Work on a copy when an edit may affect features that openpyxl does not preserve reliably.
- When editing an existing workbook, preserve its current behavior unless the requested change requires a behavioral change.
- Reopen and inspect the output after an authorized edit.

Loading with `data_only=True` returns cached results saved by a calculation-capable spreadsheet application where those results exist. It does not calculate formulas. Do not load with `data_only=True` and then save over a formula workbook.

## openpyxl limitations

openpyxl reads and writes OOXML but is not Excel. Depending on the workbook and operation, it may not fully preserve or evaluate:

- formula results and calculation state;
- VBA projects and macros;
- unsupported extensions, controls, and form objects;
- some drawing, chart, slicer, timeline, and pivot-table features;
- digital signatures;
- external-link behavior and data connections;
- application-specific metadata.

Private openpyxl attributes may be required to report some accessible features such as chart and conditional-format counts. Treat those reports as an inventory aid, not proof that every Excel feature was discovered.

## Formula preservation and validation

- Preserve formulas by loading with `data_only=False`.
- Inspect formulas as stored strings; never replace them with cached values.
- For operational datasets, prefer Excel Tables and structured references when openpyxl can preserve the required behavior safely.
- Do not manually inject or manipulate `calculatedColumnFormula` table metadata unless the exact structure is known to be valid in Microsoft Excel.
- Prefer normal cell formulas or supported openpyxl Table behavior over low-level calculated-column metadata.
- Do not assume normal formulas written into an openpyxl Table will propagate when Excel adds a row. Test a real Excel append. When calculated-column behavior is required, assign the formula through Excel to the Table column so Excel authors and maintains its own valid calculated-column metadata.
- Microsoft Excel compatibility takes priority over clever OOXML or Table optimizations.
- Do not use hundreds of preformatted blank rows or copied formulas as a substitute for tested Table expansion.
- Confirm that data validation and formula behavior remain correct when a new Table row is added; Table expansion does not guarantee every openpyxl-authored validation rule will expand automatically.
- Check formulas for obvious `#REF!` references.
- Check stored cells and cached formula results for known Excel error values.
- Treat formula-continuity detection as heuristic. A single workbook cannot prove that a value replaced an intended formula without a known-good baseline or explicit formula specification.
- Test formulas with representative inputs in a calculation-capable environment before relying on calculated outputs.

## Excel Table filters

- An Excel Table already owns an AutoFilter for its dataset.
- Do not add a worksheet-level AutoFilter covering the same range as a Table AutoFilter.
- Avoid overlapping, duplicate, or redundant worksheet and Table filter definitions.

## Worksheet views and panes

- When removing freeze panes or split panes, normalize the worksheet view and selection metadata.
- Do not leave selections referencing `topRight`, `bottomLeft`, or `bottomRight` when the corresponding pane no longer exists.
- Verify that pane, selection, active pane, and top-left-cell metadata remain internally consistent after save and reopen.

## Chart compatibility

- DrawingML `a:srgbClr` chart colors require six-digit RGB values such as `5B9BD5`, not eight-digit ARGB values such as `FF5B9BD5`.
- For column charts, explicitly set the category axis to the bottom and render representative data to confirm labels are visible; a left-positioned category axis can produce unlabeled bars.
- Treat an Excel Open failure caused by invalid chart XML as a compatibility FAIL and correct the chart generation source rather than relying on repair mode.

## Why recalculation is separate

openpyxl does not contain an Excel formula calculation engine. It can preserve and inspect formula text, but it cannot produce authoritative recalculated results.

Formula recalculation will be handled as a separate future layer because it requires a calculation-capable application or service and introduces platform, dependency, compatibility, and trust decisions. Until that layer is explicitly added, report only structural and stored-value validation and state that formulas were not recalculated.

For workbooks that will be recalculated through Excel COM, do not set `forceFullCalc=True` in generated calculation properties. This flag can leave Excel reporting `xlPending` indefinitely after a requested full rebuild. Use automatic calculation mode and let the Excel QA tool explicitly request the full rebuild.
