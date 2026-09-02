# Rendered Workbook Visual QA

Use this reference when an Excel workbook must be visually validated from its
actual rendered output.

## Required workflow

**XLSX**
→ structural inspection
→ structural validation
→ Microsoft Excel compatibility check
→ Excel recalculation
→ Excel PDF export
→ PDF page rendering to images
→ true image-based visual inspection
→ workbook correction
→ recalculate and rerender
→ final visual and structural validation

Do not call a workbook visually validated unless every relevant rendered page
image was actually inspected. OCR, PDF text extraction, workbook XML,
openpyxl style inspection, and a successful PDF export are not substitutes for
image-based inspection.

## Render PDF pages

Render the Excel-generated PDF with PyMuPDF:

```bash
python scripts/render_pdf_pages.py workbook-rendered.pdf \
  --output-dir output/rendered-pages
```

The command renders every page at 200 DPI by default and writes ordered files
such as `page-001.png`, `page-002.png`, and `page-003.png`. Increase `--dpi`
when a dense page needs closer inspection. Existing predictable page files are
not replaced unless `--replace-existing` is supplied.

Keep the source XLSX, recalculated XLSX, Excel PDF, and rendered page images as
separate artifacts.

## Image-based inspection checklist

Inspect every page image at full-page scale and zoom into dense or suspicious
areas. Evaluate:

- clipping and unreadable or truncated Russian text;
- inappropriate column widths and row heights;
- excessive whitespace and blank regions;
- horizontal and vertical alignment;
- font consistency, hierarchy, and restrained color use;
- clear distinction between input and calculated cells;
- table readability and usable data-entry layout;
- page scaling, page breaks, repeated headers, and print readiness;
- dashboard layout, KPI sizing, and KPI spacing;
- chart readability, labels, titles, axes, and legend placement;
- Setup sheet clarity;
- Transactions usability;
- Monthly Summary readability.

Check page edges, wrapped instructions, long headers, currency symbols, chart
labels, and the transitions between printed sections especially carefully.

For each observation, distinguish:

- **Workbook UX issue** — a problem users encounter in normal Excel use, such
  as clipped help text, weak hierarchy, or an unreadable chart.
- **Print-layout artifact** — a PDF-specific effect caused by page dimensions,
  scaling, margins, or page breaks that does not impair normal workbook use.
- **Expected empty-data state** — legitimate whitespace, zero values, or sparse
  output caused by limited test data rather than a design defect.

Do not automatically classify unused printable whitespace as a problem.
Screen-first usability takes priority unless the workbook is explicitly
print-first.

## Finding severity

Classify every concrete issue:

- **Critical** — content is missing, unreadable, misleading, corrupted, or
  prevents correct use of a key input, output, or chart.
- **Important** — the workbook remains usable, but presentation, print layout,
  comprehension, or data-entry quality is materially reduced.
- **Minor** — a small polish or consistency issue with limited usability impact.

Report the page number, sheet or section, observed problem, severity, and
recommended workbook correction. Fix the authored XLSX, not the PNG or PDF.

## Validation boundary

- A PDF-to-PNG conversion proves only that pages were rendered to images.
- Image inspection validates visible presentation, not formula correctness.
- A dashboard or chart intended for real data must be tested with representative non-empty data before the spreadsheet design is considered fully validated.
- Structural and formula validation remain mandatory after visual fixes.
- After any workbook correction, repeat Excel compatibility, recalculation,
  PDF export, page rendering, image inspection, and final structural checks.
