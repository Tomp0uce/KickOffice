# Chart Digitizer Quick Action Skill (Excel)

## Purpose

Extract numerical data from a chart image (equivalent to WebPlotDigitizer) and write the extracted data into the current Excel sheet, then recreate the chart.

## When to Use

- User clicks "Digitize Chart" Quick Action in Excel and uploads a chart image
- Goal: reverse-engineer a chart image into editable Excel data + a matching Excel chart

## Input Contract

- **Image**: A chart image uploaded by the user (available in `<uploaded_images>` context with `imageId`)
- **Language**: **ALWAYS respond in the UI language specified at the start of the user message as `[UI language: ...]`.**
- **Context**: Excel worksheet (active sheet receives the extracted data)

## Output Requirements

1. Extracted numeric data written to Excel via `setCellRange`
2. A matching chart created via `manageObject`
3. Visual verification via `screenshotRange`

---

## Step-by-Step Workflow

### STEP 1 — Analyze the chart image (Vision)

Look at the uploaded image and determine:
- **Chart type**: line, bar, column, scatter, area, pie
- **X-axis**: label text, min value, max value, unit (e.g., years 2010–2024, or categories)
- **Y-axis**: min value, max value, unit (e.g., 0–100 %, or 0–50 000 €)
- **Number of series**: count the distinct colors/markers
- **Series colors**: identify the exact hex color of each data series (needed for `extract_chart_data`)
- **Plot area boundaries**: estimate the fraction-based bounding box of the chart's plot area within the full image (xMinPx, xMaxPx, yMinPx, yMaxPx as 0–1 fractions)

Write down this analysis internally before calling any tools.

### STEP 2 — Extract data per series

For **each** data series, call `extract_chart_data` once:

```json
{
  "imageId": "<imageId from uploaded_images>",
  "chartType": "line",
  "xAxisRange": [2010, 2024],
  "yAxisRange": [0, 100],
  "targetColor": "#E74C3C",
  "plotAreaBox": {
    "xMinPx": 0.08,
    "xMaxPx": 0.95,
    "yMinPx": 0.05,
    "yMaxPx": 0.88
  },
  "numPoints": 50,
  "colorTolerance": 100
}
```

Repeat with a different `targetColor` for each additional series.

**Tips for accurate extraction:**
- If the chart has few discrete points (bar/column), use `numPoints` ≈ number of bars
- For smooth line/area charts, use `numPoints` 30–80
- `colorTolerance` 80–140 works for most clean charts; increase if few points are returned
- Pie/donut charts: skip `extract_chart_data` — read segment labels and percentages visually, then enter manually

### STEP 3 — Write extracted data to Excel

Find the first empty area on the active sheet using `getWorksheetData`, then write the data with `setCellRange`:

```json
{
  "sheetName": "<active sheet>",
  "address": "A1:C21",
  "values": [
    ["X", "Series 1", "Series 2"],
    [2010, 45.2, 23.1],
    ["..."]
  ]
}
```

- Use the **actual labels** from the chart axes if readable; otherwise use generic names (X, Series 1, etc.)
- Round extracted values to 2 decimal places
- If X-axis is categorical (not numeric), use the category labels as text

### STEP 3b — Convert extracted data to an Excel table

Immediately after writing the data with `setCellRange`, convert the range to a formatted Excel table:

```json
{
  "address": "A1:C21",
  "hasHeaders": true,
  "tableName": "tbl_digitized",
  "style": "TableStyleMedium2"
}
```

Use a descriptive name (e.g., `tbl_digitized_sales`, `tbl_digitized_revenue`). This enables filters, alternating row shading, and structured references.

### STEP 4 — Create the chart

Immediately after writing the data, create the chart:

```json
{
  "operation": "create",
  "objectType": "chart",
  "sheetName": "<active sheet>",
  "source": "A1:C21",
  "chartType": "Line",
  "title": "Digitized Chart",
  "seriesBy": "columns",
  "hasHeaders": true,
  "anchor": "E2"
}
```

Match the `chartType` to what was detected: `Line`, `ColumnClustered`, `BarClustered`, `Area`, `XYScatter`, `Pie`.

### STEP 5 — Verify

Call `screenshotRange` on the area containing the chart to visually compare it with the original. Report any significant discrepancies to the user.

---

## Edge Cases

### Categorical X-axis (bar/column charts)

The X-axis values are text labels, not numbers. Write them directly in the first column and choose `ColumnClustered` or `BarClustered`.

### Pie / Donut charts

`extract_chart_data` cannot reliably extract pie segments. Instead:
1. Read the segment labels and percentages visually from the image
2. Write them manually: `[["Category", "Value"], ["A", 35], ["B", 25], ...]`
3. Create a `Pie` chart from that data

### Multiple overlapping series (similar colors)

Increase `colorTolerance` to 150–200 for the lighter color; try different tolerance values if points are missing.

### No imageId available

If `<uploaded_images>` context shows no imageId, respond that the image needs to be re-uploaded and ask the user to try again.

---

## Quality Check

- ✓ Data written to Excel (not just described)?
- ✓ Chart created in the workbook?
- ✓ Chart type matches the original?
- ✓ Screenshotted and compared?
- ✓ Responded in UI language?
