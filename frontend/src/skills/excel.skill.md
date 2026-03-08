# Excel Office.js Skill

## CRITICAL EXCEL-SPECIFIC RULES

### Rule 1: ALWAYS use 2D arrays for values and formulas

Excel ranges are always 2D, even for single cells.

**WRONG:**

```javascript
range.values = 'Hello' // Error: not an array
range.values = ['A', 'B', 'C'] // Error: not 2D
```

**CORRECT:**

```javascript
range.values = [['Hello']] // Single cell
range.values = [['A', 'B', 'C']] // 1 row, 3 columns
range.values = [['A1'], ['A2'], ['A3']] // 3 rows, 1 column
range.values = [
  ['A1', 'B1'],
  ['A2', 'B2'],
] // 2x2 grid
```

### Rule 2: Array dimensions MUST match range dimensions

**WRONG:**

```javascript
const range = sheet.getRange('A1:C3') // 3x3 range
range.values = [['Only one']] // 1x1 array - MISMATCH!
```

**CORRECT:**

```javascript
const range = sheet.getRange('A1:C3') // 3x3 range
range.values = [
  ['A1', 'B1', 'C1'],
  ['A2', 'B2', 'C2'],
  ['A3', 'B3', 'C3'],
] // 3x3 array - matches!
```

### Rule 3: Formula language depends on user's Excel locale

**English Excel:**

```javascript
range.formulas = [['=SUM(A1,B1)']] // Comma separator
range.formulas = [['=VLOOKUP(A1,B:C,2,FALSE)']]
```

**French Excel:**

```javascript
range.formulas = [['=SOMME(A1;B1)']] // Semicolon separator
range.formulas = [['=RECHERCHEV(A1;B:C;2;FAUX)']]
```

**IMPORTANT**: Check the `excelFormulaLanguage` setting in the agent context.

### Rule 3b: Charts — always specify seriesBy and hasHeaders

When calling `manageObject` to create a chart, **always** pass `seriesBy` and `hasHeaders` explicitly to avoid axis label being treated as a data series:

```json
{
  "operation": "create",
  "objectType": "chart",
  "source": "A1:D10",
  "chartType": "ColumnClustered",
  "title": "Sales by Region",
  "seriesBy": "columns",
  "hasHeaders": true
}
```

- `seriesBy: "columns"` — each column is a data series (most common)
- `seriesBy: "rows"` — each row is a data series
- `hasHeaders: true` — first row contains column headers (labels), not data values

**NEVER use the default** — always specify both parameters.

### Rule 4: Use getUsedRange() to find data bounds

**WRONG — May be slow or include empty cells:**

```javascript
const range = sheet.getRange('A1:ZZ10000')
```

**CORRECT — Only populated cells:**

```javascript
const usedRange = sheet.getUsedRange()
usedRange.load('values,address')
await context.sync()
```

### Rule 5: Never modify cells while iterating

**WRONG — May corrupt iteration:**

```javascript
const range = sheet.getUsedRange()
range.load('values')
await context.sync()

for (let row of range.values) {
  // Modifying during iteration is dangerous
}
```

**CORRECT — Read all, transform, write back:**

```javascript
const range = sheet.getUsedRange();
range.load('values');
await context.sync();

const newValues = range.values.map(row =>
  row.map(cell => /* transform */)
);

range.values = newValues;
await context.sync();
```

## AVAILABLE TOOLS

### For READING:

| Tool               | When to use                         |
| ------------------ | ----------------------------------- |
| `getSelectedCells` | Get values from current selection   |
| `getWorksheetData` | Read used range from active sheet   |
| `getDataFromSheet` | Read data from any sheet by name    |
| `getWorksheetInfo` | Get workbook structure, sheet names |
| `getAllObjects`    | List charts and pivot tables        |
| `getNamedRanges`   | List named ranges                   |
| `findData`         | Search for values workbook-wide     |

### For WRITING:

| Tool              | When to use                                        |
| ----------------- | -------------------------------------------------- |
| `setCellRange`    | **PREFERRED** — Write values, formulas, formatting |
| `clearRange`      | Clear contents or formatting                       |
| `modifyStructure` | Insert/delete rows, columns, freeze panes          |

### For ANALYSIS:

| Tool                         | When to use                        |
| ---------------------------- | ---------------------------------- |
| `createTable`                | Convert range to Excel table       |
| `manageObject`               | Create/update charts, pivot tables |
| `sortRange`                  | Sort data                          |
| `applyConditionalFormatting` | Add conditional format rules       |

### For CHART IMAGE EXTRACTION:

| Tool                 | When to use                                                |
| -------------------- | ---------------------------------------------------------- |
| `extract_chart_data` | Extract data points from a chart/graph IMAGE via pixel analysis |

### ESCAPE HATCH:

| Tool            | When to use                                       |
| --------------- | ------------------------------------------------- |
| `eval_officejs` | **LAST RESORT** — Sheet rename, advanced features |

## COMMON PATTERNS

### Read active sheet data

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet()
const range = sheet.getUsedRange()
range.load('values,address,rowCount,columnCount')
await context.sync()
```

### Write to specific range

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet()
const range = sheet.getRange('A1:C3')
range.values = [
  ['Header1', 'Header2', 'Header3'],
  [1, 2, 3],
  [4, 5, 6],
]
await context.sync()
```

### Add formula with fill-down

```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet()
const range = sheet.getRange('D2:D100')
range.formulas = [['=A2+B2']] // First cell only
range.autoFill('D2:D100', 'FillDown')
await context.sync()
```

### Format range

```javascript
const range = sheet.getRange('A1:C1')
range.format.font.bold = true
range.format.fill.color = '#4472C4'
range.format.font.color = 'white'
await context.sync()
```

### Create a Pivot Table (Tableau Croisé Dynamique)

To create a pivot table, use the `manageObject` tool with `objectType: 'pivotTable'`.

```json
{
  "operation": "create",
  "objectType": "pivotTable",
  "name": "MyPivotTable",
  "sourceData": "Sheet1!A1:D100",
  "targetCell": "F1"
}
```

## WORKFLOW: Reproduce a chart from an image

When the user uploads a chart/graph image and asks to reproduce it in Excel, follow this EXACT sequence. Do NOT skip steps or call extract_chart_data without first analyzing the image.

### Step 1: Analyze the image (Vision)

Use your vision capability to inspect the uploaded chart image and determine:
- **Chart type**: line, scatter, bar, area, column, pie, etc.
- **X axis range**: read the min and max labels on the horizontal axis (e.g., [0, 100], [2020, 2025])
- **Y axis range**: read the min and max labels on the vertical axis (e.g., [0, 50000])
- **Data series color**: identify the hex color of the line/bars/points (e.g., "#0070C0" for blue)
- **Number of series**: if multiple, note each color separately
- **Title and axis labels**: note any text for the chart title

### Step 2: Extract data via pixel analysis

Call `extract_chart_data` with the parameters you identified:
```json
{
  "imageId": "<from uploaded_images context>",
  "xAxisRange": [0, 100],
  "yAxisRange": [0, 50000],
  "targetColor": "#0070C0",
  "chartType": "line",
  "numPoints": 40
}
```

If the tool returns few or zero points, increase `colorTolerance` (default 80, try 120-150).
For multiple series, call `extract_chart_data` once per color.

### Step 3: Write data to Excel

Use `setCellRange` to write the extracted points:
```json
{
  "address": "A1",
  "values": [["X", "Y"], [0, 100], [2.5, 250], ...]
}
```

### Step 4: Create the chart

Use `manageObject` to create the chart matching the original type:
```json
{
  "operation": "create",
  "objectType": "chart",
  "source": "A1:B41",
  "chartType": "Line",
  "title": "Original Chart Title",
  "seriesBy": "columns",
  "hasHeaders": true
}
```

### Important notes:
- **Always analyze the image FIRST** to extract semantic info (axis ranges, colors)
- **imageId** is found in the `<uploaded_images>` context block — do NOT fabricate it
- The pixel extraction works best on clean charts with distinct colors
- For pie/donut charts, extract_chart_data is not suitable — use vision to read percentages and enter data manually
