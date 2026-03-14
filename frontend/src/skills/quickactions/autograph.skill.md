# Auto-Graph Quick Action Skill (Excel)

## Purpose
Analyze selected data, generate derived columns if needed, and automatically insert the most appropriate chart into the Excel workbook.

## When to Use
- User clicks "Auto-Graph" Quick Action in Excel
- Selected data represents a dataset suitable for visualization
- Goal: Create insightful chart directly in the workbook without manual chart configuration

## Input Contract
- **Selected cells**: Data range (may be table, range, or mixed data)
- **Context**: Excel worksheet with data to visualize
- **Expectations**: Chart inserted into workbook, not just recommendations

## Output Requirements
1. **Analyze data structure**: Identify x-axis, y-axis, categories, time series, etc.
2. **Generate derived columns if beneficial**: Create calculated fields and highlight them with `formatRange`
3. **Select appropriate chart type**: Column, line, scatter, pie based on data characteristics
4. **Insert chart using `manageObject`**: MUST call tool to actually create the chart in Excel
5. **Set `hasHeaders: true`**: If first row/column contains labels
6. **Infer source address**: Use `getSelectedCells` if needed to determine `source` parameter

## Tool Execution Order

**CRITICAL SEQUENCE**:
1. **Inspect Data** — `getSelectedCells` or `getWorksheetData` to understand structure
2. **Generate Derived Columns (Optional)** — `setCellRange` to add calculated fields
3. **Highlight New Columns** — `formatRange` with distinctive formatting (e.g., yellow fill)
4. **Insert Chart** — `manageObject` with action='create', type='chart', config with proper source address

## Chart Type Selection Guide

### Column Chart (Vertical Bars)
**Use when**:
- Comparing categories (e.g., sales by region, product performance)
- Discrete categories on x-axis
- 2-10 categories ideal

**Data structure**:
```
| Region    | Sales |
|-----------|-------|
| North     | 1200  |
| South     | 900   |
| East      | 1500  |
| West      | 1100  |
```

### Line Chart
**Use when**:
- Time series data (dates on x-axis)
- Showing trends over time
- Continuous data

**Data structure**:
```
| Month | Revenue |
|-------|---------|
| Jan   | 50000   |
| Feb   | 52000   |
| Mar   | 48000   |
```

### Scatter Plot
**Use when**:
- Showing relationship between two numeric variables
- Looking for correlations
- No clear x/y hierarchy

**Data structure**:
```
| Height | Weight |
|--------|--------|
| 170    | 68     |
| 165    | 62     |
| 180    | 75     |
```

### Pie Chart
**Use when**:
- Showing parts of a whole (percentages)
- 3-7 categories maximum
- All values positive

**Data structure**:
```
| Category | Percentage |
|----------|------------|
| A        | 35         |
| B        | 25         |
| C        | 40         |
```

### Bar Chart (Horizontal Bars)
**Use when**:
- Many categories (>10)
- Long category names
- Better for readability than column chart

## Derived Column Examples

### Example 1: Add Growth Rate
If data has sales over time:
```
Original:
| Month | Sales |
| Jan   | 10000 |
| Feb   | 12000 |

Add column:
| Month | Sales | Growth % |
| Jan   | 10000 | -        |
| Feb   | 12000 | 20%      |
```

Use `formatRange` to highlight "Growth %" column (yellow fill)

### Example 2: Add Running Total
```
Original:
| Day   | Orders |
| Mon   | 50     |
| Tue   | 65     |

Add column:
| Day   | Orders | Cumulative |
| Mon   | 50     | 50         |
| Tue   | 65     | 115        |
```

### Example 3: Add Percentage of Total
```
Original:
| Product | Revenue |
| A       | 50000   |
| B       | 30000   |

Add column:
| Product | Revenue | % of Total |
| A       | 50000   | 62.5%      |
| B       | 30000   | 37.5%      |
```

## Tool Usage

### Required Tools
- **`getSelectedCells`**: Determine source range for chart
- **`setCellRange`** (optional): Add derived columns
- **`formatRange`** (optional): Highlight new columns
- **`manageObject`**: INSERT the chart (MANDATORY)

### manageObject Parameters
```typescript
manageObject({
  action: 'create',
  type: 'chart',
  config: {
    chartType: 'columnClustered',  // or 'line', 'scatter', 'pie', 'barClustered'
    source: 'A1:B10',              // Data range INCLUDING HEADERS
    hasHeaders: true,              // TRUE if first row is headers
    title: 'Sales by Region'       // Chart title (optional)
  }
})
```

**Chart Type Values**:
- `columnClustered` — vertical bars
- `line` — line chart
- `scatter` — scatter plot (XY)
- `pie` — pie chart
- `barClustered` — horizontal bars

## Example Executions

### Example 1: Simple Column Chart
**Data** (A1:B5):
```
| Region | Sales |
|--------|-------|
| North  | 1200  |
| South  | 900   |
| East   | 1500  |
| West   | 1100  |
```

**Execution**:
```javascript
// Step 1: Inspect
getSelectedCells()  // Returns A1:B5

// Step 2: Create chart
manageObject({
  action: 'create',
  type: 'chart',
  config: {
    chartType: 'columnClustered',
    source: 'A1:B5',
    hasHeaders: true,
    title: 'Sales by Region'
  }
})
```

### Example 2: Line Chart with Derived Growth Column
**Original Data** (A1:B6):
```
| Month | Revenue |
|-------|---------|
| Jan   | 50000   |
| Feb   | 52000   |
| Mar   | 48000   |
| Apr   | 55000   |
| May   | 57000   |
```

**Execution**:
```javascript
// Step 1: Add growth % column
setCellRange({
  address: "C1:C6",
  values: [
    ["Growth %"],
    [null],           // Jan has no previous month
    [4.0],            // Feb: (52000-50000)/50000 = 4%
    [-7.7],           // Mar: (48000-52000)/52000 = -7.7%
    [14.6],           // Apr
    [3.6]             // May
  ]
})

// Step 2: Highlight new column
formatRange({
  address: "C1:C6",
  format: {
    fill: { color: "#FFFF00" },  // Yellow
    font: { bold: true }
  }
})

// Step 3: Create line chart (Revenue + Growth)
manageObject({
  action: 'create',
  type: 'chart',
  config: {
    chartType: 'line',
    source: 'A1:C6',
    hasHeaders: true,
    title: 'Monthly Revenue and Growth'
  }
})
```

### Example 3: Scatter Plot
**Data** (A1:C10 — Height, Weight, Age):
```
| Height | Weight | Age |
|--------|--------|-----|
| 170    | 68     | 25  |
| 165    | 62     | 30  |
...
```

**Execution**:
```javascript
manageObject({
  action: 'create',
  type: 'chart',
  config: {
    chartType: 'scatter',
    source: 'A1:B10',  // Height vs Weight
    hasHeaders: true,
    title: 'Height vs Weight Correlation'
  }
})
```

## Edge Cases

### Data has no headers
Set `hasHeaders: false`:
```javascript
manageObject({
  action: 'create',
  type: 'chart',
  config: {
    chartType: 'columnClustered',
    source: 'A1:B5',
    hasHeaders: false  // Excel will auto-label as Series 1, Series 2, etc.
  }
})
```

### Multiple Y-values per X-value
Excel will automatically create grouped/stacked chart:
```
| Month | Product A | Product B |
|-------|-----------|-----------|
| Jan   | 100       | 150       |
| Feb   | 120       | 140       |
```
→ Creates multi-series column chart

### User selected non-contiguous range
Use the first contiguous block or ask via error message

### Data is already in a table
`source` should reference table name or explicit range (e.g., "Table1" or "A1:B10")

## Derived Column Best Practices

### When to generate derived columns
- Time series → Add growth rates, moving averages
- Sales data → Add % of total, rank
- Financial data → Add ratios (profit margin, ROI)

### How to highlight
```javascript
formatRange({
  address: "C1:C10",
  format: {
    fill: { color: "#FFFF00" },      // Yellow background
    font: { bold: true, color: "#000000" }
  }
})
```

## Quality Check
After chart creation, verify:
- ✓ Chart appears in Excel workbook?
- ✓ Correct chart type for data structure?
- ✓ Headers properly used as axis labels?
- ✓ Derived columns highlighted (if added)?
- ✓ Chart title descriptive?

## Auto-Graph vs Other Excel Actions
- **Auto-Graph** = analyze data → generate columns → create chart (visualization)
- **Ingest** = clean raw data → create table (data structuring)
- **Explain** = describe formula/data (education)
- **Formula Generator** = help build custom formulas (calculation)
