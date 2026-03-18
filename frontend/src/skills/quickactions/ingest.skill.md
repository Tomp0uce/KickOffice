---
name: Nettoyer les données
description: "Nettoie et normalise les données de la plage sélectionnée : suppression des doublons, correction des formats incohérents, harmonisation de la casse et des types de données."
host: excel
executionMode: agent
icon: Database
actionKey: ingest
---

# Ingest Quick Action Skill (Excel)

## Purpose

Automatically clean, validate, and structure raw pasted data into a formatted Excel table with proper data types and formatting.

## When to Use

- User clicks "Smart Ingestion" Quick Action in Excel
- Raw data has been pasted (from CSV, web, text files, or other sources)
- Goal: Convert messy raw data into clean, structured Excel table

## Input Contract

- **Selected cells**: Raw pasted data (may have formatting issues, wrong delimiters, inconsistent data types)
- **Context**: Excel worksheet with unformatted data
- **Common issues**: Wrong decimal separators (dots vs commas), inconsistent date formats, text that should be numbers, mixed delimiters

## Output Requirements

1. **Silently fix locale issues**: Correct decimal separators, date formats without asking
2. **Use `setCellRange` for corrections**: Modify cells directly to fix data type issues
3. **Create Excel table**: MUST call `createTable` with hasHeaders=true to convert range to formatted table
4. **Auto-detect headers**: Identify first row as headers if applicable
5. **Apply data types**: Ensure numbers are numbers, dates are dates, text is text
6. **No user confirmation**: Execute cleaning + table creation immediately

## Tool Execution Order

**CRITICAL SEQUENCE**:

1. **Analyze** — Use `getSelectedCells` or `getWorksheetData` to inspect raw data
2. **Clean** — Use `setCellRange` to fix locale/formatting issues (decimal separators, dates)
3. **Convert to Table** — Use `createTable` with `hasHeaders: true` to finalize

## Common Data Cleaning Patterns

### 1. Decimal Separator Issues

European formats (comma as decimal): `1,234.56` (US) vs `1.234,56` (EU)

**Fix**: Detect pattern and use `setCellRange` to update values:

```javascript
// If data shows "12,5" but should be 12.5
setCellRange({ address: "A2:A10", values: [[12.5], [15.3], ...] })
```

### 2. Date Format Inconsistencies

Common issues: "03/14/2026" vs "14/03/2026" vs "2026-03-14"

**Fix**: Parse and normalize to Excel-compatible date serial numbers or ISO format

### 3. Text Numbers

Numbers stored as text (often from CSV imports): `"1234"` instead of `1234`

**Fix**: Convert strings to numeric values via `setCellRange`

### 4. Mixed Delimiters

Data split incorrectly (tab vs comma vs semicolon)

**Fix**: If data is in single column but should be multiple, may need to split (or instruct user to re-paste with correct delimiter)

## Tool Usage

### Required Tools

- **`getSelectedCells`**: Inspect current selection and data structure
- **`setCellRange`**: Modify cells to fix locale/data type issues (multiple calls OK)
- **`createTable`**: MANDATORY final step — converts range to Excel table

### Tool Parameters

```typescript
// Clean data
setCellRange({
  address: 'A2:C10', // Target range
  values: [
    // 2D array with corrected values
    [12.5, '2026-03-14', 'Category A'],
    [15.3, '2026-03-15', 'Category B'],
    // ...
  ],
});

// Create table
createTable({
  address: 'A1:C10', // Full range including headers
  hasHeaders: true, // ALWAYS true if first row is headers
});
```

## Example Execution

### Scenario: CSV data pasted with European decimal format

**Initial State** (cells A1:C5):

```
| Product  | Price  | Quantity |
|----------|--------|----------|
| Widget A | 12,50  | 100      |
| Widget B | 15,75  | 250      |
| Widget C | 8,25   | 150      |
```

**Step 1**: Analyze

```javascript
getSelectedCells(); // Returns A1:C5, shows comma decimals
```

**Step 2**: Fix decimals

```javascript
setCellRange({
  address: 'B2:B4',
  values: [[12.5], [15.75], [8.25]], // Converted to proper decimals
});
```

**Step 3**: Create table

```javascript
createTable({
  address: 'A1:C4',
  hasHeaders: true,
});
```

**Final State**: Properly formatted Excel table with:

- Headers in bold
- Filter buttons on header row
- Alternating row shading
- Correct numeric data types

## Edge Cases

### No Headers Detected

If first row looks like data (not headers):

- Set `hasHeaders: false` in `createTable`
- Excel will auto-generate headers (Column1, Column2, etc.)

### Very Large Dataset (>10,000 rows)

- Process in chunks if needed
- Prioritize critical columns (dates, currencies)
- Still call `createTable` on full range

### Already a Table

If `getSelectedCells` indicates data is already in table format:

- Skip `createTable` call
- Only apply cleaning via `setCellRange` if issues detected

### Mixed Data Types in Column

Example: Column has both numbers and text ("N/A", "TBD")

- Keep as text if >30% are text values
- Convert to numbers if <30% are text (replace text with null or 0)

## Locale Detection Heuristics

**Decimal separator**:

- If numbers have `.` for decimals → US/UK format (1234.56)
- If numbers have `,` for decimals → European format (1234,56)
- Detect by scanning first 10 numeric cells

**Date format**:

- MM/DD/YYYY → US format
- DD/MM/YYYY → European format
- YYYY-MM-DD → ISO format (keep as-is)
- Detect by checking if day values exceed 12

## Quality Check

After ingestion, verify:

- ✓ All decimal separators corrected?
- ✓ Dates converted to proper Excel date format?
- ✓ Numeric columns are actual numbers (not text)?
- ✓ Excel table created with filters enabled?
- ✓ Headers properly identified?

## Ingest vs Other Excel Actions

- **Ingest** = clean + structure raw data → Excel table (one-time transformation)
- **Auto-Graph** = analyze structured data → create chart
- **Explain** = describe what a formula/data means
- **Formula Generator** = help user build custom formulas
