---
name: Expliquer les formules
description: "Explique en langage clair les formules, la structure et la logique des cellules sélectionnées. Décrit ce que chaque formule calcule et comment les données sont organisées."
host: excel
executionMode: immediate
icon: HelpCircle
actionKey: explain
---

# Explain Formula Quick Action Skill (Excel)

## Purpose

Explain Excel formulas or data in simple, understandable terms—covering what it does, how it works, and potential edge cases or pitfalls.

## When to Use

- User clicks "Explain Formula" Quick Action in Excel
- Selected cell(s) contain a formula or complex data structure
- Goal: Educational explanation for the user

## Input Contract

- **Selected cells**: Cell(s) with formula or data to explain
- **Language**: **ALWAYS respond in the UI language specified at the start of the user message as `[UI language: ...]`.** If it says `[UI language: Français]`, respond entirely in French. This is mandatory — ignore the language of the spreadsheet content for the conversation language.
- **Context**: Excel worksheet

## Output Requirements

1. **Clear explanation**: What does this formula/data do?
2. **How it works**: Break down the logic step-by-step
3. **Edge cases**: Potential errors, limitations, or unusual behaviors
4. **Examples**: Show example inputs/outputs if helpful
5. **Conversational tone**: Friendly, educational (not overly technical)
6. **Return text explanation**: No tool calls needed beyond the initial getSelectedCells

## Tool Usage

**STEP 1 — MANDATORY before explaining:**
Call `getSelectedCells` to get the **actual formula** of the selected cell(s). Do NOT skip this step. The user message context may contain data values, but the formula (if any) is only returned by `getSelectedCells`.

```json
{}
```

- If the returned cell has a formula (starts with `=`), focus the explanation on **that formula only** — do NOT explain the surrounding table data.
- If the returned cell has no formula (plain value or the selection is a data range with no formulas), then explain the data structure/pattern.

**DO NOT** modify data or create charts. This is a read-only educational action.

## Explanation Structure

### 1. Summary (What)

One-sentence description of what the formula/data does.

Example: "This formula calculates the total sales for each product category."

### 2. Breakdown (How)

Step-by-step explanation of the logic.

Example:

```
The formula works in three steps:
1. SUMIF looks at the Category column (B:B)
2. It matches the category name in cell D2
3. It sums all corresponding values from the Sales column (C:C)
```

### 3. Example (When helpful)

Show a concrete example with sample data.

Example:

```
If your data looks like:
| Product  | Category | Sales |
| Widget A | Gadgets  | 100   |
| Widget B | Tools    | 150   |
| Widget C | Gadgets  | 200   |

And D2 contains "Gadgets", the formula returns 300 (100 + 200)
```

### 4. Edge Cases (Important)

Highlight potential issues or limitations.

Example:

```
⚠️ Edge cases to watch for:
- If the category name has extra spaces, the match will fail
- Empty cells are treated as zero
- Text values in the Sales column will cause #VALUE! error
```

## Common Formula Types

### VLOOKUP

**What**: Searches for a value in the first column of a range and returns a value from another column
**How**: `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
**Edge cases**: #N/A if not found, #REF! if col_index exceeds range, requires sorted data if range_lookup=TRUE

### SUMIF / SUMIFS

**What**: Sums values that meet specific criteria
**How**: `SUMIF(range, criteria, [sum_range])`
**Edge cases**: Text criteria need quotes, wildcards (\* and ?) supported, case-insensitive

### INDEX MATCH

**What**: More flexible alternative to VLOOKUP
**How**: `INDEX(return_range, MATCH(lookup_value, lookup_range, 0))`
**Edge cases**: #N/A if no match, can search left unlike VLOOKUP

### IF / IFS

**What**: Conditional logic
**How**: `IF(condition, value_if_true, value_if_false)`
**Edge cases**: Nested IFs can become hard to maintain, consider IFS for multiple conditions

### CONCATENATE / TEXTJOIN

**What**: Combines text from multiple cells
**How**: `TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)`
**Edge cases**: CONCATENATE doesn't skip blanks, TEXTJOIN does if ignore_empty=TRUE

## Example Explanations

### Example 1: Complex Nested Formula

**Formula**:

```
=IF(AND(B2>10, C2="Yes"), "Approved", IF(B2>5, "Review", "Rejected"))
```

**Explanation**:

```
# Formula Explanation

**What it does**: This formula categorizes items as "Approved", "Review", or "Rejected" based on two conditions.

**How it works**:
1. First, it checks if BOTH conditions are true:
   - B2 is greater than 10
   - C2 equals "Yes"
2. If both are true → returns "Approved"
3. If not, it checks a second condition: is B2 greater than 5?
   - If yes → returns "Review"
   - If no → returns "Rejected"

**Example scenarios**:
- B2 = 12, C2 = "Yes" → "Approved" ✅
- B2 = 8, C2 = "No" → "Review" (fails first condition but passes second)
- B2 = 3, C2 = "Yes" → "Rejected" (fails both conditions)

**⚠️ Edge cases**:
- If C2 contains "yes" (lowercase), it won't match — Excel is case-sensitive in direct comparisons. Consider using `UPPER(C2)="YES"` for case-insensitive matching.
- If B2 is blank or contains text, you'll get a #VALUE! error
```

### Example 2: VLOOKUP with Common Issues

**Formula**:

```
=VLOOKUP(A2, ProductDatabase!A:D, 3, FALSE)
```

**Explanation**:

```
# VLOOKUP Explanation

**What it does**: Looks up the value in cell A2 within the ProductDatabase sheet and returns the corresponding value from the 3rd column.

**How it works**:
1. Takes the value in A2 (e.g., a product ID)
2. Searches for that value in column A of the ProductDatabase sheet
3. When found, returns the value from the 3rd column (column C) of that row
4. FALSE means exact match required (no approximation)

**Example**:
If A2 contains "PROD-123" and your ProductDatabase looks like:
| ID        | Name    | Price | Stock |
| PROD-123  | Widget  | 29.99 | 50    |
| PROD-456  | Gadget  | 49.99 | 20    |

The formula returns 29.99 (the value in the 3rd column)

**⚠️ Common issues**:
- **#N/A error**: The lookup value doesn't exist in the first column. Consider using `IFERROR(VLOOKUP(...), "Not Found")` to handle this gracefully.
- **#REF! error**: Column index 3 doesn't exist in your range (e.g., if you only selected columns A:B)
- **Wrong result**: Make sure your range starts with the lookup column (column A in this case)
- **Performance**: Searching entire columns (A:D) can be slow on large datasets—consider using specific ranges like A2:D1000
```

### Example 3: Array Formula (SUMPRODUCT)

**Formula**:

```
=SUMPRODUCT((Category="Electronics")*(Status="Sold")*Price)
```

**Explanation**:

```
# SUMPRODUCT Explanation

**What it does**: Calculates the total price of all items that are BOTH in the "Electronics" category AND have "Sold" status.

**How it works**:
1. (Category="Electronics") creates an array of TRUE/FALSE values (1 for matches, 0 for non-matches)
2. (Status="Sold") creates another TRUE/FALSE array
3. Multiplying these arrays together gives 1 only when BOTH conditions are true
4. That result is multiplied by the Price array
5. SUMPRODUCT sums all the resulting values

**Example**:
| Category    | Status | Price |
| Electronics | Sold   | 100   | → 1 × 1 × 100 = 100 ✅
| Electronics | Pending| 200   | → 1 × 0 × 200 = 0
| Furniture   | Sold   | 150   | → 0 × 1 × 150 = 0
| Electronics | Sold   | 50    | → 1 × 1 × 50  = 50 ✅

Total: 150

**⚠️ Things to know**:
- Arrays (Category, Status, Price) must be the same length
- Text comparisons are case-sensitive
- This is more efficient than SUMIFS for complex multi-condition logic
- Named ranges make this more readable: `=SUMPRODUCT((CategoryRange="Electronics")*(StatusRange="Sold")*PriceRange)`
```

### Example 4: French Explanation

**Formula**:

```
=SI(A2>100; "Élevé"; "Normal")
```

**Explanation** (in French if UI language is French):

```
# Explication de la Formule

**Ce qu'elle fait**: Cette formule classe les valeurs comme "Élevé" ou "Normal" selon qu'elles dépassent 100.

**Comment ça marche**:
1. Vérifie si la valeur dans A2 est supérieure à 100
2. Si oui → retourne "Élevé"
3. Si non → retourne "Normal"

**Exemple**:
- A2 = 150 → "Élevé"
- A2 = 80 → "Normal"

**⚠️ Points d'attention**:
- Notez l'utilisation du point-virgule (;) comme séparateur dans les versions françaises d'Excel
- Si A2 est vide, la formule retournera "Normal" (car 0 n'est pas > 100)
```

## Explaining Data Patterns (Not Formulas)

If the selected cells contain data but no formulas:

**Identify pattern**:

- Time series data → "This appears to be monthly sales data showing an upward trend"
- Categories with values → "This is a breakdown of expenses by department"
- Table structure → "This is a structured dataset with products, quantities, and prices"

**Provide insights**:

- "The highest value is in row 5 (March) with 15,000"
- "There's a noticeable drop in Q2"
- "The data appears to be aggregated by week"

## Edge Cases

### No formula in selected cell

Explain the data structure or pattern instead

### Multiple formulas selected

Explain each unique formula (if 2-3 formulas) or the general pattern (if many similar formulas)

### Complex array formula

Break down into logical components, use visual aids (step 1, step 2, etc.)

### Circular reference detected

Explain what a circular reference is and how to fix it

## Quality Check

After explaining, verify:

- ✓ Explanation is clear and educational?
- ✓ Technical jargon minimized?
- ✓ Edge cases mentioned?
- ✓ Examples provided (when helpful)?
- ✓ Correct language used (UI language)?

## Explain vs Other Excel Actions

- **Explain** = educational description (read-only)
- **Formula Generator** = help user BUILD a new formula
- **Ingest** = clean and structure raw data
- **Auto-Graph** = create charts from data
