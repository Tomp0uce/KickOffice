# Formula Generator Quick Action Skill (Excel)

## Purpose
Help users build custom Excel formulas by understanding their intent and providing working formula solutions with explanations.

## When to Use
- User clicks "Formula Generator" Quick Action in Excel (draft mode)
- User has described what they want to calculate or accomplish
- Goal: Provide a ready-to-use formula based on user's requirements

## Input Contract
- **User request**: Natural language description of desired calculation (from draft text field)
- **Context**: Excel worksheet (user may reference columns, ranges, or data structure)
- **Language**: Respond in UI language
- **Mode**: Draft mode — formula appears in chat as a suggestion, not auto-inserted

## Output Requirements
1. **Working formula**: Complete, syntactically correct Excel formula
2. **Cell reference format**: Use appropriate format (A1, R1C1, named ranges)
3. **Clear explanation**: What the formula does and how to use it
4. **Where to put it**: Suggest which cell(s) to place the formula
5. **Formula separator**: Use correct separator for user's Excel language (`,` vs `;`)
6. **Example**: Show with sample data when helpful
7. **Return text only**: No tool calls (this is a consultation, not execution)

## Formula Language Rules

**CRITICAL**: Excel formulas use different argument separators depending on locale:

### Comma (`,`) Languages
- **English**: `=SUM(A1,B1)`
- **Chinese**: `=SUM(A1,B1)`
- **Japanese**: `=SUM(A1,B1)`
- **Korean**: `=SUM(A1,B1)`
- **Arabic**: `=SUM(A1,B1)`

### Semicolon (`;`) Languages
- **French**: `=SOMME(A1;B1)`
- **German**: `=SUMME(A1;B1)`
- **Spanish**: `=SUMA(A1;B1)`
- **Italian**: `=SOMMA(A1;B1)`
- **Portuguese**: `=SOMA(A1;B1)`
- **Dutch**: `=SOM(A1;B1)`
- **Polish**: `=SUMA(A1;B1)`
- **Russian**: `=СУММ(A1;B1)`

**Detection**: Infer from UI language or ask user if ambiguous. Default to comma for English.

## Common Formula Patterns

### 1. Conditional Logic
**User request**: "Show 'Yes' if sales > 1000, otherwise 'No'"

**Formula**:
```excel
=IF(A2>1000, "Yes", "No")
```

**Explanation**: Place this in column B (assuming sales are in column A). Drag down to apply to all rows.

### 2. Lookup Values
**User request**: "Find the price for a product ID"

**Formula**:
```excel
=VLOOKUP(A2, ProductTable, 3, FALSE)
```

**Explanation**: This looks up the ID in A2, searches the ProductTable, and returns the value from the 3rd column. Make sure ProductTable is defined as a named range or use explicit range like `Sheet1!A:D`.

### 3. Sum with Conditions
**User request**: "Total sales for 'Electronics' category"

**Formula**:
```excel
=SUMIF(Category, "Electronics", Sales)
```

**Explanation**: Replace `Category` and `Sales` with your actual column ranges (e.g., `B:B` and `C:C`). Place this formula in a summary cell.

### 4. Count Unique Values
**User request**: "How many unique customers?"

**Formula** (Excel 365):
```excel
=COUNTA(UNIQUE(A2:A100))
```

**Formula** (older Excel):
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))
```

**Explanation**: Use the first formula if you have Excel 365 with dynamic arrays. Otherwise, use the second formula.

### 5. Concatenate with Separator
**User request**: "Combine first and last name with a space"

**Formula** (Excel 365 / 2019+):
```excel
=TEXTJOIN(" ", TRUE, A2, B2)
```

**Formula** (older Excel):
```excel
=A2&" "&B2
```

**Explanation**: Place in column C. The TEXTJOIN method is more flexible if you need to combine many cells or skip blanks.

## Response Structure

### Format:
```
# Formula Solution

**Formula**:
```[language]
=[YOUR FORMULA HERE]
```

**Where to put it**: [Suggest cell location]

**How it works**: [Brief explanation]

**Example** (optional):
[Show with sample data if helpful]

**⚠️ Notes** (if applicable):
[Edge cases, requirements, or tips]
```

## Example Responses

### Example 1: Calculate Percentage
**User request**: "Help me build a formula to calculate what percentage each product is of total sales"

**Response**:
```
# Formula Solution

**Formula**:
```excel
=B2/SUM($B$2:$B$10)*100
```

**Where to put it**: Column C (assuming sales are in column B)

**How it works**:
- Divides each product's sales (B2) by the total of all sales (SUM($B$2:$B$10))
- Multiplies by 100 to get percentage
- The $ signs lock the total range so it doesn't change when you copy the formula down

**Example**:
If B2 = 500 and total sales = 2000, the formula returns 25 (25%)

**💡 Tip**: Format column C as Percentage (Home → Number → Percentage) to show as "25%" instead of "25"
```

### Example 2: Multi-Condition Lookup
**User request**: "I need a formula that finds the price for a specific product and size combination"

**Response**:
```
# Formula Solution

**Formula**:
```excel
=INDEX(Price, MATCH(1, (Product=A2)*(Size=B2), 0))
```

This is an array formula. In older Excel, press **Ctrl+Shift+Enter** instead of just Enter. In Excel 365, just press Enter.

**Where to put it**: Column C (assuming Product is in A2, Size is in B2)

**How it works**:
- Creates two arrays: one matching the product, one matching the size
- Multiplies them together (only rows matching BOTH conditions get value 1)
- INDEX finds the corresponding price from the Price column

**Setup**:
Replace `Product`, `Size`, and `Price` with your actual named ranges or cell references like:
```excel
=INDEX($E$2:$E$100, MATCH(1, ($C$2:$C$100=A2)*($D$2:$D$100=B2), 0))
```

**⚠️ Alternative** (simpler for Excel 365):
```excel
=FILTER(Price, (Product=A2)*(Size=B2))
```
```

### Example 3: Running Total
**User request**: "I want a running total of daily sales"

**Response**:
```
# Formula Solution

**Formula**:
```excel
=SUM($B$2:B2)
```

**Where to put it**: Cell C2, then copy down to other rows

**How it works**:
- The first $ locks the starting cell (B2) so it always starts from row 2
- The second reference (B2) changes as you copy down (B3, B4, etc.)
- Result: Row 2 sums B2:B2 (just day 1), Row 3 sums B2:B3 (days 1-2), Row 4 sums B2:B4 (days 1-3), etc.

**Example**:
| Day | Sales | Running Total |
|-----|-------|---------------|
| 1   | 100   | 100           | ← =SUM($B$2:B2)
| 2   | 150   | 250           | ← =SUM($B$2:B3)
| 3   | 120   | 370           | ← =SUM($B$2:B4)
```

### Example 4: French Formula
**User request** (in French): "Aide-moi à créer une formule pour calculer la TVA à 20%"

**Response** (in French):
```
# Solution de Formule

**Formule**:
```excel
=A2*0,2
```

**Où la placer**: Colonne B (en supposant que le prix HT est en colonne A)

**Comment ça marche**:
- Multiplie le prix HT (A2) par 0,2 pour obtenir 20% de TVA
- Notez l'utilisation de la virgule (`,`) comme séparateur décimal dans Excel français

**Exemple**:
Si A2 = 100 €, la formule retourne 20 € (la TVA)

**💡 Conseil**: Pour obtenir le prix TTC, utilisez `=A2*1,2` (ou `=A2+B2` si la TVA est calculée en colonne B)
```

## Common User Requests and Solutions

| User Request | Suggested Formula |
|--------------|-------------------|
| "Sum values if condition" | `=SUMIF(range, criteria, sum_range)` |
| "Count non-empty cells" | `=COUNTA(range)` |
| "Average excluding zero" | `=AVERAGEIF(range, "<>0")` |
| "Find max value" | `=MAX(range)` |
| "Get today's date" | `=TODAY()` |
| "Calculate days between dates" | `=B2-A2` |
| "Extract first name from full name" | `=LEFT(A2, FIND(" ", A2)-1)` |
| "Remove spaces" | `=TRIM(A2)` |
| "Rank values" | `=RANK(A2, $A$2:$A$10)` |
| "Random number between X and Y" | `=RANDBETWEEN(X, Y)` |

## Edge Cases

### User provides incomplete information
Ask clarifying questions:
- "Which columns contain your data?"
- "Do you want an exact match or approximate match?"
- "Should the formula return text or numbers?"

### User asks for something Excel can't do
Explain the limitation and suggest alternatives:
- "Excel can't directly send emails, but you can use VBA macros or Power Automate"
- "Excel doesn't have a built-in AI function, but you can use Python or Office Scripts"

### Formula requires array formulas (pre-365)
Explicitly mention Ctrl+Shift+Enter requirement

### User's Excel version unclear
Provide both modern (365) and legacy solutions

## Quality Check
Before delivering formula, verify:
- ✓ Syntax correct?
- ✓ Cell references appropriate?
- ✓ Correct separator (`,` vs `;`) for language?
- ✓ Explanation clear?
- ✓ Example helpful?

## Formula Generator vs Other Excel Actions
- **Formula Generator** = help BUILD custom formula (consultative)
- **Explain** = describe existing formula (educational)
- **Auto-Graph** = create charts (visualization)
- **Ingest** = clean data (data transformation)
