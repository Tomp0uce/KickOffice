# Data Trend Quick Action Skill (Excel)

## Purpose

Analyze trends in selected data, identify patterns, outliers, growth rates, and provide a concise summary with actionable insights.

## When to Use

- User clicks "Data Trend" Quick Action in Excel
- Selected data represents time series, sequential measurements, or comparable values
- Goal: Provide analytical insights about trends, not just description

## Input Contract

- **Selected cells**: Numeric data (ideally with labels/dates)
- **Language**: Respond in UI language
- **Context**: Excel worksheet with data to analyze
- **Mode**: Immediate execution (agent loop)

## Output Requirements

1. **Overall trend**: Is data increasing, decreasing, stable, or cyclical?
2. **Key patterns**: Seasonality, peaks, troughs, anomalies
3. **Outliers**: Identify unusual data points that deviate significantly
4. **Growth rate**: Calculate percentage changes or growth rates
5. **Actionable insights**: What should the user do with this information?
6. **Concise format**: 3-5 numbered insights, not a long essay
7. **Use numbers**: Quantify findings (e.g., "15% growth", "peak in March")

## Tool Usage

**Optional**:

- `getSelectedCells` or `getWorksheetData` — If you need to inspect the actual data structure

**DO NOT** modify data or create charts. This is a read-only analytical action. (User can use Auto-Graph for visualization)

## Analysis Framework

### 1. Identify Trend Direction

- **Upward trend**: Values generally increasing over time
- **Downward trend**: Values generally decreasing over time
- **Stable/flat**: Values remain relatively constant
- **Cyclical**: Values show repeating patterns (seasonal)
- **Volatile**: High variability, no clear pattern

### 2. Calculate Growth Metrics

- **Total change**: `(Last value - First value) / First value * 100`
- **Average growth per period**: `((Last/First)^(1/periods) - 1) * 100`
- **Peak-to-trough**: Identify highest and lowest points
- **Volatility**: Standard deviation or variance

### 3. Detect Patterns

- **Seasonality**: Regular patterns (e.g., higher sales in Q4)
- **Peaks/troughs**: When do maximums and minimums occur?
- **Acceleration/deceleration**: Is growth speeding up or slowing down?
- **Breakpoints**: Sudden shifts in trend direction

### 4. Identify Outliers

- **Statistical outliers**: Values >2 standard deviations from mean
- **Contextual outliers**: Values unusual given surrounding data

### 5. Provide Insights

- **What does this mean?** Interpret the trend
- **Why might this be happening?** Suggest possible causes
- **What should be done?** Recommend actions

## Response Structure

### Format:

```
# Data Trend Analysis

**Overall Trend**: [One sentence summary]

**Key Findings**:
1. [Specific insight with numbers]
2. [Pattern or outlier identified]
3. [Growth rate or change metric]
4. [Noteworthy observation]
5. [Actionable recommendation]

**📊 Quick Stats**:
- [Relevant metric 1]
- [Relevant metric 2]
- [Relevant metric 3]
```

## Example Analyses

### Example 1: Monthly Sales Data

**Data** (Monthly Revenue, Jan-Dec):

```
| Month | Revenue |
|-------|---------|
| Jan   | 50000   |
| Feb   | 52000   |
| Mar   | 48000   |
| Apr   | 55000   |
| May   | 57000   |
| Jun   | 54000   |
| Jul   | 61000   |
| Aug   | 63000   |
| Sep   | 59000   |
| Oct   | 68000   |
| Nov   | 72000   |
| Dec   | 75000   |
```

**Analysis**:

```
# Data Trend Analysis

**Overall Trend**: Strong upward trajectory with 50% annual growth and consistent monthly gains.

**Key Findings**:
1. **Consistent growth**: Revenue increased every quarter, from $50K (Jan) to $75K (Dec)—a 50% increase over 12 months
2. **Accelerating momentum**: Growth rate improved through the year—Q1 averaged 4% monthly growth, while Q4 averaged 7% monthly growth
3. **Only dip**: March saw a -7.7% decline (from $52K to $48K), but recovered immediately in April
4. **Strong finish**: Q4 (Oct-Dec) delivered the highest revenues, with November and December exceeding $70K for the first time
5. **Forecast**: If this trend continues, January 2026 revenue could reach $78K-$80K (4-6% month-over-month growth)

**📊 Quick Stats**:
- **Total annual growth**: 50% (Jan $50K → Dec $75K)
- **Average monthly increase**: $2,083
- **Best month**: December ($75K)
- **Worst dip**: March (-7.7%)
```

### Example 2: Website Traffic Data

**Data** (Daily visitors for 14 days):

```
| Day | Visitors |
|-----|----------|
| 1   | 1200     |
| 2   | 1300     |
| 3   | 1250     |
| 4   | 1400     |
| 5   | 1350     |
| 6   | 900      | ← Weekend
| 7   | 850      | ← Weekend
| 8   | 1450     |
| 9   | 1500     |
| 10  | 1550     |
| 11  | 1600     |
| 12  | 1620     |
| 13  | 950      | ← Weekend
| 14  | 920      | ← Weekend
```

**Analysis**:

```
# Data Trend Analysis

**Overall Trend**: Clear weekday/weekend pattern with weekday traffic steadily increasing, weekend traffic consistently low.

**Key Findings**:
1. **Strong weekday growth**: Weekday traffic increased from ~1200 (Day 1) to ~1600 (Day 12)—a 33% gain in just two weeks
2. **Weekend drop**: Traffic falls ~40% on weekends (Sat/Sun: ~900 visitors vs weekday average: ~1400)
3. **Consistent pattern**: The weekday/weekend cycle is highly predictable—every weekend shows similar dip
4. **No outliers**: All data points fit the expected pattern; no unusual spikes or drops
5. **Actionable insight**: Consider running targeted campaigns on weekends to reduce the traffic gap, or focus marketing budgets on weekdays when engagement is naturally higher

**📊 Quick Stats**:
- **Weekday average**: 1,421 visitors
- **Weekend average**: 905 visitors
- **Gap**: 57% more traffic on weekdays
- **Trend**: Weekday traffic +33% over two weeks
```

### Example 3: Product Defect Rates (Quarterly)

**Data**:

```
| Quarter | Defect Rate (%) |
|---------|-----------------|
| Q1 2024 | 5.2             |
| Q2 2024 | 4.8             |
| Q3 2024 | 4.5             |
| Q4 2024 | 4.1             |
| Q1 2025 | 3.9             |
| Q2 2025 | 3.5             |
| Q3 2025 | 3.2             |
| Q4 2025 | 2.9             |
```

**Analysis**:

```
# Data Trend Analysis

**Overall Trend**: Consistent downward trend—defect rates decreased by 44% over two years, indicating effective quality improvements.

**Key Findings**:
1. **Steady improvement**: Defect rate dropped from 5.2% (Q1 2024) to 2.9% (Q4 2025)—a 44% reduction
2. **Linear decline**: Each quarter shows improvement; no setbacks or plateaus
3. **Accelerating success**: Rate of improvement increased—2024 averaged -0.275% per quarter, 2025 averaged -0.3% per quarter
4. **Below industry benchmark**: Assuming 4% is industry standard, you've been below that since Q4 2024
5. **Next target**: Aim for sub-2.5% by Q2 2026 if current trajectory holds

**📊 Quick Stats**:
- **Total improvement**: -2.3 percentage points (-44%)
- **Average quarterly reduction**: -0.29%
- **Current rate**: 2.9% (Q4 2025)
- **Best quarter**: Q4 2025 (lowest defect rate)
```

### Example 4: French Data (Sales Performance)

**Data** (French monthly sales):

```
| Mois | Ventes |
|------|--------|
| Jan  | 15000  |
| Fév  | 16000  |
| Mar  | 14500  |
| Avr  | 17000  |
```

**Analysis** (in French):

```
# Analyse des Tendances

**Tendance Globale**: Croissance positive avec légère volatilité—les ventes ont augmenté de 13% sur 4 mois malgré un recul en mars.

**Observations Clés**:
1. **Croissance solide**: Les ventes sont passées de 15 000 € (janvier) à 17 000 € (avril), soit +13% en 4 mois
2. **Pic en avril**: Avril représente le meilleur mois avec 17 000 € (+17% vs janvier)
3. **Baisse en mars**: Mars a connu une baisse de -9% (16 000 → 14 500 €), mais récupération immédiate en avril
4. **Volatilité modérée**: Les variations mensuelles oscillent entre -9% et +17%, indiquant une certaine instabilité
5. **Action recommandée**: Analyser la cause de la baisse de mars pour éviter sa répétition—facteur saisonnier, campagne marketing, ou autre?

**📊 Statistiques Rapides**:
- **Croissance totale**: +13% (jan→avr)
- **Meilleur mois**: Avril (17 000 €)
- **Pire baisse**: Mars (-9%)
- **Moyenne mensuelle**: 15 625 €
```

## Outlier Detection

### Statistical Method

Calculate mean and standard deviation:

- **Mild outlier**: Value >1.5 SD from mean
- **Extreme outlier**: Value >2 SD from mean

### Contextual Method

Compare to surrounding values:

- If value differs by >20% from previous/next values → flag as outlier

## Actionable Insights Examples

### Growth Insights

- "Revenue is growing but decelerating—consider new strategies to maintain momentum"
- "Q4 consistently outperforms—allocate more budget to year-end campaigns"

### Stability Insights

- "Traffic is flat despite marketing spend—reevaluate campaign effectiveness"
- "Costs remain stable despite volume increase—excellent operational efficiency"

### Volatility Insights

- "High variability suggests external factors—investigate correlation with market events"
- "Unpredictable sales pattern—implement demand forecasting system"

### Outlier Insights

- "March spike (+50%) coincided with product launch—replicate strategy for future releases"
- "August dip likely due to vacation season—plan reduced staffing accordingly"

## Edge Cases

### Insufficient data (<5 data points)

"Limited data makes trend analysis unreliable. Collect more data points for meaningful insights."

### All values identical

"No trend detected—all values are constant at [X]. This could indicate data source issue or truly stable metric."

### Extreme outliers

Mention outliers but don't let them dominate the analysis—focus on overall trend

### Non-numeric data

"Selected data is not numeric. Trend analysis requires numerical values."

## Quality Check

After analysis, verify:

- ✓ Quantified findings (percentages, numbers)?
- ✓ Identified actual patterns (not just describing data)?
- ✓ Provided actionable insights?
- ✓ Concise (3-5 points, not essay)?
- ✓ Correct language?

## Data Trend vs Other Excel Actions

- **Data Trend** = analytical insights about patterns and changes (read-only)
- **Explain** = describe formula/data structure (educational)
- **Auto-Graph** = create visual charts (visualization)
- **Formula Generator** = help build formulas (calculation assistance)
