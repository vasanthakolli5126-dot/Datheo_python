# Task 3: Financial KPI Dashboard (Power BI): Build Guide

A `.pbix` can only be saved from Power BI Desktop, so this package gives you everything to assemble it in about 30–40 minutes: data prep, model, all DAX, the waterfall, scenario sliders and the page layout.

**Files:** `financial_sample.xlsx` (actuals) and `Budget_Targets.csv` (budget, see the note below).

> **Budget note:** Microsoft's Financial Sample has **no budget column**. `Budget_Targets.csv` is an *illustrative placeholder* I generated (Segment × Country × Product × Year; 5% discount rate, 80% COGS ratio, gross-sales target ±10% of actual). Replace it with your real targets if you have them. Keep the same columns and nothing else changes.

---

## 1. Load the data (Power Query)

**Get Data → Excel →** `financial_sample.xlsx` → Sheet1 → *Transform Data*.

1. Trim column names (the file has trailing spaces, e.g. `"Segment "`). Select all columns → Transform → Format → Trim, or in M:
   ```
   = Table.TransformColumnNames(Source, Text.Trim)
   ```
2. Set types: Date → Date; Units Sold, Gross Sales, Discounts, Sales, COGS, Profit → Fixed decimal; Year, Month Number → Whole number.
3. Rename `Sales` to `Net Sales` (it is net of discounts) and the table to `FactSales`.
4. `Discount Band` is null where there is no discount. Replace null with `None`.
5. Add an index column named `RowID` (optional).

**Get Data → Text/CSV →** `Budget_Targets.csv`, name it `FactBudget`, and set types.

### Dimension tables (Reference `FactSales`, keep one column, Remove Duplicates)
| Table | Source |
|---|---|
| `DimSegment` | Segment |
| `DimCountry` | Country |
| `DimProduct` | Product |
| `DimDate` | DAX calendar below |

```DAX
DimDate =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2013, 1, 1 ), DATE ( 2014, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "MMM" ),
    "Quarter", "Q" & ROUNDUP ( MONTH ( [Date] ) / 3, 0 )
)
```
Mark as date table (Table tools → Mark as date table → `Date`). Sort `Month` by `Month Number`.

## 2. Relationships (star schema)

| From (many) | To (one) | Direction |
|---|---|---|
| FactSales[Segment] | DimSegment[Segment] | Single |
| FactSales[Country] | DimCountry[Country] | Single |
| FactSales[Product] | DimProduct[Product] | Single |
| FactSales[Date] | DimDate[Date] | Single |
| FactBudget[Segment / Country / Product] | the same dims | Single |
| FactBudget[Year] | DimDate[Year] | *Do not link.* Budget is annual; use the `Year` slicer from DimDate with the budget measure below |

Because the budget is annual, add a small `DimYear` table (`DimYear = DISTINCT(DimDate[Year])`), relate `FactBudget[Year]` → `DimYear[Year]`, and use `DimYear[Year]` for the slicer. For the actuals measures use `FactSales[Year]`, or relate DimYear to FactSales[Year] too.

## 3. Income statement hierarchy

Gross Sales → (–) Discounts → **Net Sales** → (–) COGS → **Gross Profit**

In the dataset: `Gross Sales = Units Sold × Sale Price`, `Net Sales = Gross Sales – Discounts`, `Profit = Net Sales – COGS`.

## 4. Base measures

Create a blank table `_Measures` (Enter Data) and put all measures in it.

```DAX
Gross Sales   = SUM ( FactSales[Gross Sales] )
Discounts     = SUM ( FactSales[Discounts] )
Net Sales     = SUM ( FactSales[Net Sales] )
COGS          = SUM ( FactSales[COGS] )
Gross Profit  = SUM ( FactSales[Profit] )
Units Sold    = SUM ( FactSales[Units Sold] )

-- Required KPIs
Gross Profit % = DIVIDE ( [Gross Profit], [Net Sales] )
COGS Ratio     = DIVIDE ( [COGS], [Net Sales] )
Discount Rate  = DIVIDE ( [Discounts], [Gross Sales] )

Segment Margin =
VAR _segProfit = [Gross Profit]
VAR _segSales  = [Net Sales]
RETURN DIVIDE ( _segProfit, _segSales )

-- Contribution of the selected segment to total company profit
Segment Profit Share =
DIVIDE ( [Gross Profit], CALCULATE ( [Gross Profit], ALL ( DimSegment ) ) )
```
*Note:* `Segment Margin` is evaluated in the filter context of whatever segment is on the visual, so it works in a matrix or bar by Segment. `[Gross Profit %]` and `[Segment Margin]` give the same number at total level.

## 5. Budget and variance measures

```DAX
Budget Gross Sales = SUM ( FactBudget[Budget Gross Sales] )
Budget Discounts   = SUM ( FactBudget[Budget Discounts] )
Budget Net Sales   = SUM ( FactBudget[Budget Sales] )
Budget COGS        = SUM ( FactBudget[Budget COGS] )
Budget Profit      = SUM ( FactBudget[Budget Profit] )

Net Sales Var      = [Net Sales] - [Budget Net Sales]
Net Sales Var %    = DIVIDE ( [Net Sales Var], [Budget Net Sales] )
Profit Var         = [Gross Profit] - [Budget Profit]
Profit Var %       = DIVIDE ( [Profit Var], ABS ( [Budget Profit] ) )
Budget Profit %    = DIVIDE ( [Budget Profit], [Budget Net Sales] )
Margin Var (pts)   = [Gross Profit %] - [Budget Profit %]

Var Status =
SWITCH ( TRUE (),
    ISBLANK ( [Budget Net Sales] ), BLANK (),
    [Net Sales Var %] >= 0, "▲ On / above budget",
    [Net Sales Var %] >= -0.05, "● Within 5%",
    "▼ Below budget" )
```
Use `Var Status` or `Net Sales Var %` for conditional formatting (green ≥ 0, amber 0 to –5%, red < –5%).

The same measures work across **geographic segments** (DimCountry), **product lines** (DimProduct) and **segments / cost centres** (DimSegment). For *manufacturing cost centres*, use `Product` × `Manufacturing Price` (add a column `Mfg Cost Centre = FactSales[Product]`) and compare `COGS` vs `Budget COGS`:

```DAX
COGS Var   = [COGS] - [Budget COGS]            -- positive = overspend
COGS Var % = DIVIDE ( [COGS Var], [Budget COGS] )
```

## 6. Waterfall chart (revenue erosion)

Create a helper table, since the native Waterfall visual needs one category column and one value column:

```DAX
WF Steps =
DATATABLE ( "Step", STRING, "Order", INTEGER,
    { { "Gross Sales", 1 }, { "Discounts", 2 }, { "Net Sales", 3 },
      { "COGS", 4 }, { "Gross Profit", 5 } } )
```
```DAX
WF Value =
SWITCH ( SELECTEDVALUE ( 'WF Steps'[Step] ),
    "Gross Sales",  [Gross Sales],
    "Discounts",   -[Discounts],
    "Net Sales",    [Net Sales],
    "COGS",        -[COGS],
    "Gross Profit", [Gross Profit] )
```
Visual: **Waterfall chart** → Category = `WF Steps[Step]` (sort by `Order`), Y = `[WF Value]`. In Format → Sentiment, set increase = green, decrease = red, **total** = blue. Then right-click the *Net Sales* and *Gross Profit* bars → **Set as total** (or choose "Set as total" in the data point formatting) so they show the running subtotal, not a delta.

Tip: since the subtotal bars are set as totals, the chart reads correctly. If Power BI adds them to the running sum instead, make `Net Sales` and `Gross Profit` rows return `BLANK()` and let the visual compute totals.

## 7. Interactive scenario forecasting sliders

Modeling → **New parameter → Numeric range** (tick *Add slider to this page*):

| Parameter | Min | Max | Step | Default |
|---|---|---|---|---|
| `Volume Change %` | -30% | 30% | 1% | 0 |
| `Price Change %` | -15% | 15% | 0.5% | 0 |
| `Discount Rate Change (pts)` | -10% | 10% | 0.5% | 0 |
| `COGS Change %` | -20% | 20% | 1% | 0 |

Each creates a `[... Value]` measure. Use those below:

```DAX
Scn Gross Sales =
[Gross Sales] * ( 1 + [Volume Change % Value] ) * ( 1 + [Price Change % Value] )

Scn Discount Rate = [Discount Rate] + [Discount Rate Change (pts) Value]

Scn Discounts = [Scn Gross Sales] * [Scn Discount Rate]

Scn Net Sales = [Scn Gross Sales] - [Scn Discounts]

-- COGS scales with volume (not price), then with the cost lever
Scn COGS =
[COGS] * ( 1 + [Volume Change % Value] ) * ( 1 + [COGS Change % Value] )

Scn Gross Profit = [Scn Net Sales] - [Scn COGS]
Scn Gross Profit % = DIVIDE ( [Scn Gross Profit], [Scn Net Sales] )

Scn Profit vs Base   = [Scn Gross Profit] - [Gross Profit]
Scn Profit vs Budget = [Scn Gross Profit] - [Budget Profit]
```
Show them as a card row (Base vs Scenario vs Budget), plus a second waterfall using `Scn` measures (duplicate `WF Value` as `WF Value Scn`). Sliders respond to the segment, country, product and year slicers.

## 8. Suggested report pages

1. **Executive Summary:** cards (Gross Sales, Net Sales, Gross Profit, Gross Profit %, COGS Ratio); revenue-erosion waterfall; year / month slicers; Net Sales vs Budget clustered column by month.
2. **Variance Analysis:** matrix with Segment → Country → Product rows and Net Sales, Budget, Var, Var %, Gross Profit %, Margin Var; data bars and conditional-format icons; a bar chart of Profit Var by Country.
3. **Segment & Product Margin:** bar of `Segment Margin` by Segment; scatter (X = Net Sales, Y = Gross Profit %, size = COGS, details = Product); Discount Band × Segment matrix for discount rate.
4. **Scenario Planner:** the four sliders, base vs scenario cards, scenario waterfall, delta vs budget.

Pick a consistent palette (navy, teal; green / red only for variance), turn on tooltips, and add a Reset-slicers bookmark.

## 9. Findings in the supplied data (for your commentary)

Computed from the 700 rows (Sep 2013 to Dec 2014):

| Step | Amount | % of Gross |
|---|---|---|
| Gross Sales | 127.93M | 100% |
| Discounts | (9.21M) | 7.2% |
| **Net Sales** | **118.73M** | 92.8% |
| COGS | (101.83M) | 79.6% |
| **Gross Profit** | **16.89M** | 13.2% |

- Gross Profit % is **14.2%** of Net Sales; COGS Ratio is **85.8%**.
- **Segment margin:** Channel Partners 73.1%, Midmarket 27.7%, Government 21.7%, Small Business 9.8%, **Enterprise –3.1% (loss-making, –0.61M)**.
- Government is the profit engine at about 11.4M of 16.9M profit.
- **High discount band** carries 5.32M of the 9.21M discounts (58%) on about a third of gross sales, so discounting is concentrated there.
- 2013 contains only Sep–Dec, so avoid year-over-year growth claims or flag the partial year.
- Variance vs the **placeholder budget** is illustrative only (budget net sales 120.0M, profit 24.0M); real conclusions need real targets.

## 10. Validation checklist

- `Net Sales` total = **118,726,350** and `Gross Profit` = **16,893,702**.
- `Gross Sales – Discounts = Net Sales` and `Net Sales – COGS = Gross Profit` on every row.
- Slicing by Year 2013 shows Net Sales ≈ 26.4M, 2014 ≈ 92.3M.
- With all sliders at default, `Scn` measures equal the base measures.
- Save as `Financial_KPI_Dashboard.pbix`.
