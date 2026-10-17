# Power BI Superstore Executive Dashboard — Setup Guide

## Files
- `Fact_Orders.csv`: sales fact table at order-line grain.
- `Dim_Customer.csv`: customer and segment dimension.
- `Dim_Product.csv`: product, category and sub-category dimension.
- `Dim_Date.csv`: contiguous calendar dimension.
- `PowerBI_DAX_Measures.txt`: ready-to-copy DAX measures.
- `Superstore_Executive_Dashboard_Report.pdf`: static executive report.

## Build the model in Power BI Desktop
1. Home → Get data → Text/CSV. Import all four CSV files.
2. In Power Query, confirm data types: dates as Date, Sales/Profit/Discount as Decimal Number, Quantity as Whole Number, IDs as Text where applicable.
3. In Model view, create:
   - `Dim_Date[Date]` (1) → `Fact_Orders[Order Date]` (*)
   - `Dim_Customer[Customer ID]` (1) → `Fact_Orders[Customer ID]` (*)
   - `Dim_Product[Product ID]` (1) → `Fact_Orders[Product ID]` (*)
   Use single-direction filtering from dimensions to fact.
4. Mark `Dim_Date` as the Date table using `Dim_Date[Date]`.
5. Create measures by copying expressions from `PowerBI_DAX_Measures.txt`.

## Suggested report layout
### Page 1 — Executive Overview
- Four KPI cards: Total Sales, Total Profit, Profit Margin %, Total Orders.
- Line chart: `Dim_Date[Year-Month]` on X-axis and `[Total Sales]` on Y-axis.
- Treemap: Category and Sub-Category hierarchy; Values = `[Total Sales]`.
- Slicers: Date, Region, Segment, Category.

### Page 2 — Regional & Product Profitability
- Clustered bar chart: Region vs Total Sales and Total Profit.
- Matrix: Rows = Category > Sub-Category; Values = Total Sales, Total Profit, Profit Margin %.
- Conditional formatting on Profit Margin % and Total Profit.
- Scatter plot: Discount on X-axis, Profit on Y-axis, Category in Legend, Sales in Size (optional).

### Page 3 — Discount & Customer Analysis
- Column chart: Average Discount % and Total Profit by discount bands (if you add a discount-band column).
- Bar chart: Segment by Sales and Profit.
- Drill-through or tooltip page for category/sub-category details.

## Important implementation note
A `.pbix` file must be created/saved by Power BI Desktop. This starter pack provides the modeled CSV tables, DAX measures, instructions, and PDF report; it is not itself a `.pbix` file. Once the report is assembled in Power BI Desktop, save as `.pbix` and export the report to PDF using File → Export → Export to PDF.
