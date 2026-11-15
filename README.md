# Sales Performance Dashboard 2024

An interactive **Excel dashboard** that analyzes 2024 sales across five Egyptian regions (Alexandria, Aswan, Cairo, Giza, Mansoura). The project covers the full analytics workflow: cleaning messy source data with **Power Query**, building a **star-schema data model** in Power Pivot, writing **DAX measures**, and presenting the results with pivot charts and slicers.

---

## Key Metrics

| KPI | Value |
|-----|-------|
| Total Sales | $101,208,243 |
| Total Quantity | 108,203 |
| Total Orders | 19,826 |
| Customers | 300 |

---

## Dashboard Visuals
<img width="1305" height="474" alt="image" src="https://github.com/user-attachments/assets/32caccd8-f5b3-4da5-a5cc-d11ca9ea938b" />

| Visual | Chart type | Fields |
|--------|-----------|--------|
| Total Sales per Region | Donut | `Customers[Region]`, Total Sales |
| Regional Sales by Product Category | Stacked column | Category (axis), Region (legend), Total Sales |
| Customer Distribution by Region | Pie | `Customers[Region]`, Customers count |
| Monthly Order Volume Peaks | Line (smoothed) | `Calendar[Month]`, Orders count |
| Top 5 Best-Selling | Bar | Total Sales (Top 5 filter) |
| Bottom 5 Lowest-Selling Products | Bar | `Products[ProductName]`, Total Sales (Bottom 5 filter) |
| Filters | Slicers | Region, Category, Date (timeline) |

---

## Tools & Skills

- **Excel** (pivot tables, pivot charts, slicers, timeline)
- **Power Query (M)** for data cleaning and transformation
- **Power Pivot** for the data model and relationships
- **DAX** for calculated measures

---

## Data Model
<img width="978" height="548" alt="image" src="https://github.com/user-attachments/assets/f9080c0c-f7a7-4d3b-aac0-22ecfb894f0c" />
The model is a star schema: one fact table (`Orders`) and three dimension tables.

```
Products (1) ──────┐
                   ├──< Orders (*) >──── (1) Calendar
Customers (1) ─────┘
```

| Table | Type | Key columns |
|-------|------|-------------|
| Orders | Fact | Order ID, CustomerID, ProductID, Quantity, New Date |
| Customers | Dimension | CustomerID, First Name, Last Name, Region |
| Products | Dimension | ProductID, ProductName, Category, Price |
| Calendar | Dimension | Date, Year, Month Number, Month, MMM-YYYY, Day Of Week Number, Day Of Week |

**Relationships**

- `Products[ProductID]` → `Orders[ProductID]` (one-to-many)
- `Customers[CustomerID]` → `Orders[CustomerID]` (one-to-many)
- `Calendar[Date]` → `Orders[New Date]` (one-to-many)

---

## Data Cleaning (Power Query)

### Orders
- Combined multiple monthly files from a folder (`Folder.Files` + a custom transform function)
- Extracted the **month** from each source file name
- Removed duplicate orders and rows with a missing `ProductID`
- Fixed **mixed date formats** (`/` and `-` separators, day/month/year in different positions) by splitting the date, detecting the year and day, and rebuilding a clean `New Date` column using an explicit locale
- Generated a new sequential `Order ID` and removed helper columns

### Products
- Removed duplicate `ProductID` values
- Trimmed and capitalized `ProductName` and `Category`
- Fixed text prices (for example, "Two Hundred" → 200), removed the `$` sign, and converted `Price` to a currency type

### Customers
- Removed duplicate `CustomerID` values
- Trimmed and capitalized names, then split `CustomerName` into **First Name** and **Last Name**
- Cleaned `Region` (removed spaces, proper case)

---

## DAX Measures

```DAX
Total Sales =
SUMX ( Orders, Orders[Quantity] * RELATED ( Products[Price] ) )

Total Quantity =
SUM ( Orders[Quantity] )

Orders Count =
COUNTROWS ( Orders )

Average Sales Per Order =
DIVIDE ( [Total Sales], [Orders Count] )
```

---

## Power Query Source Code

<details>
<summary>Orders query</summary>

```m
let
    Source = Folder.Files("C:\...\PowerQuery\Orders"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"OrderID", Int64.Type}, {"CustomerID", Int64.Type}, {"ProductID", Int64.Type}, {"OrderDate", type text}, {"Quantity", Int64.Type}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"OrderID"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Duplicates", each [ProductID] <> null and [ProductID] <> ""),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","/","-",Replacer.ReplaceText,{"OrderDate"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Replaced Value", "OrderDate", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"OrderDate.1", "OrderDate.2", "OrderDate.3"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"OrderDate.1", Int64.Type}, {"OrderDate.2", Int64.Type}, {"OrderDate.3", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Year", each if [OrderDate.1] >= 2000 then [OrderDate.1] else if [OrderDate.3] >= 2000 then [OrderDate.3] else "Not found"),
    #"Filtered Rows1" = Table.SelectRows(#"Added Conditional Column", each true),
    #"Extracted Text Between Delimiters" = Table.TransformColumns(#"Filtered Rows1", {{"Source.Name", each Text.BetweenDelimiters(_, "_", "."), type text}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Extracted Text Between Delimiters",{{"Source.Name", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type2",{{"Source.Name", "Month"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"OrderID", "CustomerID", "ProductID", "OrderDate.1", "OrderDate.2", "OrderDate.3", "Quantity", "Year", "Month"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns", "Custom", each if[OrderDate.1]<=31 and [OrderDate.1]<>[Month] then [OrderDate.1]
else if [OrderDate.2]<=31 and [OrderDate.2]<>[Month] then [OrderDate.2]
else if [OrderDate.3]<=31 and [OrderDate.3]<>[Month] then [OrderDate.3]
else
[Month]),
    #"Renamed Columns2" = Table.RenameColumns(#"Added Custom",{{"Custom", "Day"}}),
    #"Inserted Merged Column" = Table.AddColumn(#"Renamed Columns2", "New Date", each Text.Combine({Text.From([Day], "en-US"), Text.From([Month], "en-US"), Text.From([Year], "en-US")}, "-"), type text),
    #"Changed Type with Locale" = Table.TransformColumnTypes(#"Inserted Merged Column", {{"New Date", type date}}, "en-AE"),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type with Locale",{"OrderDate.1", "OrderDate.2", "OrderDate.3", "Year", "Month", "Day", "OrderID"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Index",{"Index", "CustomerID", "ProductID", "Quantity", "New Date"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Reordered Columns1",{{"Index", "Order ID"}})
in
    #"Renamed Columns3"
```

</details>

<details>
<summary>Products query</summary>

```m
let
    Source = Csv.Document(File.Contents("C:\...\PowerQuery\Products.csv"),[Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ProductID", Int64.Type}, {"ProductName", type text}, {"Category", type text}, {"Price", type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"ProductID"}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Removed Duplicates",{{"ProductName", Text.Proper, type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Capitalized Each Word",{{"ProductName", Text.Trim, type text}}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Trimmed Text",{{"Category", Text.Proper, type text}}),
    #"Trimmed Text1" = Table.TransformColumns(#"Capitalized Each Word1",{{"Category", Text.Trim, type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Trimmed Text1","Two Hundred","200",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","$","",Replacer.ReplaceText,{"Price"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value1",{{"Price", Currency.Type}})
in
    #"Changed Type1"
```

</details>

<details>
<summary>Customers query</summary>

```m
let
    Source = Csv.Document(File.Contents("C:\...\PowerQuery\Customers.csv"),[Delimiter=",", Columns=3, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CustomerID", Int64.Type}, {"CustomerName", type text}, {"Region", type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"CustomerID"}),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Duplicates",{{"CustomerName", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text",{{"CustomerName", Text.Proper, type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Capitalized Each Word", "CustomerName", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"CustomerName.1", "CustomerName.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"CustomerName.1", type text}, {"CustomerName.2", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"CustomerName.1", "First Name"}, {"CustomerName.2", "Last Name"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns"," ","",Replacer.ReplaceText,{"Region"}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Replaced Value",{{"Region", Text.Proper, type text}}),
    #"Trimmed Text1" = Table.TransformColumns(#"Capitalized Each Word1",{{"Region", Text.Trim, type text}})
in
    #"Trimmed Text1"
```

</details>

---

## Repository Structure

```
sales-performance-dashboard-2024/
├── Sales_Performance_Dashboard_2024.xlsx
├── data/
│   ├── Orders/          # monthly order files
│   ├── Products.csv
│   └── Customers.csv
├── power-query/
│   ├── orders.m
│   ├── products.m
│   └── customers.m
├── images/
│   └── dashboard.png
└── README.md
```

---

## How to Use

1. Download or clone this repository.
2. Open `Sales_Performance_Dashboard_2024.xlsx` in Excel (Excel 2016 or later with Power Query and Power Pivot).
3. If the data does not refresh, go to **Data → Get Data → Query Options** and update the source file paths to your local folders, then click **Refresh All**.
4. Use the **Region**, **Category**, and **Date** slicers to filter the dashboard.

---

## Insights

- Sales are spread fairly evenly across the five regions (about 18% to 21% each).
- Customer counts per region are close to equal (56 to 62).
- Clothing and Toys are the highest-selling categories.
- Monthly sales stay in a narrow range (about $8.3M to $8.7M), with the highest month in May and April.

---

## Author

**[Mahmoud Ahmed]**
Technical Translator| Excel, Power Query & Dashboards
[www.linkedin.com/in/mahmoud-hashem-b8598852

](https://www.linkedin.com/) · [Email](mailto:M_translator@ymail.com)
