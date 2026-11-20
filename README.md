# 📊 Excel Company Sales Dashboard

An end-to-end sales analytics project built entirely in **Microsoft Excel**: messy raw data is cleaned with **Power Query**, modeled as a **star schema in Power Pivot** with **DAX measures**, analysed with **PivotTables**, and presented in an interactive **dashboard with slicers**.

---

## 📌 Table of Contents
1. [Project Overview](#-project-overview)
2. [Tools & Skills](#-tools--skills)
3. [Dataset](#-dataset)
4. [Stage 1 – Data Cleaning (Power Query)](#-stage-1--data-cleaning-power-query)
5. [Stage 2 – Data Modeling (Power Pivot Star Schema)](#-stage-2--data-modeling-power-pivot-star-schema)
6. [Stage 3 – Analysis (PivotTables)](#-stage-3--analysis-pivottables)
7. [Stage 4 – Dashboard](#-stage-4--dashboard)
8. [Key Insights](#-key-insights)
9. [Repository Structure](#-repository-structure)
10. [How to Use](#-how-to-use)

---

## 🎯 Project Overview

The goal is to turn raw, inconsistent sales files into a reliable data model and an interactive dashboard that answers questions such as:

- How much did the company sell in total, and how many orders and customers were there?
- Which product categories, products and regions perform best and worst?
- How do sales and quantities change by month?

**Workflow:**

```
Raw CSV files  →  Power Query (clean & transform)  →  Power Pivot (star schema + DAX)
              →  PivotTables  →  Interactive Dashboard (charts + slicers)
```

---

## 🛠 Tools & Skills

| Area | Tools / Features |
|------|------------------|
| ETL / Cleaning | Power Query (M language), Folder import, Custom Functions |
| Data Modeling | Power Pivot, Star Schema, Relationships (1-to-many) |
| Calculations | DAX Measures, Calendar (Date) table, Date Hierarchy |
| Analysis | PivotTables, PivotCharts, Top/Bottom N filters |
| Visualization | Excel Dashboard, Slicers, KPI cards, combo/column/bar/pie/line charts |

---

## 📂 Dataset

The data comes from three sources:

| Source | Format | Description |
|--------|--------|-------------|
| **Customers** | `Customers.csv` | Customer ID, name, and region |
| **Products** | `Products.csv` | Product ID, name, category, and price |
| **Orders** | Folder of monthly files (`..._<month>.csv`) | Order lines: customer, product, date, quantity |

**Regions:** Alexandria, Aswan, Cairo, Giza, Mansoura
**Period analysed:** 2024

---

## 🧹 Stage 1 – Data Cleaning (Power Query)

All cleaning was done in Power Query so the process is repeatable: when source files change, a single **Refresh** reruns every step.

### 👥 Customers table

| Step | Purpose |
|------|---------|
| Promote headers / change data types | `CustomerID` → whole number, other columns → text |
| Capitalize each word (`Text.Proper`) + Trim | Standardise name formatting and remove extra spaces |
| Split `CustomerName` by space | Create **First Name** and **Last Name** columns |
| Replace spaces in `Region` + Capitalize | Make region names consistent (e.g., no stray spaces / mixed case) |

```m
let
    Source = Csv.Document(File.Contents(".../Customers.csv"),[Delimiter=",", Columns=3, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CustomerID", Int64.Type}, {"CustomerName", type text}, {"Region", type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Changed Type",{{"CustomerName", Text.Proper, type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Capitalized Each Word",{{"CustomerName", Text.Trim, type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Trimmed Text", "CustomerName", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"CustomerName.1", "CustomerName.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"CustomerName.1", type text}, {"CustomerName.2", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"CustomerName.1", "First Name"}, {"CustomerName.2", "Last Name"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns"," ","",Replacer.ReplaceText,{"Region"}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Replaced Value",{{"Region", Text.Proper, type text}})
in
    #"Capitalized Each Word1"
```

### 📦 Products table

**Problem:** the `Price` column contained dirty values such as `1,296$`, `872$`, `1,423$`, `496$` and even text like `Two Hundred`.

| Step | Purpose |
|------|---------|
| Capitalize + Trim `ProductName` and `Category` | Consistent text values |
| Replace `1,296$` → `1296`, `872$` → `872`, `1,423$` → `1423`, `496$` → `496` | Remove currency symbols and thousand separators |
| Replace `Two Hundred` → `200` | Convert a text value to a number |
| Change `Price` type to **Currency** | Make the column usable in calculations |

```m
let
    Source = Csv.Document(File.Contents(".../Products.csv"),[Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ProductID", Int64.Type}, {"ProductName", type text}, {"Category", type text}, {"Price", type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Changed Type",{{"ProductName", Text.Proper, type text}, {"Category", Text.Proper, type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Capitalized Each Word",{{"Category", Text.Trim, type text}, {"ProductName", Text.Trim, type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Trimmed Text","1,296$","1296",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","Two Hundred","200",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","872$","872",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","1,423$","1423",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","496$","496",Replacer.ReplaceText,{"Price"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value4",{{"Price", Currency.Type}})
in
    #"Changed Type1"
```

### 🧾 Orders table (the most complex one)

Orders arrive as **multiple monthly files in a folder**, with **inconsistent date formats** and **missing values**.

| Step | Purpose |
|------|---------|
| Import from **Folder** + filter hidden files | Load all monthly order files at once |
| Invoke **Custom Function** (`Transform File`) | Apply the same transformation to every file and combine them |
| Expand the table column | Merge all files into a single table |
| Filter out null/empty `ProductID` and `Quantity` | Remove incomplete rows |
| Remove original `OrderID`, add **Index** column → `Order ID` | Rebuild a clean, unique order key |
| Replace `/` with `-` in `OrderDate` | Unify date separators |
| Split `OrderDate` by `-` into 3 parts | Separate day / month / year components (formats vary) |
| Conditional column **Year** | The part greater than 31 must be the year |
| Extract **Month** from the file name (`Text.BetweenDelimiters` between `_` and `.`) | The month is reliable in the file name |
| Conditional column **Day** | The part ≤ 31 that is not the month is the day |
| Merge Day-Month-Year (culture `en-GB`) → `new date`, change type to **Date** | One clean date column used to relate to the Calendar table |

```m
let
    Source = Folder.Files(".../Orders"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"OrderID", Int64.Type}, {"CustomerID", Int64.Type}, {"ProductID", Int64.Type}, {"OrderDate", type text}, {"Quantity", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each [ProductID] <> null and [ProductID] <> ""),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each [Quantity] <> null and [Quantity] <> ""),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows1",{"OrderID"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns", "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "Order ID"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Source.Name", "Order ID", "CustomerID", "ProductID", "OrderDate", "Quantity"}),
    #"Replaced Value" = Table.ReplaceValue(#"Reordered Columns","/","-",Replacer.ReplaceText,{"OrderDate"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Replaced Value", "OrderDate", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"OrderDate.1", "OrderDate.2", "OrderDate.3"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"OrderDate.1", Int64.Type}, {"OrderDate.2", Int64.Type}, {"OrderDate.3", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Year", each if [OrderDate.1] > 31 then [OrderDate.1] else if [OrderDate.3] > 31 then [OrderDate.3] else "not found"),
    #"Extracted Text Between Delimiters" = Table.TransformColumns(#"Added Conditional Column", {{"Source.Name", each Text.BetweenDelimiters(_, "_", "."), type text}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Extracted Text Between Delimiters",{{"Source.Name", Int64.Type}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type2",{{"Source.Name", "Month"}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns2",{"Order ID", "CustomerID", "ProductID", "OrderDate.1", "OrderDate.2", "OrderDate.3", "Quantity", "Year", "Month"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns1", "Day", each if [OrderDate.1]<=31 and [OrderDate.1]<> [Month] then [OrderDate.1]
        else if [OrderDate.2]<=31 and [OrderDate.2]<> [Month] then [OrderDate.2]
        else if [OrderDate.3]<=31 and [OrderDate.3]<> [Month] then [OrderDate.3]
        else [Month]),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom",{"OrderDate.1", "OrderDate.2", "OrderDate.3"}),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Removed Columns1", {{"Day", type text}, {"Month", type text}, {"Year", type text}}, "en-GB"),{"Day", "Month", "Year"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"new date"),
    #"Changed Type3" = Table.TransformColumnTypes(#"Merged Columns",{{"new date", type date}})
in
    #"Changed Type3"
```

> The query above is condensed from the full applied-steps file; a few no-op "filter rows" steps were omitted.

---

## 🧩 Stage 2 – Data Modeling (Power Pivot Star Schema)

The cleaned tables were loaded into the **Excel Data Model** and arranged as a **star schema**: one central **fact table** (Orders) surrounded by **dimension tables**.

<img width="645" height="415" alt="new star" src="https://github.com/user-attachments/assets/40172c59-3274-45e4-8254-c3d0fdb6b5c4" />


| Table | Type | Columns |
|-------|------|---------|
| **Orders** | Fact | Order ID, CustomerID, ProductID, Quantity, new date + measures |
| **Products** | Dimension | ProductID, ProductName, Category, Price |
| **Customers** | Dimension | CustomerID, First Name, Last Name, Region |
| **Calendar** | Dimension (Date table) | Date, Year, Month Number, Month, MMM-YYYY, Day Of Week Number, Day Of Week, Date Hierarchy (Year → Month → Date) |

**Relationships (one-to-many, single direction):**

- `Products[ProductID]` **1 ──► \*** `Orders[ProductID]`
- `Customers[CustomerID]` **1 ──► \*** `Orders[CustomerID]`
- `Calendar[Date]` **1 ──► \*** `Orders[new date]`

**DAX measures created in the Orders table:**

| Measure | Description |
|---------|-------------|
| `Total sales` | Sum of quantity × product price |
| `Total quantity` | Sum of quantity sold |
| `Total orders` | Number of orders |
| `Total customer` | Distinct count of customers |

Example DAX (adjust to your own model):

```dax
Total quantity = SUM ( Orders[Quantity] )
Total orders   = DISTINCTCOUNT ( Orders[Order ID] )
Total customer = DISTINCTCOUNT ( Orders[CustomerID] )
Total sales    = SUMX ( Orders, Orders[Quantity] * RELATED ( Products[Price] ) )
```

**Why a star schema?** It keeps filters flowing from dimensions to the fact table, makes DAX simpler and faster, avoids duplicated data, and makes the model easy to extend.

---

## 📑 Stage 3 – Analysis (PivotTables)

PivotTables built on the Data Model feed every chart and KPI card on the dashboard.

<img width="527" height="312" alt="pivot table" src="https://github.com/user-attachments/assets/86c47c72-9767-4074-9ab3-274230503ff2" />


| PivotTable | Content |
|------------|---------|
| KPI cards | Total customers, total orders, total quantity, total sales |
| Sales by month | Total sales and total orders per month |
| Sales by category | Total sales per product category |
| Sales by region | Total sales per region |
| Customers by region | Distinct customers per region |
| Top 5 products | Top-N filter on total sales |
| Bottom 5 products | Bottom-N filter on total sales |
| Quantity per product | Total quantity per product (filtered to 2024) |

---

## 📈 Stage 4 – Dashboard

The final **Company Sales Dashboard** combines KPI cards, charts, and slicers in a single interactive view.

<img width="454" height="257" alt="Sales Dashboard" src="https://github.com/user-attachments/assets/724e10ec-5623-4e4d-98a2-1de154480407" />


**KPIs**

| Total Sales | Total Customers | Total Orders | Total Quantity |
|:-----------:|:---------------:|:------------:|:--------------:|
| **$101,146,137** | **300** | **19,594** | **108,120** |

**Visuals**

- 🥧 **Total Sales per Category** – pie chart
- 📊 **Top & Bottom Region Sales** – bar chart
- 🏆 **Top 5 Sales Products** – bar chart
- 📉 **Bottom 5 Sales Products** – bar chart
- 📅 **Sales Against Orders (by month)** – combo chart (columns + line)
- 👥 **Total Customers per Region** – column chart
- 📈 **Sales Quantity per Item** – smooth line chart

**Slicers (Filters):** Category, ProductName, Region, Month — all connected to every PivotTable/PivotChart so the whole dashboard updates together.

---

## 🔍 Key Insights

- **Clothing** is the top category (~$29.1M), followed by **Toys** (~$25.4M) and **Electronics** (~$19.0M); **Furniture** is the lowest (~$10.4M).
- **Top 5 products:** Doll (~$8.15M), Shirt (~$7.21M), Pants (~$6.57M), Tablet (~$6.35M), Shoes (~$6.33M).
- **Bottom 5 products:** Sofa (~$0.26M), Headphones (~$0.89M), Chair (~$1.93M), Notebook (~$1.98M), Lego (~$2.06M).
- **Mansoura** has the highest regional sales (~$21.0M) while **Alexandria** has the lowest (~$18.7M).
- Customers are spread almost evenly across regions (56–62 per region).
- Monthly sales are stable (~$8.3M–$8.7M), peaking in **May** and lowest in **February**.
- Tablet has the highest units sold (4,637) and Phone the lowest (3,894).

---

## 🗂 Repository Structure

```
Excel-Company-Sales-Dashboard/
│
├── README.md
├── Company_Sales_Dashboard.xlsx      # Final workbook (Power Query + Power Pivot + Dashboard)
├── data/
│   ├── Customers.csv
│   ├── Products.csv
│   └── Orders/                       # Monthly order files
├── power-query/
│   └── applied_steps.txt             # M code for Customers, Products, Orders
└── images/
    ├── star_schema.png
    ├── pivot_table.png
    └── sales_dashboard.png
```

---

## ▶️ How to Use

1. Download the repository and open `Company_Sales_Dashboard.xlsx` in **Excel (Windows, 2016 or later / Microsoft 365)** — Power Pivot is required.
2. If prompted, update the source file paths in **Data → Queries & Connections** to point to the `data/` folder.
3. Click **Data → Refresh All** to reload the data.
4. Use the slicers on the dashboard to explore sales by category, product, region, and month.

---

## 👤 Author

**Your Name** — [LinkedIn](https://linkedin.com/in/your-profile) · [GitHub](https://github.com/your-username)

⭐ If you found this project useful, please give it a star!
