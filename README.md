# 📊 Superstore Sales & Profitability Analysis

An end-to-end **Excel Data Analytics project** focused on analyzing sales, profitability, customers, products, returns, discounts, and regional performance.

The project demonstrates how raw and uncleaned business data can be transformed into a structured analytical dataset, KPI calculations, PivotTable analysis, and an interactive Excel dashboard that supports business decision-making.

---

## 📌 Project Overview

This project analyzes a Superstore-style retail dataset covering **2014–2017**.

The analysis answers key business questions such as:

* How are sales and profit growing over time?
* Which regions generate the most sales and profit?
* Which categories and sub-categories are most profitable?
* Which products generate losses?
* What is the return rate?
* Which categories and regions have higher return rates?
* How does discounting affect profitability?
* Which customer segments contribute the most revenue and profit?
* Which products should be investigated for pricing or profitability issues?

---

# 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Tables
* Power Query / Data Transformation
* PivotTables
* PivotCharts
* Excel Formulas
* Slicers
* Conditional Formatting
* Data Visualization
* Business Analysis

---

# 📂 Dataset Structure

The original workbook contains three main data sources:

```text
📁 Raw Data
│
├── Orders
│   ├── Order ID
│   ├── Order Date
│   ├── Ship Date
│   ├── Customer
│   ├── Product
│   ├── Category
│   ├── Sub-Category
│   ├── Sales
│   ├── Quantity
│   ├── Discount
│   ├── Profit
│   └── Region
│
├── Returns
│   ├── Order ID
│   └── Returned
│
└── People
    ├── Person
    └── Region
```

### Purpose of each table

**Orders**
Contains the main transaction-level sales data.

**Returns**
Contains information about whether an order was returned.

**People**
Maps people/managers to their respective regions.

---

# 🔄 Project Workflow

The complete project follows this workflow:

```text
Raw / Uncleaned Data
        ↓
Data Cleaning
        ↓
Data Transformation
        ↓
Merge Orders + Returns
        ↓
Create Calculated Columns
        ↓
Create KPI Values
        ↓
Create PivotTables
        ↓
Business Analysis
        ↓
Dashboard
        ↓
Business Insights
        ↓
Recommendations
```

---

# 1️⃣ Raw / Uncleaned Data

The project starts with raw Excel data.

The first step is to inspect:

* Missing values
* Duplicate records
* Incorrect data types
* Date formatting
* Text inconsistencies
* Numeric fields
* Column names
* Unnecessary columns
* Return information
* Region mapping

Before performing analysis, the raw data should not be modified directly.

Keep the original sheets as a backup/reference.

---

# 2️⃣ Data Cleaning

The Orders dataset was prepared for analysis by checking and standardizing:

### Date columns

Convert:

```text
08/11/2016 00:00:00
```

into a proper Excel date.

Then create:

```text
Year
Month-Year
```

Example:

```excel
=TEXT([@[Order Date]], "yyyy")
```

For monthly trend analysis:

```excel
=TEXT([@[Order Date]],"mmm-yyyy")
```

---

## Numeric columns

Verify that the following fields are numeric:

```text
Sales
Quantity
Discount
Profit
```

Discount should be stored as a decimal/percentage:

```text
0      → 0%
0.10   → 10%
0.20   → 20%
0.50   → 50%
```

---

# 3️⃣ Convert Data into Excel Tables

Convert the raw ranges into Excel Tables:

**Insert → Table**

Recommended table names:

```text
Orders
Returns
```

Using Excel Tables makes formulas, PivotTables, and future data refreshes easier.

---

# 4️⃣ Merge Orders and Returns

The `Returns` table is connected to the Orders data using:

```text
Order ID
```

The purpose is to identify which orders were returned.

The resulting master dataset contains information such as:

```text
Order ID
Sales
Profit
Category
Sub-Category
Region
Discount
Returned
```

---

The final master dataset can therefore support:

```text
Region
Sales
Profit
Orders
Returns
Return Rate
```

---

# 6️⃣ Create the Orders Master Table

After cleaning and combining the data, create an `Orders_Master` table.

Example structure:

```text
Orders_Master
│
├── Order ID
├── Order Date
├── Year
├── Month-Year
├── Customer
├── Category
├── Sub-Category
├── Product
├── Sales
├── Quantity
├── Discount
├── Profit
├── Region
├── Returned
├── Unique Order Flag
├── Unique Returned Order Flag
└── Discount Band
```

This becomes the main dataset used for analysis.

---

# 7️⃣ Create Unique Order Flag

The dataset can contain multiple rows for the same Order ID because one order can contain multiple products.

To count unique orders correctly, create:

```text
Unique Order Flag
```

Formula:

```excel
=IF(COUNTIFS(B$2:B2,B2)=1,1,0)
```

The first occurrence of an Order ID receives:

```text
1
```

Other occurrences receive:

```text
0
```

Then:

```text
SUM(Unique Order Flag)
```

gives the number of unique orders.

---

# 8️⃣ Create Unique Returned Order Flag

To calculate unique returned orders, create:

```text
Unique Returned Order Flag
```

The flag should be `1` only for the first occurrence of a returned Order ID.

This allows:

```text
Returned Orders
```

to be calculated without counting the same returned order multiple times.

---

# 9️⃣ Create Discount Band

To analyze the effect of discounting, create:

```text
Discount Band
```

Example:

```text
0%
1–10%
11–20%
21–30%
31–40%
40%+
```

Example formula:

```excel
=IF([@Discount]=0,"0%",
IF([@Discount]<=10%,"1-10%",
IF([@Discount]<=20%,"11-20%",
IF([@Discount]<=30%,"21-30%",
IF([@Discount]<=40%,"31-40%","40%+")))))
```

This allows profitability to be analyzed by discount level.

---

# 🔟 KPI Calculations

Create a separate:

```text
KPI Value
```

sheet.

Recommended KPIs:

| KPI                 | Calculation                     |
| ------------------- | ------------------------------- |
| Total Sales         | SUM(Sales)                      |
| Total Profit        | SUM(Profit)                     |
| Total Orders        | SUM(Unique Order Flag)          |
| Total Customers     | Unique Customer ID              |
| Total Products      | Unique Product                  |
| Total Quantity      | SUM(Quantity)                   |
| Returned Orders     | SUM(Unique Returned Order Flag) |
| Return Rate         | Returned Orders / Orders        |
| Profit Margin       | Profit / Sales                  |
| Average Order Value | Sales / Orders                  |
| Average Discount    | AVERAGE(Discount)               |

---

# 📊 KPI Formulas

### Total Sales

```excel
=SUM(Orders[Sales])
```

### Total Profit

```excel
=SUM(Orders[Profit])
```

### Total Orders

```excel
=SUM(Orders[Unique Order Flag])
```

### Returned Orders

```excel
=SUM(Orders[Unique Returned Order Flag])
```

### Return Rate

```excel
=Returned Orders / Total Orders
```

### Profit Margin

```excel
=Total Profit / Total Sales
```

### Average Order Value

```excel
=Total Sales / Total Orders
```

### Average Discount

```excel
=AVERAGE(Orders[Discount])
```

---

# 📈 1. Sales Trend Analysis

Create a PivotTable using:

### Rows

```text
Year
Month-Year
```

### Values

```text
Sales
Profit
```

Create a:

**Line Chart**

to analyze sales and profit trends over time.

Questions answered:

* Is revenue growing?
* Which year had the highest sales?
* Are sales and profit growing together?
* Are there seasonal patterns?

---

# 🌎 2. Regional Analysis

Create a PivotTable:

### Rows

```text
Region
```

### Values

```text
Sales
Profit
Unique Order Flag
```

Create:

* Sales by Region
* Profit by Region
* Return Rate by Region

This helps identify strong and weak regional performance.

---

# 🏷️ 3. Category Analysis

Create a PivotTable:

### Rows

```text
Category
```

### Values

```text
Sales
Profit
Quantity
```

Recommended chart:

**Sales & Profit by Category**

This helps identify categories with:

```text
High Sales + High Profit
High Sales + Low Profit
Low Sales + High Profit
Loss-Making Categories
```

---

# 📦 4. Sub-Category Analysis

Create a PivotTable:

### Rows

```text
Sub-Category
```

### Values

```text
Sales
Profit
Quantity
```

Recommended charts:

* Sales by Sub-Category
* Profit by Sub-Category

This provides a deeper product-performance analysis.

---

# 🔄 5. Return Analysis

Create PivotTables for:

### Returns by Category

```text
Rows:
Category

Values:
Unique Returned Order Flag
```

### Return Rate by Region

```text
Region
Returned Orders
Total Orders
Return Rate
```

Formula:

```excel
=Returned Orders / Total Orders
```

Recommended visuals:

* Returned Orders by Category
* Return Rate by Region
* Return Rate by Category

---

# 👥 6. Regional Performance

Use the People mapping to analyze:

```text
Region
Sales
Profit
Orders
Returned Orders
Return Rate
```

This provides a management-performance view.

Do not evaluate performance using sales alone. Compare:

```text
Sales
Profit
Margin
Orders
Returns
```

---

# 🏷️ 7. Discount & Profitability Analysis

Create a PivotTable using:

### Rows

```text
Discount Band
```

### Values

```text
Sales
Profit
Unique Order Flag
```

Then calculate:

```text
Profit Margin = Profit / Sales
```

Recommended table:

| Discount Band | Sales | Profit | Profit Margin |
| ------------- | ----: | -----: | ------------: |
| 0%            |   ... |    ... |          ...% |
| 1–10%         |   ... |    ... |          ...% |
| 11–20%        |   ... |    ... |          ...% |
| 21–30%        |   ... |    ... |          ...% |
| 31–40%        |   ... |    ... |          ...% |
| 40%+          |   ... |    ... |          ...% |

Recommended charts:

* Sales by Discount Band
* Profit by Discount Band

This identifies whether aggressive discounting is creating or destroying profitability.

---

# 📊 Dashboard Design

The final dashboard combines the analysis into a single management view.

## KPI Cards

Recommended cards:

```text
Total Sales
Total Profit
Total Orders
Total Customers
Profit Margin
Returned Orders
Return Rate
Average Order Value
Average Discount
```

---

## Dashboard Layout

```text
┌─────────────────────────────────────────────────────────────┐
│           SUPERSTORE SALES & PROFITABILITY                  │
│                                                             │
│ Year | Region | Category | Segment | Person                │
├──────────┬──────────┬──────────┬──────────┬────────────────┤
│ Sales    │ Profit   │ Orders   │ Customers│ Profit Margin  │
├──────────┼──────────┼──────────┼──────────┼────────────────┤
│ Returns  │ Return   │ Avg      │ Avg      │                │
│ Orders   │ Rate     │ Order    │ Discount │                │
├──────────────────────────────┬──────────────────────────────┤
│ Sales Trend                  │ Sales by Region              │
├──────────────────────────────┼──────────────────────────────┤
│ Category Performance         │ Profit by Region             │
├──────────────────────────────┼──────────────────────────────┤
│ Top Products                 │ Return Analysis              │
├──────────────────────────────┼──────────────────────────────┤
│ Sub-Category Analysis        │ Discount vs Profit           │
├─────────────────────────────────────────────────────────────┤
│                    Business Insights                         │
└─────────────────────────────────────────────────────────────┘
```

---

# 💡 Business Insights

The analysis should go beyond simply displaying charts.

### Sales Growth

Analyze year-over-year sales and profit growth.

### Regional Performance

Identify regions with strong sales but weak profitability.

### Product Performance

Identify:

* Top-selling products
* Most profitable products
* Loss-making products

### Returns

Identify categories, regions, and products with high return rates.

### Discount

Analyze whether higher discounts are associated with lower profitability.

### Customer Segments

Compare:

```text
Consumer
Corporate
Home Office
```

using Sales, Profit, Orders, and Margin.

---

# 🚨 Key Business Questions

The final analysis should answer:

1. Which year generated the highest sales?
2. Which region generates the most profit?
3. Which category is most profitable?
4. Which sub-category generates losses?
5. Which products generate the highest profit?
6. Which products generate losses?
7. What is the overall return rate?
8. Which region has the highest return rate?
9. Which category has the highest return rate?
10. Does higher discount reduce profitability?
11. Which customer segment generates the highest profit?
12. Where should management focus to improve profitability?

---

# 🎯 Business Recommendations

Based on the analysis, management can investigate:

### 1. Improve Furniture Profitability

Furniture has relatively strong sales but weak profitability.

Investigate:

* Pricing
* Product costs
* Discounts
* Returns
* Product mix

### 2. Control High Discounts

High discount bands should be reviewed because aggressive discounting can significantly reduce margins.

### 3. Investigate Loss-Making Products

Create a dedicated:

```text
Loss-Making Products
```

table containing:

```text
Product
Sales
Profit
Profit Margin
Discount
Returns
```

### 4. Reduce Returns

Investigate recurring return patterns by:

```text
Product
Category
Region
Customer
```

### 5. Focus on Profitable Growth

Do not optimize only for:

```text
Sales
```

Track:

```text
Sales + Profit + Margin + Returns
```

together.

---

# 📁 Recommended Final Workbook Structure

```text
Superstore_Analysis.xlsx
│
├── Orders
├── Returns
├── People
│
├── Orders_Master
│
├── KPI Value
│
├── Sales Analysis
├── Return Analysis
├── Sub-Category Analysis
├── People Performance
├── Discount Analysis
│
└── Dashboard
```

---

# 🔍 Data-to-Dashboard Architecture

```text
              RAW DATA
                  │
        ┌─────────┼─────────┐
        ↓         ↓         ↓
     Orders    Returns    People
        │         │         │
        └─────────┼─────────┘
                  ↓
          DATA CLEANING
                  ↓
        DATA TRANSFORMATION
                  ↓
           ORDERS_MASTER
                  │
        ┌─────────┼─────────┐
        ↓         ↓         ↓
      KPIs    PivotTables  Analysis
        │         │         │
        └─────────┼─────────┘
                  ↓
             DASHBOARD
                  ↓
          BUSINESS INSIGHTS
                  ↓
          BUSINESS ACTIONS
```

---

# 🚀 Skills Demonstrated

This project demonstrates practical skills in:

* Excel Data Cleaning
* Data Transformation
* Data Integration
* Excel Tables
* Lookup / Merge Logic
* Calculated Columns
* Excel Formulas
* KPI Development
* PivotTables
* PivotCharts
* Dashboard Design
* Business Intelligence
* Sales Analysis
* Profitability Analysis
* Return Analysis
* Discount Analysis
* Product Analysis
* Regional Analysis
* Business Decision Support

---

# 📌 Project Outcome

The final deliverable transforms raw transactional data into a structured business intelligence solution.

The project demonstrates the complete analytics lifecycle:

```text
Raw Data
   ↓
Clean Data
   ↓
Integrated Data
   ↓
Calculated Metrics
   ↓
KPI Analysis
   ↓
Business Analysis
   ↓
Interactive Dashboard
   ↓
Actionable Insights
```

This project can be used as an **Excel Data Analyst / MIS / Business Intelligence portfolio project** and can later be recreated in **Power BI or SQL + Power BI** for a more advanced analytics workflow.
