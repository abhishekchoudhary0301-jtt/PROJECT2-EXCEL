::: {align="center"}
# 📊 SALES PERFORMANCE & CUSTOMER ANALYTICS

### Excel Data Analysis Project \| PR. 2 Analyzer

```{=html}
<p>
```
`<img src="https://img.shields.io/badge/Microsoft-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Microsoft Excel">`{=html}
`<img src="https://img.shields.io/badge/Analysis-Data%20Analytics-2563EB?style=for-the-badge" alt="Data Analytics">`{=html}
`<img src="https://img.shields.io/badge/Status-Completed-16A34A?style=for-the-badge" alt="Completed">`{=html}
```{=html}
</p>
```
**Turning raw sales records into meaningful business insights.**
:::

------------------------------------------------------------------------

## 👨‍💻 Project Information

  Field               Details
  ------------------- -----------------------------------------------------
  **Project Title**   Sales Performance & Customer Analytics
  **Author**          **ABHISHEK CHOUDHARY**
  **Project Type**    Microsoft Excel -- Data Analysis
  **Workbook**        `PR_2_Analyzer_Completed.xlsx`
  **Dataset Size**    200 sales records
  **Tools**           Microsoft Excel, Formulas, What-If Analysis, Charts
  **Domain**          Sales, Retail & Customer Analytics

------------------------------------------------------------------------

## 📌 1. Project Overview

This project analyzes a sales dataset containing customer details,
regions, product categories, sales amounts, quantities, discounts, order
dates, and profit.

The objective is to transform raw transactional data into an organized
analysis workbook. Excel formulas and analytical features are used to
understand sales performance, evaluate discount decisions, identify
valuable customers, study profit patterns, and communicate findings
through a dashboard.

The workbook is designed as a practical demonstration of core and
intermediate Excel data-analysis skills.

## 🎯 2. Project Objectives

-   Analyze overall sales, profit, orders, and discount performance.
-   Highlight the top 10 customers based on total purchase.
-   Measure how different discount rates affect total profit.
-   Study the relationship between sales and profit using linear
    regression.
-   Generate descriptive statistics for important numeric fields.
-   Track monthly sales trends and growth.
-   Add a dynamic analysis timestamp using `NOW()`.
-   Identify high-value customers using `INDEX()` and `MATCH()`.
-   Compare sales across regions and product categories.
-   Present key performance indicators (KPIs) through charts and a
    dashboard.

## ❓ 3. Problem Statements

Businesses collect sales data every day, but raw records alone do not
explain business performance. This project addresses the following
questions:

1.  **Customer performance:** Which 10 customers contribute the highest
    purchase value?
2.  **Discount decisions:** How does increasing or decreasing the
    discount affect total profit?
3.  **Sales and profit relationship:** Is there a measurable
    relationship between sales and profit?
4.  **Data distribution:** What are the average, median, minimum,
    maximum, and standard deviation of sales, profit, and discount?
5.  **Monthly performance:** Which months show sales growth or decline?
6.  **Customer value:** Who are the highest-value customers based on
    purchase amount?
7.  **Regional performance:** Which region generates the highest sales?
8.  **Product performance:** Which product category contributes the most
    sales?
9.  **Business reporting:** How can the main findings be communicated
    clearly in one dashboard?

## 🗂️ 4. Workbook Structure

  -----------------------------------------------------------------------
  Worksheet                           Purpose
  ----------------------------------- -----------------------------------
  **Project Instructions**            Original practical instructions and
                                      required tasks

  **Dashboard**                       KPI cards and visual
                                      sales-performance summary

  **Dataset**                         Source data containing 200 sales
                                      records

  **Customer Analysis**               Customer purchase totals, profit,
                                      order count, ranking, and
                                      high-value customer lookup

  **What-If Analysis**                Discount scenarios and their impact
                                      on profit

  **Regression & Stats**              Linear regression, descriptive
                                      statistics, and sales histogram

  **Pivot Summary**                   Region-wise and product-wise sales
                                      summary
  -----------------------------------------------------------------------

## 🧮 5. Excel Formulas & Techniques Used

### A. Conditional Formatting -- Top 10 Customers

Highlights the 10 highest total-purchase values so that the most
valuable customers can be identified quickly.

**Excel path:**
`Home → Conditional Formatting → Top/Bottom Rules → Top 10 Items`

### B. Profit Calculation

The workbook uses the profit values supplied in the source dataset. For
a simple transaction model where profit is calculated from sales after
discount and cost, the formula would be:

``` excel
=Sales-(Sales*Discount)-Cost
```

> Note: Use this formula only when the dataset contains the required
> cost value and the business defines profit this way.

### C. What-If Analysis -- Discount Impact

Tests alternative discount percentages and estimates the resulting total
profit.

**Scenario formula logic:**

``` excel
=SUMPRODUCT(ProfitRange-SalesRange*(ScenarioDiscount-DiscountRange))
```

This approach adjusts recorded profit by the difference between the
scenario discount and each transaction's original discount, assuming the
discount change affects profit rupee-for-rupee with sales.

**Excel path:**
`Data → What-If Analysis → Scenario Manager / Data Table`

### D. Linear Regression -- Profit vs Sales

Studies how profit changes in relation to sales.

``` excel
=SLOPE(ProfitRange,SalesRange)
=INTERCEPT(ProfitRange,SalesRange)
=RSQ(ProfitRange,SalesRange)
=CORREL(ProfitRange,SalesRange)
=FORECAST.LINEAR(1000,ProfitRange,SalesRange)
```

-   **Slope:** Estimated change in profit for a one-unit increase in
    sales.
-   **Intercept:** Estimated profit when sales are zero.
-   **R-squared:** Indicates how much of the variation in profit is
    explained by the linear model.
-   **Correlation:** Measures the direction and strength of the linear
    relationship.
-   **Forecast:** Estimates profit for a sales value of 1,000.

### E. Descriptive Statistics

The analysis includes:

``` excel
=COUNT(SalesRange)
=AVERAGE(SalesRange)
=MEDIAN(SalesRange)
=STDEV.S(SalesRange)
=MIN(SalesRange)
=MAX(SalesRange)
=SUM(SalesRange)
```

The same measures are calculated for profit and discount.

### F. Monthly Sales Growth & Arrows

Monthly sales are compared with the previous month. Up/down symbols or
custom number formatting can be used to make the direction of change
easier to interpret.

Example growth formula:

``` excel
=(CurrentMonthSales-PreviousMonthSales)/PreviousMonthSales
```

### G. Dynamic Timestamp

``` excel
=NOW()
```

This returns the current date and time and updates when Excel
recalculates the workbook.

### H. High-Value Customer Lookup

``` excel
=INDEX(CustomerNameRange,MATCH(MAX(TotalPurchaseRange),TotalPurchaseRange,0))
```

-   `MAX()` finds the highest purchase amount.
-   `MATCH()` finds its position.
-   `INDEX()` returns the corresponding customer name.

### I. Region & Product Sales Summary

The summary uses `SUMIFS()` to aggregate sales by region and product
category.

``` excel
=SUMIFS(SalesRange,RegionRange,SelectedRegion,ProductRange,SelectedProduct)
```

This workbook contains a formula-based Pivot-style summary. It is not a
native Excel PivotTable object.

## 📊 6. Dashboard & Visualizations

The dashboard brings important business indicators together, such as:

-   Total Sales
-   Total Profit
-   Number of Orders
-   Average Discount

Charts are used to communicate performance:

  Chart            Analytical Purpose
  ---------------- ----------------------------------------------------
  **Bar Chart**    Compare sales across regions or product categories
  **Line Chart**   Observe monthly sales trends
  **Pie Chart**    Show the contribution of categories or regions

## 🔍 7. Dataset Description

The source dataset contains 200 records and the following fields:

  Column                 Description
  ---------------------- ------------------------------------------
  `Customer_ID`          Unique customer identifier
  `Customer_Name`        Customer name
  `Region`               Sales region
  `Product_Category`     Product category
  `Sales`                Sales amount
  `Quantity`             Units purchased
  `Discount`             Discount rate
  `Order_Date`           Date of order
  `Profit`               Recorded profit
  `Analysis_Timestamp`   Dynamic timestamp generated with `NOW()`

## 📈 8. Initial Dataset Snapshot

The following figures are calculated from the source records and provide
a quick overview:

  KPI                                            Value
  ----------------------- ----------------------------
  Total Sales                               195,217.76
  Total Recorded Profit                      68,287.01
  Number of Orders                                 200
  Average Discount                              9.725%
  Unique Customers                                  27
  Regions                                            5
  Product Categories                                 5
  Date Range                10-Apr-2024 to 07-Apr-2025

**Initial observations:** - The West region has the highest total sales
in this dataset. - Books is the highest-selling product category by
sales value. - These are descriptive observations from the supplied
records, not predictions of future performance.

## 💡 9. Business Insights & Recommendations

-   **Customer focus:** Use the Top 10 analysis to identify customers
    who may benefit from loyalty offers.
-   **Discount control:** Compare profit under several discount
    scenarios before choosing a promotion.
-   **Regional planning:** Review the sales contribution of each region
    to guide inventory and marketing decisions.
-   **Product planning:** Compare category sales to support product
    assortment and stock planning.
-   **Performance monitoring:** Use monthly trends and dashboard KPIs to
    identify changes that need investigation.

## 🚀 10. How to Use This Project

1.  Download or clone this repository.
2.  Open `PR_2_Analyzer_Completed.xlsx` in Microsoft Excel.
3.  Start with the **Project Instructions** worksheet.
4.  Review the **Dataset** worksheet to understand the source columns.
5.  Explore **Customer Analysis** and **What-If Analysis**.
6.  Review **Regression & Stats** and **Pivot Summary**.
7.  Open **Dashboard** for the visual summary.
8.  Allow Excel to recalculate formulas so that dynamic values such as
    `NOW()` are refreshed.

## ⚠️ 11. Notes & Limitations

-   The workbook is built for learning and practical demonstration.
-   Formula results may recalculate when the workbook is opened in
    Microsoft Excel.
-   `NOW()` is volatile and changes when Excel recalculates.
-   The discount scenario calculation assumes that changing the discount
    changes profit by the corresponding sales amount; actual business
    margins may also depend on taxes, fees, variable costs, and pricing
    rules.
-   The region/product summary is formula-based rather than a native
    PivotTable.
-   Regression indicates association, not proof that sales changes cause
    profit changes.
-   The source data's recorded profit values are used for the analysis;
    they are not independently audited.

## 🛠️ 12. Skills Demonstrated

`Microsoft Excel` · `Data Cleaning & Review` · `Excel Formulas` ·
`Conditional Formatting` · `What-If Analysis` · `Linear Regression` ·
`Descriptive Statistics` · `Customer Analysis` · `SUMIFS` ·
`INDEX & MATCH` · `Data Visualization` · `Dashboard Design` ·
`Business Insights`

## 📁 13. Suggested GitHub Repository Layout

``` text
Sales-Performance-Analysis/
│
├── README.md
├── PR_2_Analyzer_Completed.xlsx
└── images/
    ├── dashboard.png
    ├── customer-analysis.png
    └── what-if-analysis.png
```

Add screenshots to the `images/` folder and embed them in this README
when you are ready to showcase the dashboard.

------------------------------------------------------------------------

::: {align="center"}
### ⭐ Thank you for visiting this project!

**Created by ABHISHEK CHOUDHARY**

*Learning Excel. Exploring Data. Discovering Insights.*
:::
