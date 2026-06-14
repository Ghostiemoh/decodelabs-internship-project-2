# Decodelabs Internship: Project 2 - E-Commerce Data Analytics & Executive Dashboard

This repository contains the dataset, dashboard implementation, and documentation for Project 2 of the DecodeLabs Data Analytics Internship. The project involves generating descriptive statistics, monthly/yearly trends, acquisition metrics, and coupon performance insights for an e-commerce platform.

## Project Overview

To drive revenue and solve operational inefficiencies, this project adds a dynamic, formula-driven **Executive Summary** dashboard to the transactional dataset. The dashboard consolidates raw sales records into high-level business intelligence metrics, allowing stakeholders to identify key growth areas and leakage points immediately.

### Project File Inventory

* **`Dataset for Data Analytics P2.xlsx`**: The transactional dataset containing the main data and the new, styled `Executive Summary` dashboard tab.
* **`Data analytics P2.pdf`**: The official project guidelines and instruction manual.

---

## Dashboard Implementation Details

The `Executive Summary` dashboard was inserted as the first worksheet tab. To ensure the metrics recalculate automatically when raw data changes, the dashboard relies entirely on Excel formulas rather than static values.

### 1. Core Performance Metrics Table
Calculates standard descriptive statistics across the numeric columns (`Quantity`, `UnitPrice`, `ItemsInCart`, `TotalPrice`):
* **Row Count**: `=COUNTA(Sheet1!E2:E1201)` (verifies dataset size)
* **Average (Mean)**: `=AVERAGE(Sheet1!E2:E1201)`
* **Median**: `=MEDIAN(Sheet1!E2:E1201)`
* **Minimum**: `=MIN(Sheet1!E2:E1201)`
* **Maximum**: `=MAX(Sheet1!E2:E1201)`
* **Standard Deviation**: `=STDEV(Sheet1!E2:E1201)`

### 2. Order Fulfillment Funnel
Tracks the operational health of the checkout and logistics pipeline:
* **Status Count**: `=COUNTIF(Sheet1!I2:I1201, "StatusName")`
* **Status Percentage**: `=StatusCount/COUNTA(Sheet1!A2:A1201)`

### 3. Product Performance Profile
Groups and aggregates transaction volumes and revenue across the 7 product catalog lines:
* **Order Count**: `=COUNTIF(Sheet1!D2:D1201, "Product_Name")`
* **Total Quantity**: `=SUMIF(Sheet1!D2:D1201, "Product_Name", Sheet1!E2:E1201)`
* **Total Revenue**: `=SUMIF(Sheet1!D2:D1201, "Product_Name", Sheet1!N2:N1201)`
* **Avg. Unit Price**: `=AVERAGEIF(Sheet1!D2:D1201, "Product_Name", Sheet1!F2:F1201)`

### 4. Coupon Code Return Correlation
Investigates how coupon usage (or lack thereof) impacts post-purchase customer satisfaction:
* **Return Count per Coupon**: `=COUNTIFS(Sheet1!L2:L1201, "Coupon_Name", Sheet1!I2:I1201, "Returned")`
* **Return Rate Percentage**: `=Return_Count/Total_Coupon_Orders`
* **No-Coupon Returns**: Uses `=SUM(Sheet1!N2:N1201)-SUM(Coupons)` and relative counts to isolate orders placed without a discount code.

---

## Key Strategic Observations

### 🚨 Operational Red Flag: Fulfillment Leakage
A critical finding is that **41.41%** of all generated orders are lost to either **Cancellations (20.83%)** or **Returns (20.58%)**. Completed, successful deliveries account for only **19.25%** of transactions. This indicates severe friction in product descriptions, shipping operations, or payment verification.

### 📈 Compounding Revenue Contraction
Comparing the first half (H1) of each year to eliminate seasonal bias:
* **H1 2023 Revenue**: **$286,501.52**
* **H1 2024 Revenue**: **$257,059.34** *(10.28% YoY decline)*
* **H1 2025 Revenue**: **$231,882.85** *(9.79% YoY decline)*

This shows a steady, compounding contraction in sales of approximately 10% year-over-year.

### 🎟️ Coupon Anchoring Effect
Orders placed without a coupon code suffer from the highest return rate on the platform (**24.60%**). Conversely, using the **`SAVE10`** coupon correlates with the lowest return rate (**16.43%**). Providing small checkout incentives could serve as a valuable tool for lowering buyer remorse and returns.
