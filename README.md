# Excel-Sales-and-Promotion-Analysis 📊
Advanced Excel analysis using Power Query, Power Pivot, DAX, and Macros to evaluate sales performance, coupon effectiveness, and referral sources.

## 📌 Project Summary
This project analyzes an e-commerce order dataset using **Microsoft Excel advanced features** to evaluate sales performance, promotion effectiveness, and customer acquisition channels.

The goal is to demonstrate real-world **business analytics workflows in Excel**, including data cleaning, modeling, and analysis using **Power Query, Power Pivot and DAX measures** .

The project focuses on:
- Measuring coupon and promotion effectiveness
- Analyzing referral sources and revenue contribution
- Tracking order fulfillment stages
- Building reusable analytical models in Excel

---

## 📥 Dataset Download (Raw Excel File)

The raw dataset used in this project was downloaded from ExcelX, a website that provides free practice datasets for data analysis and Excel projects. It is named "2. Online Store Orders Dataset" in the link:
https://excelx.com/practice-data/sales-retail/

## 📂 Dataset
The dataset contains transactional order data with the following columns:

`OrderID, Date, CustomerID, Product, Quantity, UnitPrice, ShippingAddress, PaymentMethod, OrderStatus, TrackingNumber, ItemsInCart, CouponCode, ReferralSource, TotalPrice`


---

## 🧰 Tools & Techniques Used
- **Power Query**
  - Load Excel data into the Data Model
  - Clean and normalize text fields
  - Create helper columns (e.g., Coupon Used: Yes/No)
- **Power Pivot**
  - Build the Data Model
  - Create relationships and calculated measures
- **DAX (Data Analysis Expressions)**
  - Coupon usage analysis
  - Revenue and order KPIs
  - Referral source performance metrics
- **PivotTables & PivotCharts**
  - Interactive analysis and filtering

---

## 📊 Key Analyses & Metrics

### ✅ Promotion Effectiveness
- Orders with coupons used
- Revenue generated from coupon orders
- Average order value of coupon vs non-coupon orders
- Coupon usage percentage

### ✅ Referral Source Analysis
- Orders per referral source
- Revenue contribution by referral channel <!-- Not implement yet -->
- Average order value per referral source <!-- Not implement yet -->
- Coupon usage by referral channel <!-- Not implement yet -->

### ✅ Order Fulfillment Tracking
- Order volume by fulfillment status (Completed, Pending, Cancelled, Shipped, Returned)
- Orders' fulfillment rate 

---
## 📊 Dashboard Overview

All key metrics and visualizations produced in this project are consolidated into a single worksheet named **Dashboard**.

The **Dashboard** sheet includes:
- Promotion & coupon effectiveness
- Referral source performance
- Order fulfillment tracking
- Interactive PivotTables and PivotCharts

This sheet serves as the main analytical interface of the workbook, allowing users to explore insights without modifying the underlying data.
---
## 🧾 How to Open the Excel Workbook

1. Download the repository or clone it locally.
2. Open the file
3. When prompted by Excel:
- Click **Enable Editing**
- Click **Enable Content** (to allow macros)
4. Go to **Data → Refresh All** to reload the data model.
5. Use Power Pivot and PivotTables to explore the analysis.
