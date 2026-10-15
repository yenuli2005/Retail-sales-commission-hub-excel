# 📊 Retail Sales Performance & Commission Intelligence Hub

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-Retail_Analytics-blue?style=for-the-badge)
![Corporate Finance](https://img.shields.io/badge/Corporate_Finance-Commission_Modeling-00758F?style=for-the-badge)
![Project](https://img.shields.io/badge/Project-Milestone_2-orange?style=for-the-badge)

An enterprise-standard Excel workbook built to model retail sales transactions, automated lookup-driven commission payouts, multi-criteria data aggregations, and executive dashboard analytics.

---

## 📷 Dashboard Overview & Preview

<img width="1366" height="768" alt="Image" src="https://github.com/user-attachments/assets/a07f3345-fd3a-4e01-b9cd-c88d9ad3d43e" />

---

## 🎯 Project Purpose & Key Skills Tested

This project served as my practical sandbox to apply advanced data modeling and dynamic reporting in Excel. The goal was to build an end-to-end automated system that transforms raw transactional data into executive business intelligence.

* **Dynamic Approximate Lookups:** Automating tier-based sales commission rates using `XLOOKUP`.
* **Multi-Criteria Data Aggregation:** Cross-tabulating sales, revenue, and demographics using `SUMIFS` and `COUNTIFS`.
* **Executive Summary Modeling:** Evaluating overall revenue, average order value, peak payouts, and total expenses using `AVERAGE`, `MAX`, and `SUM`.
* **Interactive Reporting:** Designing dynamic KPI cards and corporate charts for executive presentation.

---

## 📁 Model Architecture

```text
Retail_Sales_Commission_Hub.xlsx
│
├── 📊 Sheet 1: Raw Data (200 Cleaned Transaction Records)
├── ⚙️ Sheet 2: Commission Engine (Tiered Rate Table & Calculation Rows)
├── 📋 Sheet 3: Summary Tables (Category Performance & Demographic Matrices)
└── 📈 Sheet 4: Dashboard (KPI Summary Blocks & Visual Charts)
```

---

## 🛠️ Formulas & Mathematical Mechanics

### Dynamic Tiered Commission Lookup (`XLOOKUP`)

Commission rates are dynamically assigned based on transaction total sales across tiered thresholds:

**Commission Rate = XLOOKUP(D2,H$2:H$5,I$2:I$5,0,-1)**

**Commission payout = Total Sales*Commission Rate**

### Base Model Summary
* **Scope:** 200 Transaction Records (Full Year 2023)
* **Gross Revenue Analyzed:** $\$103,635.00$
* **Total Commission Expense:** $\$7,248.38$

<img width="1366" height="768" alt="Image" src="https://github.com/user-attachments/assets/e08df891-f695-451b-8b6f-c4d60409f646" />

---

## 💡 Key Takeaways & Lessons Learned

1. **Approximate Matching (`-1`):** Applied `XLOOKUP` in match mode `-1` to eliminate nested `IF` formulas for tiered rates.
2. **Multi-Condition Aggregation:** Constructed clean cross-tabulations using `SUMIFS` and `COUNTIFS` without altering raw data.
3. **Structured Design:** Enforced a clean separation between raw data, calculation engines, summary layers, and dashboards.

---

## ⚙️ How to Run & Test

1. **Download:** Clone or download `Retail_Sales_Commission_Hub.xlsx`.
2. **Launch:** Open using Microsoft Excel 2016+, Office 365, or Google Sheets.
3. **Interact:** Inspect formulas across the `Commission Engine`, `Summary Tables`, and `Dashboard` tabs.

---

## 🔒 Confidentiality & Data Disclaimer

> **Note on Data:** The dataset was adapted from open analytics repositories (Kaggle) and refined for educational and demonstration purposes.

---

## 👩‍💻 Author

**Yenuli Withanage**  
*B.Sc. in Finance Undergraduate | University of Sri Jayewardenepura (USJ)*  
*AAT SL Level 3 Candidate*  

[LinkedIn Profile](https://www.linkedin.com) • [GitHub Repository](https://github.com/yenuli2005)
