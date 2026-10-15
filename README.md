# Automated Income and Expense Dashboard Using Python and Excel

An end-to-end financial data engineering and business intelligence solution developed for **MEATOPHUM Microfinance** (Phnom Penh, Cambodia) as part of the 4th-Year Engineering Internship Thesis in Data Science at the **Institute of Technology of Cambodia (ITC)**.

---

## 📌 Project Overview

Prior to this project, monthly income and expense reports across MEATOPHUM Microfinance's **19 branches** were compiled manually by retyping figures from raw accounting exports (`IS-All`) into a static summary sheet (`SUM-IS`). This manual workflow was time-consuming, prone to transcription errors, and vulnerable to silent data duplication from inconsistent source formatting.

This project replaces the manual process with a **self-refreshing Microsoft Excel BI Dashboard** powered by a **Python (`pandas`) preprocessing pipeline**, **Power Query (M)** automated folder consolidation, a **Power Pivot Star Schema Data Model**, dynamic **DAX measures**, and **VBA one-click automation**.

---

## 🏗️ System Architecture & Technical Pipeline

1. **Historical Data Cleaning & Validation (Python / Jupyter Notebook):**
   * **Script:** `Transform_to_raw_data.ipynb`
   * Reads the multi-block raw accounting export (`20260531-Income and Expensed-Jan to July 26.xlsx`, sheet `IS-All`) with `header=None`.
   * Normalizes ordinal date banners and applies a custom **Date-Proximity Safety Check** (`0 <= h - d <= 4`) to automatically detect and skip stray, undated utility note tables that previously caused duplicate expense records.
   * Unpivots (`pd.melt`) the 19 wide branch columns into a normalized 7-column flat schema (`Date`, `GL Code`, `Description`, `Branch`, `Amount`, `Category`, `Signed Amount`) and exports the baseline master file (`IS_Master_All.xlsx`).

2. **Automated Monthly Consolidation & Categorization (Power Query / M Language):**
   * Monitors the `Dashboard_Data/` staging folder using a `Folder.Files` connector.
   * Automatically appends the historical master file (`IS_Master_All.xlsx`) with new monthly flat-format exports (`August_2026_NewData_FlatFormat.xlsx`, `September_2026_NewData_FlatFormat.xlsx`).
   * Implements sequential `Text.Contains` conditional rules in `Dim_Account` to group **60+ raw GL Code descriptions** into **12 standardized Expense Groups** (`Staff Cost`, `Loan Loss Provision`, `Travel & Vehicle`, `Rent`, `Depreciation & Amortisation`, `Professional Service`, `Office & Admin`, `Utility`, `Tax & Statutory`, `Repair & Maintenance`, `Interest Expense`, and `Other`).

3. **Dimensional Data Modeling (Power Pivot Star Schema):**
   * **`Fact_Table`:** Central transactional table (3,664 consolidated records across Jan–Sep 2026).
   * **`Dim_Date`:** Derived calendar attributes (`Date`, `Year`, `Month_Num`, `Month_Name`, `Quarter`) sorted chronologically.
   * **`Dim_Branch`:** Unique list of the 19 operating branches (`MPT`, `CHB`, `SNK`, `STR`, `KPC`, `SNL`, `PTR`, `STG`, `RTK`, `BKV`, `BLV`, `KHR 168`, `ASL`, `999 Cash`, `KTL`, `SLG`, `SST`, `ASY168`, `HO`).
   * **`Dim_Account`:** Unique chart of accounts (`GL Code`, `Description`, `Category`, `Expense Group`).

4. **Dynamic Financial KPIs (DAX Measures):**
   * `Total Income` = `CALCULATE(SUM(Fact_Table[Amount]), Dim_Account[Category] = "Income")`
   * `Total Expense` = `CALCULATE(SUM(Fact_Table[Amount]), Dim_Account[Category] = "Expense")`
   * `Net Income` = `[Total Income] - [Total Expense]`
   * `Profit Margin` = `DIVIDE([Net Income], [Total Income])`
   * `Income Growth %` & `Expense Growth %` (Month-over-Month time-intelligence calculations using `DATEADD`).

5. **Interactive Reporting & One-Click Automation (Excel & VBA):**
   * **`Analysis` Sheet:** Backend staging PivotTables feeding the visual layer.
   * **`Dashboard` Sheet:** Executive interface featuring 4 KPI Cards, 4 interactive PivotCharts, multi-sheet Slicers (`Branch`, `Month_Name`, `Expense Group`), and a `Date` Timeline.
   * **`SUM-Auto` (`Summary`) Sheet:** Automated branch-by-branch financial statement replacing the legacy manual `SUM-IS` report.
   * **VBA Macro (`RefreshDashboard`):** Executes `ThisWorkbook.RefreshAll` via a clickable header button, updating all queries, data model tables, and charts in under 5 seconds.

---

## 📂 Repository File Structure

```text
├── Dashboard_Data/
│   ├── IS_Master_All.xlsx                        # Cleaned historical master data (Jan–Jul 2026)
│   ├── August_2026_NewData_FlatFormat.xlsx       # Monthly flat-format export (August 2026)
│   └── September_2026_NewData_FlatFormat.xlsx    # Monthly flat-format export (September 2026)
├── Transform_to_raw_data.ipynb                   # Python data cleaning & validation pipeline
├── 20260531-Income and Expensed-Jan to July 26.xlsm  # Final automated Excel dashboard (.xlsm)
└── README.md                                     # Project documentation
```

---

## 🚀 How to Use the Dashboard

1. **Add New Monthly Data:** Place the new month's flat-format Excel export into the `Dashboard_Data/` directory.
2. **Update Folder Path (First-Time Setup Only):** In Excel, open `Data -> Queries & Connections -> Power Query Editor`, and update the `Dashboard_Data` folder path parameter to match your local directory path.
3. **One-Click Refresh:** Open `20260531-Income and Expensed-Jan to July 26.xlsm`, enable macros, and click the **Refresh** button in the top-right corner of the `Dashboard` sheet.
4. **Explore Insights:** Use the interactive `Branch`, `Month_Name`, and `Expense Group` slicers or click the `DASHBOARD`, `SUMMARY`, and `Analysis` navigation buttons to switch views.

---

## 👤 Author & Academic Information

* **Author:** Nov Panhavath
* **Institution:** Institute of Technology of Cambodia (ITC) — Department of Applied Mathematics and Statistics (Data Science Major)
* **Academic Advisor:** Mr. Dorn Dawin
* **Company Supervisor:** Ms. Wong Somaly (Head of Finance, MEATOPHUM Microfinance)
* **Academic Year:** 2025–2026
