# Data Cleaning & Preparation — DecodeLabs Industrial Training Kit (Project 1)

A data integrity audit and cleaning pipeline built on a raw, 1,200-row e-commerce order dataset. The goal of this project wasn't visualization or modeling — it was proving that messy, real-world data can be turned into a verified, production-ready source of truth using both code-based and formula-based approaches.

## Problem Statement

The raw dataset (`data/raw/raw_dataset.xlsx`) contained three classes of data quality issues common in real operational exports:

| Issue | Detail |
|---|---|
| Missing values | 309 of 1,200 rows (25.8%) had a blank `CouponCode` field |
| Inconsistent numeric precision | 114 `UnitPrice` and 260 `TotalPrice` values had inconsistent decimal places (0, 1, or 2 decimals) due to floating-point export artifacts |
| Formatting inconsistency | Free-text fields (`Product`, `PaymentMethod`, `OrderStatus`, `ShippingAddress`, `ReferralSource`) had inconsistent casing and stray whitespace |

Duplicate `OrderID` and `TrackingNumber` values, and full duplicate rows, were also audited as part of the integrity check (the dataset passed this check with zero duplicates found).

## Two Approaches, Same Result

This repo intentionally solves the same problem two ways, to demonstrate both tracks called out in the training kit:

### 1. Python / Pandas (`src/clean_data.py`)
A reproducible script: load → clean → verify → write formatted `.xlsx` output with an auto-generated change log. Best suited for large or recurring datasets where the same cleaning logic needs to run repeatedly without manual rework.

### 2. Excel-native formulas (`src/build_excel_formulas.py` → `excel/Cleaned_Dataset_Excel_Formulas.xlsx`)
Every cleaned value in this workbook is a live Excel formula (`TRIM`, `PROPER`, `ROUND`, `COUNTIF`, `ISNUMBER`) referencing an untouched `Raw_Data` sheet — nothing is pre-computed. Opening it in Excel and pressing F9 recalculates everything from scratch. This mirrors a Power Query–style workflow for analysts who want a fully transparent, auditable spreadsheet rather than a script.

## Cleaning Methodology

**Phase 1 — Strategic Imputation.** Missing `CouponCode` values were not statistically imputed (mean/median/mode), because their absence is meaningful, not random — it reflects an order placed without a discount. They were filled with an explicit `"No Coupon"` label instead.

**Phase 2 — Integrity Audit.** Checked for duplicate `OrderID`, duplicate `TrackingNumber`, and full duplicate rows using the SQL-equivalent logic `GROUP BY id HAVING COUNT(*) > 1`, implemented as `df.duplicated()` in pandas and `COUNTIF` in Excel.

**Phase 3 — Standardization.** Dates verified as true date types and formatted to ISO 8601 (`YYYY-MM-DD`). Text fields trimmed and title-cased. Numeric fields rounded to a consistent 2 decimal places. ID and coupon code fields were deliberately excluded from re-casing to avoid corrupting fixed codes (e.g. `SAVE10` staying `SAVE10`, not becoming `Save10`).

**Phase 4 — Documentation.** Every change is logged in a `Change_Log` sheet/section with a Change ID, description, measurable impact, and status — so the cleaning work is auditable, not just trusted.

## Verification Gate

Both outputs pass the same hard requirement: **zero duplicate IDs and zero invalid date formats.**

```
Duplicate OrderIDs:         0
Duplicate TrackingNumbers:  0
Invalid / non-date cells:   0
Remaining missing values:   0
VERDICT: PASS — 0% Error Rate
```

## Repository Structure

```
.
├── data/
│   ├── raw/
│   │   └── raw_dataset.xlsx          # Original, unmodified source data
│   └── cleaned/
│       └── Cleaned_Dataset.xlsx      # Python/pandas output + change log sheet
├── excel/
│   └── Cleaned_Dataset_Excel_Formulas.xlsx   # Formula-driven Excel workbook
├── src/
│   ├── clean_data.py                 # Pandas cleaning pipeline
│   └── build_excel_formulas.py       # Generates the formula-based workbook
├── requirements.txt
└── README.md
```

## How to Run

### Python pipeline
```bash
pip install -r requirements.txt
python src/clean_data.py
```
This reads `data/raw/raw_dataset.xlsx` (place it alongside the script or update the `SRC` path) and writes a cleaned workbook with a `Cleaned_Data` sheet and a `Change_Log` sheet.

### Excel formula workbook
```bash
pip install -r requirements.txt
python src/build_excel_formulas.py
```
Generates a workbook with `Raw_Data`, `Cleaned_Data`, `Verification_Gate`, and `Change_Log` sheets, where every cell is a live formula. Open it directly in Microsoft Excel or LibreOffice Calc.

## Tools Used

| Tool | Role |
|---|---|
| **Python (pandas)** | Data loading, transformation logic, automated change-log generation |
| **openpyxl** | Cell-level Excel formatting (fonts, number formats, conditional fills, frozen headers) |
| **Excel formulas** (`TRIM`, `PROPER`, `ROUND`, `COUNTIF`, `ISNUMBER`, `SUMPRODUCT`) | Native, transparent, recalculable cleaning logic for analysts working directly in spreadsheets |

## Key Takeaway

Cleaning is not about deleting anything inconvenient — it's about making deliberate, defensible, documented decisions about every gap, duplicate, and inconsistency in a dataset before it's trusted for analysis.

---
*Part of the DecodeLabs Industrial Training Kit — Data Analytics, Project 1 (Batch 2026).*
