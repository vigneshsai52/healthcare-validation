# 🔍 Healthcare Data Validation Dashboard

SQL-based data validation system that detects data quality issues in healthcare billing records.

## 📌 Project Overview

This project loads 55,500+ healthcare patient records into a SQLite database, runs SQL queries to validate data quality, and generates a color-coded Excel validation report highlighting issues found.

## 🛠️ Tech Stack

- Python
- Pandas (Data Processing)
- SQLite (SQL Database & Queries)
- OpenPyXL (Excel Report Generation)

## 🔍 Validation Checks Performed

- **Duplicate Records** – Detects duplicate patient names
- **Invalid Billing** – Finds zero or negative billing amounts
- **Age Outliers** – Flags ages below 0 or above 120
- **Invalid Gender** – Catches values outside Male/Female
- **Condition Summary** – SQL aggregation by medical condition
- **Test Results Breakdown** – Normal/Abnormal/Inconclusive split

## 📊 Validation Results Found

| Check | Result |
|---|---|
| Total Records | 55,500 |
| Duplicate Names | 5,507 🟡 |
| Invalid Billing | 108 🟡 |
| Age Outliers | 0 🟢 |
| Invalid Gender | 0 🟢 |

## ▶️ How to Run

```bash
pip install -r requirements.txt
cd src
python validation_report.py
```

## 📁 Project Structure

```
healthcare-validation/
├── data/
│   ├── healthcare_dataset.csv    → Input CSV dataset
│   └── healthcare.db             → SQLite database
├── output/
│   └── validation_report.xlsx    → Color-coded validation report
├── src/
│   └── validation_report.py      → Main validation script
├── requirements.txt
└── README.md
```
