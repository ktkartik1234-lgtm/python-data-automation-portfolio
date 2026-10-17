# Python & Advanced Excel Data Automation Portfolio

A curated collection of production-grade Python automation scripts, programmatic Excel workbook generators, and automated quality assurance (QA) pipelines. 

Designed for **Data Operations**, **Business Process Automation**, and **Data Integrity Auditing**.

---

## 🛠 Core Technical Competencies

- **Languages & Libraries:** Python 3 (OpenPyXL, Pandas, PyMuPDF/fitz, Regex, Datetime), SQL.
- **Spreadsheet Engineering:** Multi-sheet programmatic workbook generation, dynamic formula linking, conditional formatting, data validation dropdowns, and automated audit control boards.
- **Data Integrity & QA:** Automated text parsing, character-length constraint enforcement, substring verification against source documents, duplicate detection, and mathematical tie-out validation.
- **Process Automation:** Transforming unstructured text and raw dumps into publication-ready datasets with zero manual intervention.

---

## 📂 Featured Case Studies

### 1. [LMS Question Bank & Data Auditing Pipeline](./01-lms-data-pipeline)
**Problem:** Educational publishers and LMS platforms require thousands of curriculum questions, study notes, and options compliant with strict database character limits, balanced answer keys, and verified source grounding. Manual QA is slow and prone to formatting errors.

**Solution:**
- Built an automated Python validation engine (`lms_dataset_verification_engine.py`) that audits 2,000 row-aligned data entries against a source reference PDF using PyMuPDF.
- Enforced strict character length limits across 7 database fields (Questions $\le$ 160 chars, Options A–D $\le$ 80 chars, Notes $\le$ 160 chars).
- Automated answer key distribution auditing (exactly 25.0% A, B, C, D balance) and regex-based punctuation cleaning.
- **Key Metric:** Audited and verified **4,000 pedagogical items** across 9 chapters with **100% compliance** and zero hallucinations.

---

### 2. [Programmatic Financial Reconciliation & Audit Model](./02-financial-reconciliation-generator)
**Problem:** Small businesses and finance managers struggle with broken manual spreadsheets where P&L, Balance Sheet, and Cash Flow figures fail to balance due to orphaned or mistyped entries.

**Solution:**
- Developed a Python script (`generate_reconciliation_model.py`) using `openpyxl` that generates an end-to-end, styled 3-statement financial workbook (`.xlsx`).
- Implemented an **Automated Audit Control Board** utilizing live `=IF(ROUND(Variance, 2)=0, "PASS", "FAIL")` formulas that verify mathematical tie-outs to the cent.
- Built a structured General Ledger with strict double-entry balancing (`Total Debits - Total Credits = $0.00`) and a Bank Reconciliation clearing schedule.
- **Key Metric:** Zero variance (\$0.00 discrepancy) across Balance Sheet assets/liabilities, cash flow closing balances, and bank statements.

---

### 3. [Dynamic Operations & Content Management Dashboard](./03-excel-operations-dashboard)
**Problem:** Media and operations teams managing large digital catalogs need unified dashboards with status trackers, KPI cards, and strict data validation to prevent data entry corruption.

**Solution:**
- Engineered an automated workbook generator (`create_operations_dashboard.py`) applying executive slate/crimson styling, auto-adjusting column dimensions, and native Excel data validation rules.
- Structured separate operational views for Track Management, Channel Distribution, and Executive Summary metrics.
- **Key Metric:** Replaced unstructured manual logging with a standardized, formula-driven management workbook.

---

## 🚀 Quick Start / Running the Models

### Prerequisites
```bash
pip install openpyxl pymupdf pandas
```

### Run Financial Model Generator
```bash
cd 02-financial-reconciliation-generator
python generate_reconciliation_model.py
```
*Outputs `Financial_Reconciliation_and_Reporting_Model.xlsx` with active audit check formulas.*

### Run Operations Dashboard Generator
```bash
cd 03-excel-operations-dashboard
python create_operations_dashboard.py
```

---

## 📬 Contact & Availability
- **Role Targets:** Data Operations Specialist, Junior Data Analyst, Python Automation Engineer.
- **Freelance Engagements:** Python Scripting, Automated Data Cleaning, Excel Programmatic Engineering.
- **GitHub:** [@ktkartik1234-lgtm](https://github.com/ktkartik1234-lgtm)
