📊 Accounting Automation Engine (Python)
🔹 Overview

This project is a Python-based accounting automation engine designed to generate key financial statements and reconciliation reports directly from raw accounting data. It automates repetitive finance and audit tasks such as trial balance preparation, income statement and balance sheet generation, and data reconciliation — reducing manual effort and improving accuracy.

This project was built by a finance guy learning Python to solve real-world problems encountered during audit internships.

🔹 Key Features

✅ Generate Trial Balance from raw journal entries

✅ Create Income Statement and Balance Sheet from mapped trial balance

✅ Perform Reconciliation between two datasets (e.g., bank vs cash book)

✅ Export all outputs to Excel with structured sheets

✅ Designed for real-world audit and finance workflows

🔹 Technologies Used

Python

Pandas

Excel (via xlsxwriter)

🔹 How to Use
1️⃣ Import and Create Object
from learn_class import Accounting
import pandas as pd

acc = Accounting(Journaldf=journal_df, Mappedtb=mapped_tb, Rec1=rec_df1, Rec2=rec_df2)

2️⃣ Generate Trial Balance
acc.tb(pindex=0, debindex=1, credindex=2)

3️⃣ Generate Income Statement
acc.income_statement()

4️⃣ Generate Balance Sheet
acc.balance_sheet()

5️⃣ Perform Reconciliation
acc.reconcilation(uniqid1=0, uniqueid2=0, amtclm1=1, amtclm2=1)

6️⃣ Export to Excel
acc.export_to_excel("financial_reports.xlsx", rec=True)


This will export:

Trial Balance

Income Statement

Balance Sheet

Reconciliation (matching & non-matching items)

🔹 Data Requirements
📌 Mapped Trial Balance must contain these columns:

MAPPING1 → SOPL or SOFP

MAPPING2 → INCOME, EXPENSE, CURRENT ASSET, NON CURRENT ASSET, CURRENT LIABILITY, NON CURRENT LIABILITY

MAPPING3 → User-defined classification

ADJUSTED CLOSING → Final balances

Spelling and capitalization must match exactly.

🔹 Why This Project Matters

This project demonstrates:

Financial accounting knowledge

Audit and reconciliation logic

Automation and analytical thinking

Python and Pandas proficiency

Real-world problem-solving

Built not as a tutorial exercise, but as a tool to improve efficiency during audit and accounting work.

🔹 Status

🛠️ Still under active development — additional features and validations will be added over time.

👤 Author

Afras Abdulla
CMA (Certified Management Accountant) | BCOM Graduate
GitHub: Afras-Abdulla