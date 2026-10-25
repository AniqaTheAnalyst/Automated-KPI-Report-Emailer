# Automated KPI Report Emailer

A Python automation that pulls loan KPIs from SQL Server, builds a multi-sheet Excel report, and emails it as an attachment, with no manual steps.

Built on the `bank_loan` table from my [BFSI Bank Loan Analysis](https://github.com/AniqaTheAnalyst) project (38,576 loan applications).

## What it does

1. Connects to SQL Server with `pyodbc` (Windows authentication).
2. Runs 8 KPI queries stored as separate `.sql` files in `queries/`.
3. Combines the results with `pandas` and writes a dated Excel report (`KPI_REPORT_YYYY-MM-DD.xlsx`) with two sheets:
   - **Summary**: all headline KPIs
   - **Good_vs_Bad**: loan count split by good (Fully Paid / Current) and bad (other statuses) loans
4. Emails the report through Gmail SMTP (SSL, port 465). Credentials are read from a `.env` file and never stored in the code.

## KPIs reported

| KPI | Source query |
|---|---|
| Total loan applications | `total_applications.sql` |
| Total funded amount | `total_funded.sql` |
| Total amount received | `total_received.sql` |
| Average interest rate (%) | `avg_interest.sql` |
| Average DTI (%) | `avg_dti.sql` |
| Good vs. bad loan split | `good_vs_bad.sql` |
| Bad loan funded amount | `bad_loan_funded.sql` |
| Bad loan amount received | `bad_loan_received.sql` |

## Tech stack

Python, SQL Server (T-SQL), pyodbc, pandas, openpyxl, smtplib, python-dotenv

## Setup

```bash
git clone <this-repo>
cd Automated-KPI-Report-Emailer
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

1. Install [ODBC Driver 18 for SQL Server](https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server) and load the `bank_loan` table into a SQL Server database.
2. Update the `SERVER` and `DATABASE` values in the connection string in `main.py`.
3. Copy `.env.example` to `.env` and fill in your details. For Gmail, use an [App Password](https://myaccount.google.com/apppasswords) (2-step verification required), not your normal password.

## Run

```bash
python main.py
```

The report is saved in the project folder and emailed to the recipient.

## Scheduling (Windows Task Scheduler)

Create a daily task that runs `venv\Scripts\python.exe` with the argument `main.py`, and set **Start in** to the project folder so the `queries/` folder and `.env` are found.

## Sample output

See `sample_output/` for an example report and email screenshot.

## Project structure

```
Automated-KPI-Report-Emailer/
├── main.py
├── queries/          # one .sql file per KPI
├── requirements.txt
├── .env.example      # template; real .env is git-ignored
└── .gitignore
```
