
import pyodbc
import pandas as pd
from datetime import date
import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 18 for SQL Server};"
    "SERVER=.\\SQLEXPRESS;"
    "DATABASE=BL DB;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

queries = {
    "Total Applications": "total_applications.sql",
    "Total Funded": "total_funded.sql",
    "Total Received": "total_received.sql",
    "Avg Interest": "avg_interest.sql",
    "Avg DTI": "avg_dti.sql",
    "Good vs Bad": "good_vs_bad.sql",
    "Bad Loan Received": "bad_loan_received.sql",
    "Bad Loan Funded": "bad_loan_funded.sql",
}


# Preparing the data KPIS IN EXCEL

results = {}
for name, file in queries.items():
    with open(f"queries/{file}") as f:
        results[name] = pd.read_sql(f.read(), conn)

single_KPI = [
    "Total Applications",
    "Total Funded",
    "Total Received",
    "Avg Interest",
    "Avg DTI",
    "Good vs Bad",
    "Bad Loan Received",
    "Bad Loan Funded",
]

summery = pd.DataFrame({
    "KPIS": single_KPI,
    "values" :[results[k].iloc[0,0] for k in single_KPI]

})


filename = f"KPI_REPORT_{date.today()}.xlsx"

with pd.ExcelWriter(filename , engine = "openpyxl") as writer:
    summery.to_excel(writer , sheet_name="summery" , index=False)
    results["Good vs Bad"].to_excel(writer , sheet_name="Good_vs_Bad" , index=False)


    print("Saved", filename)


# Entering credientials

load_dotenv()
SENDER =  os.getenv("EMAIL_SENDER")
PASSWORD = os.getenv("EMAIL_PASSWORD")
RECIPIENT = os.getenv("EMAIL_RECIPIENT")

# Building the email body
msg = EmailMessage()
msg['subject'] = f"KPI Report - {date.today()}"
msg['from'] = SENDER
msg['to'] = RECIPIENT
msg.set_content(f"Dear Team,\n\nPlease find attached the KPI report for {date.today()}.\n\nBest regards,\nAutomated KPI Reporter")

with open(filename, 'rb') as f:
    msg.add_attachment(
        f.read(),
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=filename,
    )


# Sending the email
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(SENDER, PASSWORD)
    server.send_message(msg)

print("Email sent")