import pandas as pd
import smtplib
import sys 
import os
import getpass
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from sqlalchemy import create_engine, Column, String, Text, DateTime, Integer
# from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import declarative_base
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus
from datetime import datetime

# Define the SQLAlchemy base
Base = declarative_base()

# Define the EmailHistory model
class EmailHistory(Base):
    __tablename__ = 'EmailHistory'
    
    Id = Column(Integer, primary_key=True, autoincrement=True)
    Name = Column(String(255))
    Email = Column(String(255))
    Cc_Email = Column(String(255))
    Subject = Column(String(255))
    Body = Column(Text)
    SenderName = Column(String(255))
    SentDate = Column(DateTime)

# Function to send email
def send_email(to_email, cc_email, subject, body, username, password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = to_email
    if cc_email:
        msg['Cc'] = cc_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.send_message(msg)

# Read email addresses and messages from Excel file
def read_emails_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Save email history to SQL Server
def save_email_history(session, Name, to_email, Cc_Email, subject, body, SenderName):
    email_record = EmailHistory(Name=Name, Email=to_email, Cc_Email=Cc_Email, Subject=subject, Body=body, SenderName=SenderName, SentDate=datetime.now())
    session.add(email_record)
    session.commit()

def load_dotenv_file():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dotenv_path = os.path.join(script_dir, ".env")
    if not os.path.exists(dotenv_path):
        return

    with open(dotenv_path, "r", encoding="utf-8") as file:
        for raw_line in file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if len(value) >= 2 and ((value[0] == value[-1] == '"') or (value[0] == value[-1] == "'")):
                value = value[1:-1]

            if key and key not in os.environ:
                os.environ[key] = value

def load_email_password():
    env_password = os.getenv("EMAIL_PASSWORD")
    if env_password:
        return env_password.strip()

    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        send_txt_path = os.path.join(script_dir, "Send.txt")
        with open(send_txt_path, 'r', encoding="utf-8") as file:
            file_password = file.read().strip()
            if file_password:
                return file_password
    except FileNotFoundError:
        pass

    return getpass.getpass("Enter your email password / app password: ").strip()

def create_sql_server_engine():
    explicit = os.getenv("SQLSERVER_SQLALCHEMY_URL")
    if explicit:
        return create_engine(explicit, use_setinputsizes=False)

    host = os.getenv("SQLSERVER_HOST")
    user = os.getenv("SQLSERVER_USER")
    password = os.getenv("SQLSERVER_PASSWORD")
    database = os.getenv("SQLSERVER_DB")
    driver = os.getenv("SQLSERVER_DRIVER", "SQL Server")

    missing = [k for k, v in {
        "SQLSERVER_HOST": host,
        "SQLSERVER_USER": user,
        "SQLSERVER_PASSWORD": password,
        "SQLSERVER_DB": database,
    }.items() if not v]
    if missing:
        raise RuntimeError(
            "Missing SQL Server configuration in environment variables: "
            + ", ".join(missing)
            + ". Set SQLSERVER_SQLALCHEMY_URL instead if you prefer a single connection string."
        )

    odbc_connect = (
        f"DRIVER={{{driver}}};"
        f"SERVER={host};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )

    return create_engine(
        "mssql+pyodbc:///?odbc_connect=" + quote_plus(odbc_connect),
        use_setinputsizes=False
    )

def main():
    try:
        load_dotenv_file()
        print("\n--------------------------------------------------")
        print("Please Login your Email Account [ Google Mail ]")
        print("--------------------------------------------------")
        username = input("Enter your email address: ")
        sendname = input("Please enter your name: ")

        password = load_email_password()
        
        print("\n+++++++++++++++++++++++++++++++++")
        print("Please Select the Excel File... ")
        print("+++++++++++++++++++++++++++++++++\n")
        root = Tk()
        root.withdraw() # Hide the root window
        root.attributes('-topmost', True)  # Ensure the dialog is on top
        file_path = askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
        root.destroy()  # Close the root window after selection
        
        if not file_path:
            print("No file selected. Exiting.")
            return

        df = read_emails_from_excel(file_path)

        # Group by NAME, EMAIL, and CC
        print("Reading Data in Excel")
        df = df.astype(str).replace('nan', '', regex=True)
        df.fillna('',inplace=True)
        df['CC'] = df['CC'].fillna('')
        
        grouped = df.groupby(['NAME', 'EMAIL', 'CC'])

        # SQL Server connection details
        print("Checking Server Connection")
        engine = create_sql_server_engine()
        Base.metadata.create_all(engine)
        
        Session = sessionmaker(bind=engine)
        session = Session()
        
        print("Preparing the email sending")
        df.columns = df.columns.str.upper()
        df['DATE ONLINE'] = pd.to_datetime(df.get('DATE ONLINE', datetime.now()), errors='coerce')
        df['DATE ONLINE'].fillna(datetime.now(), inplace=True)       
        df['MONTH_DAY_YEAR'] = df['DATE ONLINE'].dt.strftime('%B %d, %Y')  # For full month name
        df['MONTH_YEAR'] = df['DATE ONLINE'].dt.strftime('%B %Y')
        df['DATE ONLINE'] = df['DATE ONLINE'].dt.strftime('%m-%d-%Y')

        bmsg10a = (
                "<b>REMINDERS for Cash Advances:</b><br>"
                "1. Please make all receipts (manual/tape) <b>payable to Innogen Pharmaceuticals Inc.</b> For tape receipts, the <b>name of the company must be printed</b> through POS Machine, <b>not in handwriting form</b>. Make sure that the TIN & Address are readable. <b>(Nonreadable receipts are not acceptable)</b><br><br>"
                "2. All receipts must be submitted <b>within the month based on date of purchase / activity</b>, otherwise, <b>request will not be accepted/forfeited.</b><br><br>"
                "3. Please attach a <b>Post Activity Report (PAR)</b> with complete signature/approval on your liquidation, and attach an attendance sheet & pictures. <b>(for PPE and RBAF only)</b><br><br>"
                "4. <b style='color: red;'>Liquidation for Cash Advances should be made within seven (7) working days upon completion of the activity. If no liquidation was made within 1 month after the required 7 working days, FULL AMOUNT of the cash advance will be deducted from your salary (one-time deduction)</b><br><br>"
                "5. All acknowledgement receipts must indicate the following: (a) Date when cash was received; (b) Recipient information - name and/or contact details; (c) amount received; (d) duly signed by the recipient; and (e) any other relevant additional information. <b>(NOTE: A.R. must not be in the form given by Innogen and signed by recipient--- this is not accepted)</b><br><br>"
                )   
        print("---")
        for (name, email, cc_email), group in grouped:                
            body_lines = []
            bmsg2 = name
            indic = 'REIM'
            for index, row in group.iterrows():
                month = row.get('MONTH_YEAR', '')
                
                cator = row.get('INDICATOR', '')
                if cator != 'REIM':
                    indic = cator
                    
                subject = 'ONLINE TRANSACTION REQUEST - ' + row.get('MONTH_DAY_YEAR', '')
                bmsg1 = row.get('DATE ONLINE','')
                
                nickname = row.get('NICKNAME', '')
                if nickname:
                    bmsg2 = nickname
                                
                bmsg3 = row.get('DM#', '')
                bmsg4 = row.get('DV#', '')
                bmsg5 = f"{float(row.get('AMOUNT', 0)):,.2f}"
                bmsg6 = row.get('BANK', '')
                bmsg7 = row.get('DESCRIPTION', '')
                bmsg8 = row.get('PURPOSE', '')
                bmsg9 = row.get('REF.', '')

                body_lines.append(
                    f"<tr>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg1}</td>"
                    f"<td style='border: 1px solid #ddd;'><b>{name}</b></td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg3}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg4}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg5}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg6}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg7}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg8}</td>"
                    f"<td style='border: 1px solid #ddd;'>{bmsg9}</td>"
                    f"</tr>"                    
                )

            bmsg10 = "Please acknowledge upon receipt of this email and your respective CA/REIM. Thank you.<br><br>" if indic == "REIM" else bmsg10a
        
            body_template = (
                "<br>"
                "Hi Ma'am/Sir,<br><br>"
                "Good day, Please see details below;<br><br>"
                "<b>INNOGEN PHARMACEUTICALS, INC.</b><br>"
                "REQUESTED ONLINE-<br>"
                "MONTH: {month}<br><br>"
                "<table style='border-collapse: collapse; width: 100%;'>"
                "<tr style='background-color: #f2f2f2;'>"
                "<th style='text-align: left; border: 1px solid #ddd;'>DATE ONLINE</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>NAME</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>DM#</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>DV#</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>AMOUNT</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>BANK</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>DESCRIPTION</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>PURPOSE</th>"
                "<th style='text-align: left; border: 1px solid #ddd;'>REF.</th>"
                "</tr>"
                "{body_rows}"
                "</table>"
                "<br><br>"
                "{bmsg10}<br>"
                "<br>Regards,<br><br>{sendname}<br>"
                "<b>Finance Department</b>"
            )

            body = body_template.format(month=month, bmsg2=bmsg2, body_rows=''.join(body_lines), sendname=sendname, bmsg10=bmsg10)

            try:
                send_email(email, cc_email if pd.notnull(cc_email) else '', subject, body, username, password)
                print(f'Email sent TO: {email}',f'\nCC: {cc_email}\n')
                save_email_history(session, name, email, cc_email, subject, body, sendname)
            except Exception as e:
                print(f'Failed to send email TO: {email}: {e}')
                save_email_history(session, f"ERROR: {name}", f"ERROR: {email}", f"ERROR: {cc_email}", f"ERROR: {subject}", f"ERROR: {e}", sendname)
                input("\n\nPRESS ANY KEY TO EXIT\n\n")
                sys.exit(1)
                
        session.close()
        
        # This is additional codes only for Email Contacts References.
        # ‾----------------------------------------------‾
        tbl_name = "EmailHistory"
        sql_query = f"SELECT DISTINCT [Name],[Email],[Cc_Email] FROM {tbl_name} ORDER BY [Name]"

        df_contacts = pd.read_sql(sql_query, engine)
        df_contacts.to_excel("Email_Contacts.xlsx")
        # _----------------------------------------------_
        
        # This code is just to pause the program before it exit.
        input("\n\nDONE sending all email(s) in the list.\nPRESS ANY KEY TO EXIT\n\n")        
    
    except Exception as e:
        print()
        input(f"ERROR has been occured: {e} \n\nPRESS ANY KEY TO EXIT\n\n") 
        
if __name__ == '__main__':
    main()
