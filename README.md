# Bulk Send Email Script (Finance)

Sends formatted HTML emails to multiple recipients based on an Excel file, groups rows per recipient, and logs each sent email to a SQL Server table for history/audit purposes.

## What this script does

- Prompts you for:
  - Sender Gmail address
  - Sender display name (used in the email signature and history log)
- Lets you pick an Excel file via a file picker dialog.
- Reads the Excel rows, then groups them by `NAME`, `EMAIL`, and `CC`.
- Sends **one email per group** using Gmail SMTP (`smtp.gmail.com:587`) with an HTML table containing all rows in that group.
- Writes a history record per sent email (or error) to SQL Server table `EmailHistory`.
- Exports a reference file `Email_Contacts.xlsx` containing distinct contacts from `EmailHistory`.

## Key features

- One-email-per-recipient grouping (by `NAME` + `EMAIL` + `CC`)
- HTML email body with a table of requests
- Conditional body text based on `INDICATOR` (e.g., `REIM` vs others)
- SQL Server history logging (`EmailHistory` table auto-created if missing)
- Contact export from history (`Email_Contacts.xlsx`)

## Excel input expectations

The Excel file should contain these columns (case-insensitive; the script uses the values as strings):

- Required for grouping/sending:
  - `NAME`
  - `EMAIL`
  - `CC` (can be blank)
- Used in the email content:
  - `DATE ONLINE`
  - `INDICATOR`
  - `NICKNAME`
  - `DM#`
  - `DV#`
  - `AMOUNT`
  - `BANK`
  - `DESCRIPTION`
  - `PURPOSE`
  - `REF.`

The subject is formatted as:

`ONLINE TRANSACTION REQUEST - <Month Day, Year>`

## Output files

- `Email_Contacts.xlsx` is generated at the end of a successful run.
- `Send.txt` is an optional local secret file that can store your Gmail app password (see Security).

## Portable setup (recommended)

To make this portable (no Windows environment variable setup needed), create a local `.env` file in the same folder as the script. The script will auto-load it at startup.

1. Copy `.env.example` to `.env`
2. Fill in your real values
3. Run the script normally

Example `.env`:

```env
EMAIL_PASSWORD="your_gmail_app_password"
SQLSERVER_HOST="192.168.16.2\\INNOGENBC"
SQLSERVER_DB="RXTtracking"
SQLSERVER_USER="your_user"
SQLSERVER_PASSWORD="your_password"
SQLSERVER_DRIVER="ODBC Driver 17 for SQL Server"
```

## Security / Secrets (important for GitHub)

This repo is set up to keep secrets out of GitHub:

- `Send.txt` is ignored via `.gitignore`
- `.env` files are ignored via `.gitignore`

### Email password

Use one of these options (in order of priority):

1. Set environment variable `EMAIL_PASSWORD` (recommended)
2. Put the password/app-password into `Send.txt` (kept out of git by `.gitignore`)
3. If neither is provided, the script prompts securely (hidden input)

For Gmail, use an **App Password** (recommended) instead of your normal account password.

### SQL Server credentials

Set the SQL Server connection info using environment variables (recommended), or provide a full SQLAlchemy URL.

Option A (recommended): environment variables

- `SQLSERVER_HOST` (e.g. `192.168.16.2\\INSTANCE` or a hostname)
- `SQLSERVER_DB`
- `SQLSERVER_USER`
- `SQLSERVER_PASSWORD`
- `SQLSERVER_DRIVER` (optional, default: `SQL Server`)

Option B: one variable with a full SQLAlchemy URL

- `SQLSERVER_SQLALCHEMY_URL`

Example (Option A) PowerShell:

```powershell
$env:EMAIL_PASSWORD="your_gmail_app_password"
$env:SQLSERVER_HOST="192.168.16.2\INNOGENBC"
$env:SQLSERVER_DB="RXTtracking"
$env:SQLSERVER_USER="your_user"
$env:SQLSERVER_PASSWORD="your_password"
$env:SQLSERVER_DRIVER="ODBC Driver 17 for SQL Server"
python "Send Email To All v2b.py"
```

## Setup

Install dependencies:

```bash
pip install pandas sqlalchemy pyodbc openpyxl
```

Notes:

- `tkinter` is included with most standard Python installs on Windows.
- For `.xlsx` input, `openpyxl` is commonly required by `pandas.read_excel`.
- Ensure you have an ODBC driver installed (commonly: **ODBC Driver 17 for SQL Server**).

## How to run

```bash
python "Send Email To All v2b.py"
```

Then:

1. Enter your Gmail address
2. Enter your name (used in signature + logging)
3. Select the Excel file when the file picker opens
4. Wait for completion; the script prints per-recipient send results

## Database logging

Table: `EmailHistory`

Columns:

- `Name`, `Email`, `Cc_Email`, `Subject`, `Body`, `SenderName`, `SentDate`

If sending fails, an error record is saved with `ERROR:` prefixes for traceability.

## Troubleshooting

- Gmail authentication failing:
  - Use an App Password
  - Confirm `EMAIL_PASSWORD` is set (or `Send.txt` exists)
- SQL Server connection errors:
  - Confirm ODBC driver name in `SQLSERVER_DRIVER`
  - Confirm network/VPN access to the SQL Server host
  - Confirm credentials and database name

