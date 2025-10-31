# 🚀 Quick Start Guide - Email to Excel Agent

## What This Does

This application extracts emails from your inbox (Gmail, Outlook, or any IMAP-compatible email service) and exports them to a formatted Excel spreadsheet.

## Installation

```bash
# Install dependencies
pip install -r requirements.txt
```

## Running the Application

```bash
# Start the Streamlit app
streamlit run email_to_excel_app.py
```

The app will open in your browser at `http://localhost:8501`

## Quick Setup (Gmail Users)

### Step 1: Get an App Password

1. Go to your Google Account: https://myaccount.google.com
2. Select **Security**
3. Enable **2-Step Verification** (if not already enabled)
4. Go to **App passwords**: https://myaccount.google.com/apppasswords
5. Select **Mail** and your device
6. Click **Generate**
7. Copy the 16-character password

### Step 2: Connect in the App

1. In the sidebar, select **"IMAP (Gmail, Outlook, etc.)"**
2. Enter your details:
   - **Email Address**: your.email@gmail.com
   - **Password**: Paste the App Password (not your regular password)
   - **IMAP Server**: imap.gmail.com (default)
3. Click **"Connect"**

### Step 3: Fetch Emails

1. Select folder: **INBOX** (or any other folder)
2. Choose number of emails: **50** (or any number)
3. Apply filter if needed: **All**, **Unread**, **Last 7 days**, etc.
4. Click **"Fetch Emails"**

### Step 4: Export to Excel

1. Preview the emails in the table
2. Choose export format: **Excel (XLSX)**
3. Click **"Export to File"**
4. Click **"Download File"** to save

## For Outlook Users

Use these settings:
- **Email Address**: your.email@outlook.com
- **Password**: Your regular password
- **IMAP Server**: outlook.office365.com

## Features

✅ **Multiple Email Providers**: Gmail, Outlook, Yahoo, and more  
✅ **Flexible Filtering**: By date, folder, read status  
✅ **Preview Before Export**: See emails before downloading  
✅ **Professional Excel Format**: Formatted headers, borders, and styling  
✅ **Attachment Tracking**: See which emails have attachments  
✅ **Multiple Export Formats**: Excel, Excel with Summary, or CSV  

## Excel Output

Your exported file will contain:

| Column | Description |
|--------|-------------|
| Date | When the email was sent |
| From | Sender's email address |
| To | Recipient's email address |
| Subject | Email subject line |
| Body | Email content (plain text) |
| Has Attachments | Yes/No |
| Attachments | List of attachment names |

## Troubleshooting

### "Connection failed" error

**For Gmail:**
- Make sure you're using an **App Password**, not your regular password
- Verify 2-Step Verification is enabled
- Check that IMAP is enabled in Gmail settings

**For Outlook:**
- Verify your password is correct
- Check that IMAP is enabled in Outlook settings
- Try server: `imap-mail.outlook.com` if `outlook.office365.com` doesn't work

### "No emails found"

- Check the folder name is correct
- Try selecting "All" instead of filtered options
- Verify the folder contains emails

### Import errors

Make sure all dependencies are installed:
```bash
pip install -r requirements.txt
```

## Security Tips

🔒 **Never share your App Password**  
🔒 **Use App Passwords instead of regular passwords**  
🔒 **Don't commit credentials to version control**  

## Need Help?

1. Check the full documentation: `README_EMAIL_TO_EXCEL.md`
2. Review the troubleshooting section above
3. Verify your email provider's IMAP settings

## Example Use Cases

- 📊 **Email Analytics**: Export emails for analysis
- 📝 **Record Keeping**: Archive important emails
- 🔍 **Email Search**: Export and search in Excel
- 📧 **Backup**: Create backups of important emails
- 📈 **Reporting**: Generate email reports for teams

---

**Ready to start?** Run `streamlit run email_to_excel_app.py` and follow the steps above!
