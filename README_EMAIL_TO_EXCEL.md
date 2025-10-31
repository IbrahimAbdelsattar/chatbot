# 📧 Email to Excel Converter

A Python application that extracts emails from your inbox and exports them to Excel format with a user-friendly Streamlit interface.

## Features

- ✅ **Multiple Connection Methods**: IMAP (Gmail, Outlook, etc.) and Gmail API
- ✅ **Flexible Filtering**: Filter by folder, date range, read status, or custom criteria
- ✅ **Email Preview**: View emails before exporting
- ✅ **Multiple Export Formats**: Excel (XLSX), Excel with Summary, or CSV
- ✅ **Attachment Tracking**: Track which emails have attachments
- ✅ **Statistics Dashboard**: View email statistics and metrics
- ✅ **Formatted Excel Output**: Professional formatting with headers and styling

## Installation

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

```bash
streamlit run email_to_excel_app.py
```

The application will open in your default web browser.

### Connection Methods

#### Option 1: IMAP Connection (Recommended)

**For Gmail:**
1. Enable 2-factor authentication on your Google account
2. Generate an App Password:
   - Go to [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Select "Mail" and your device
   - Copy the generated 16-character password
3. In the app:
   - Email: your.email@gmail.com
   - Password: Use the App Password (not your regular password)
   - IMAP Server: imap.gmail.com

**For Outlook/Office 365:**
- Email: your.email@outlook.com
- Password: Your regular password
- IMAP Server: outlook.office365.com

**For Other Providers:**
- Find your IMAP server settings from your email provider
- Enter the server address and credentials

#### Option 2: Gmail API (Advanced)

1. Create a project in [Google Cloud Console](https://console.cloud.google.com)
2. Enable Gmail API
3. Create OAuth 2.0 credentials (Desktop app)
4. Download credentials.json
5. Upload the file in the app sidebar
6. Complete the OAuth flow in your browser

### Fetching Emails

1. Connect to your email account using the sidebar
2. Select the folder/mailbox (e.g., INBOX, Sent, etc.)
3. Choose the number of emails to fetch
4. Apply filters if needed:
   - All emails
   - Unread only
   - Last 7 days
   - Last 30 days
   - Custom search criteria
5. Click "Fetch Emails"

### Exporting to Excel

1. After fetching emails, preview them in the table
2. Choose export format:
   - **Excel (XLSX)**: Standard Excel file with email data
   - **Excel with Summary**: Includes a summary sheet with statistics
   - **CSV**: Comma-separated values format
3. Optionally enter a custom filename
4. Click "Export to File"
5. Download the generated file

## Excel Output Format

The exported Excel file contains the following columns:

| Column | Description |
|--------|-------------|
| Date | Email date and time |
| From | Sender email address |
| To | Recipient email address |
| Subject | Email subject line |
| Body | Email body content (plain text) |
| Has Attachments | Yes/No indicator |
| Attachments | List of attachment filenames |

### Excel Features

- **Header Row**: Bold, colored header with centered text
- **Column Widths**: Auto-adjusted for readability
- **Text Wrapping**: Enabled for long content
- **Borders**: Clean borders around all cells
- **Frozen Header**: Header row stays visible when scrolling
- **Date Formatting**: Proper date/time formatting

## Project Structure

```
.
├── email_to_excel_app.py    # Main Streamlit application
├── email_processor.py        # Email fetching and parsing logic
├── excel_exporter.py         # Excel file generation
├── requirements.txt          # Python dependencies
└── README_EMAIL_TO_EXCEL.md  # This file
```

## Modules

### email_processor.py

Contains two classes:
- **EmailProcessor**: IMAP-based email fetching
  - `connect_imap()`: Connect to email server
  - `get_folders()`: List available folders
  - `fetch_emails()`: Fetch emails with filters
  - `disconnect()`: Close connection

- **GmailAPIProcessor**: Gmail API-based fetching
  - `connect_gmail_api()`: Authenticate with OAuth2
  - `fetch_emails_api()`: Fetch emails using API

### excel_exporter.py

Contains the ExcelExporter class:
- `create_excel_from_emails()`: Create basic Excel file
- `create_excel_buffer()`: Create in-memory Excel for download
- `create_summary_sheet()`: Create Excel with summary statistics
- `export_to_csv()`: Export to CSV format

### email_to_excel_app.py

Main Streamlit application with:
- Connection configuration UI
- Email fetching interface
- Preview and statistics
- Export functionality
- Detailed email viewer

## Troubleshooting

### IMAP Connection Issues

**Gmail:**
- Ensure 2-factor authentication is enabled
- Use App Password, not regular password
- Check if IMAP is enabled in Gmail settings

**Outlook:**
- Verify IMAP is enabled in account settings
- Use the correct server: outlook.office365.com

**Other Providers:**
- Verify IMAP server address
- Check if IMAP access is enabled
- Some providers require app-specific passwords

### Gmail API Issues

- Ensure Gmail API is enabled in Google Cloud Console
- Verify OAuth consent screen is configured
- Check that credentials.json is for a Desktop application
- Delete token.pickle and re-authenticate if needed

### Export Issues

- Ensure you have write permissions in the directory
- Check available disk space
- For large exports, try reducing the number of emails

## Security Notes

- **Never commit credentials**: Don't store passwords in code
- **App Passwords**: Use app-specific passwords when available
- **OAuth Tokens**: token.pickle contains sensitive data
- **Credentials File**: Keep credentials.json secure

## Dependencies

- **streamlit**: Web application framework
- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file creation and formatting
- **google-auth**: Google authentication library
- **google-api-python-client**: Gmail API client

## License

See LICENSE file for details.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Review email provider documentation
3. Verify IMAP/API settings

## Future Enhancements

Potential features for future versions:
- Email search by keywords
- Bulk operations
- Email threading support
- Attachment download
- Multiple account support
- Scheduled exports
- Email analytics and charts
