import streamlit as st
from email_processor import EmailProcessor, GmailAPIProcessor
from excel_exporter import ExcelExporter
import pandas as pd
from datetime import datetime, timedelta
import traceback


st.set_page_config(
    page_title="Email to Excel Converter",
    page_icon="📧",
    layout="wide"
)

st.title("📧 Email to Excel Converter")
st.write("Extract emails from your inbox and export them to Excel format")


if 'emails' not in st.session_state:
    st.session_state.emails = []
if 'connected' not in st.session_state:
    st.session_state.connected = False
if 'processor' not in st.session_state:
    st.session_state.processor = None


st.sidebar.header("⚙️ Configuration")

connection_type = st.sidebar.radio(
    "Connection Method",
    ["IMAP (Gmail, Outlook, etc.)", "Gmail API"],
    help="Choose how to connect to your email"
)

st.sidebar.markdown("---")

if connection_type == "IMAP (Gmail, Outlook, etc.)":
    st.sidebar.subheader("IMAP Settings")
    
    email_address = st.sidebar.text_input(
        "Email Address",
        placeholder="your.email@gmail.com"
    )
    
    password = st.sidebar.text_input(
        "Password / App Password",
        type="password",
        help="For Gmail, use an App Password: https://myaccount.google.com/apppasswords"
    )
    
    imap_server = st.sidebar.text_input(
        "IMAP Server",
        value="imap.gmail.com",
        help="Gmail: imap.gmail.com, Outlook: outlook.office365.com"
    )
    
    if st.sidebar.button("🔌 Connect", type="primary"):
        if not email_address or not password:
            st.sidebar.error("Please enter email and password")
        else:
            try:
                with st.spinner("Connecting to email server..."):
                    processor = EmailProcessor()
                    processor.connect_imap(email_address, password, imap_server)
                    st.session_state.processor = processor
                    st.session_state.connected = True
                    st.sidebar.success("✅ Connected successfully!")
            except Exception as e:
                st.sidebar.error(f"❌ Connection failed: {str(e)}")
                st.session_state.connected = False

else:
    st.sidebar.subheader("Gmail API Settings")
    st.sidebar.info("Gmail API requires OAuth2 credentials. Upload your credentials.json file.")
    
    credentials_file = st.sidebar.file_uploader(
        "Upload credentials.json",
        type=['json'],
        help="Download from Google Cloud Console"
    )
    
    if st.sidebar.button("🔌 Connect with Gmail API", type="primary"):
        if credentials_file is None:
            st.sidebar.error("Please upload credentials file")
        else:
            try:
                with open("credentials.json", "wb") as f:
                    f.write(credentials_file.getbuffer())
                
                with st.spinner("Authenticating with Gmail..."):
                    processor = GmailAPIProcessor()
                    processor.connect_gmail_api("credentials.json")
                    st.session_state.processor = processor
                    st.session_state.connected = True
                    st.sidebar.success("✅ Connected successfully!")
            except Exception as e:
                st.sidebar.error(f"❌ Connection failed: {str(e)}")
                st.session_state.connected = False

if st.session_state.connected:
    st.sidebar.success("🟢 Connected")
    
    if st.sidebar.button("🔌 Disconnect"):
        if hasattr(st.session_state.processor, 'disconnect'):
            st.session_state.processor.disconnect()
        st.session_state.connected = False
        st.session_state.processor = None
        st.session_state.emails = []
        st.rerun()


if st.session_state.connected:
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📥 Fetch Emails")
        
        fetch_col1, fetch_col2, fetch_col3 = st.columns(3)
        
        with fetch_col1:
            if connection_type == "IMAP (Gmail, Outlook, etc.)":
                try:
                    folders = st.session_state.processor.get_folders()
                    selected_folder = st.selectbox("Select Folder", folders, index=0 if "INBOX" in folders else 0)
                except:
                    selected_folder = st.text_input("Folder Name", value="INBOX")
            else:
                selected_folder = "All Mail"
                st.info("Gmail API fetches from all mail")
        
        with fetch_col2:
            email_limit = st.number_input(
                "Number of Emails",
                min_value=1,
                max_value=1000,
                value=50,
                help="Maximum number of emails to fetch"
            )
        
        with fetch_col3:
            search_option = st.selectbox(
                "Filter",
                ["All", "Unread", "Last 7 days", "Last 30 days", "Custom"]
            )
        
        if search_option == "Custom":
            custom_search = st.text_input(
                "Custom Search Criteria",
                placeholder="e.g., FROM 'sender@example.com'",
                help="IMAP search criteria"
            )
        else:
            search_map = {
                "All": "ALL",
                "Unread": "UNSEEN",
                "Last 7 days": f'SINCE {(datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")}',
                "Last 30 days": f'SINCE {(datetime.now() - timedelta(days=30)).strftime("%d-%b-%Y")}'
            }
            custom_search = search_map.get(search_option, "ALL")
        
        if st.button("📨 Fetch Emails", type="primary"):
            try:
                with st.spinner(f"Fetching emails from {selected_folder}..."):
                    if connection_type == "IMAP (Gmail, Outlook, etc.)":
                        emails = st.session_state.processor.fetch_emails(
                            folder=selected_folder,
                            limit=email_limit,
                            search_criteria=custom_search
                        )
                    else:
                        query = ""
                        if search_option == "Unread":
                            query = "is:unread"
                        elif search_option == "Last 7 days":
                            query = f"after:{(datetime.now() - timedelta(days=7)).strftime('%Y/%m/%d')}"
                        elif search_option == "Last 30 days":
                            query = f"after:{(datetime.now() - timedelta(days=30)).strftime('%Y/%m/%d')}"
                        
                        emails = st.session_state.processor.fetch_emails_api(
                            max_results=email_limit,
                            query=query
                        )
                    
                    st.session_state.emails = emails
                    st.success(f"✅ Fetched {len(emails)} emails successfully!")
            except Exception as e:
                st.error(f"❌ Error fetching emails: {str(e)}")
                st.code(traceback.format_exc())
    
    with col2:
        st.subheader("📊 Statistics")
        
        if st.session_state.emails:
            st.metric("Total Emails", len(st.session_state.emails))
            
            with_attachments = sum(1 for e in st.session_state.emails if e.get('has_attachments', False))
            st.metric("With Attachments", with_attachments)
            
            unique_senders = len(set(e.get('from', '') for e in st.session_state.emails))
            st.metric("Unique Senders", unique_senders)
        else:
            st.info("No emails fetched yet")
    
    
    if st.session_state.emails:
        st.markdown("---")
        st.subheader("📋 Email Preview")
        
        df = pd.DataFrame([
            {
                'Date': e.get('date', ''),
                'From': e.get('from', ''),
                'Subject': e.get('subject', ''),
                'Has Attachments': '✅' if e.get('has_attachments', False) else '❌',
                'Body Preview': (e.get('body', '')[:100] + '...') if len(e.get('body', '')) > 100 else e.get('body', '')
            }
            for e in st.session_state.emails
        ])
        
        st.dataframe(df, use_container_width=True, height=300)
        
        
        st.markdown("---")
        st.subheader("💾 Export Options")
        
        export_col1, export_col2, export_col3 = st.columns(3)
        
        with export_col1:
            export_format = st.radio(
                "Export Format",
                ["Excel (XLSX)", "Excel with Summary", "CSV"],
                help="Choose export format"
            )
        
        with export_col2:
            filename = st.text_input(
                "Filename (optional)",
                placeholder="emails_export",
                help="Leave empty for auto-generated name"
            )
        
        with export_col3:
            st.write("")
            st.write("")
            if st.button("📥 Export to File", type="primary"):
                try:
                    exporter = ExcelExporter()
                    
                    if export_format == "Excel (XLSX)":
                        output_file = exporter.create_excel_from_emails(
                            st.session_state.emails,
                            filename if filename else None
                        )
                    elif export_format == "Excel with Summary":
                        output_file = exporter.create_summary_sheet(
                            st.session_state.emails,
                            (filename if filename else f"emails_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}") + ".xlsx"
                        )
                    else:
                        output_file = exporter.export_to_csv(
                            st.session_state.emails,
                            filename if filename else None
                        )
                    
                    st.success(f"✅ Exported successfully: {output_file}")
                    
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="⬇️ Download File",
                            data=file,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if output_file.endswith('.xlsx') else "text/csv"
                        )
                except Exception as e:
                    st.error(f"❌ Export failed: {str(e)}")
                    st.code(traceback.format_exc())
        
        
        st.markdown("---")
        st.subheader("🔍 Detailed View")
        
        selected_email_idx = st.selectbox(
            "Select email to view details",
            range(len(st.session_state.emails)),
            format_func=lambda i: f"{st.session_state.emails[i].get('subject', 'No Subject')} - {st.session_state.emails[i].get('from', 'Unknown')}"
        )
        
        if selected_email_idx is not None:
            email = st.session_state.emails[selected_email_idx]
            
            detail_col1, detail_col2 = st.columns(2)
            
            with detail_col1:
                st.write("**From:**", email.get('from', 'N/A'))
                st.write("**To:**", email.get('to', 'N/A'))
            
            with detail_col2:
                st.write("**Date:**", email.get('date', 'N/A'))
                st.write("**Attachments:**", ', '.join(email.get('attachments', [])) or 'None')
            
            st.write("**Subject:**", email.get('subject', 'N/A'))
            
            st.text_area("**Body:**", email.get('body', 'No content'), height=200)

else:
    st.info("👈 Please connect to your email account using the sidebar to get started")
    
    st.markdown("---")
    st.subheader("📖 How to Use")
    
    st.markdown("""
    ### IMAP Connection (Recommended for most users)
    
    1. **For Gmail users:**
       - Enable 2-factor authentication
       - Generate an App Password: [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
       - Use the App Password instead of your regular password
       - IMAP Server: `imap.gmail.com`
    
    2. **For Outlook/Office 365 users:**
       - IMAP Server: `outlook.office365.com`
       - Use your regular email password
    
    3. **For other email providers:**
       - Find your IMAP server settings from your email provider
       - Enter the server address and your credentials
    
    ### Gmail API Connection (Advanced)
    
    1. Create a project in Google Cloud Console
    2. Enable Gmail API
    3. Create OAuth 2.0 credentials
    4. Download credentials.json
    5. Upload the file in the sidebar
    
    ### Features
    
    - ✅ Fetch emails from any folder
    - ✅ Filter by date, read status, or custom criteria
    - ✅ Preview emails before export
    - ✅ Export to Excel or CSV format
    - ✅ Include attachments information
    - ✅ Generate summary statistics
    """)


st.sidebar.markdown("---")
st.sidebar.markdown("### 📚 Resources")
st.sidebar.markdown("[Gmail App Passwords](https://myaccount.google.com/apppasswords)")
st.sidebar.markdown("[Google Cloud Console](https://console.cloud.google.com)")
