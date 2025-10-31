"""
Example Usage - Email to Excel Agent

This script demonstrates how to use the email agent programmatically
without the Streamlit UI.

NOTE: This is for demonstration purposes. For actual use, run the Streamlit app:
      streamlit run email_to_excel_app.py
"""

from email_processor import EmailProcessor
from excel_exporter import ExcelExporter
from datetime import datetime


def example_imap_usage():
    """
    Example: Connect via IMAP and export emails to Excel
    
    IMPORTANT: Replace with your actual credentials
    For Gmail, use an App Password: https://myaccount.google.com/apppasswords
    """
    
    print("=" * 60)
    print("Example: IMAP Email to Excel Export")
    print("=" * 60)
    
    EMAIL = "your.email@gmail.com"
    PASSWORD = "your-app-password-here"
    IMAP_SERVER = "imap.gmail.com"
    
    print("\n⚠️  This is a demonstration script.")
    print("Replace EMAIL and PASSWORD with your actual credentials.\n")
    
    try:
        print("Step 1: Connecting to email server...")
        processor = EmailProcessor()
        processor.connect_imap(EMAIL, PASSWORD, IMAP_SERVER)
        print("✅ Connected successfully!\n")
        
        print("Step 2: Fetching available folders...")
        folders = processor.get_folders()
        print(f"✅ Found {len(folders)} folders:")
        for folder in folders[:5]:
            print(f"   - {folder}")
        if len(folders) > 5:
            print(f"   ... and {len(folders) - 5} more\n")
        
        print("Step 3: Fetching emails from INBOX...")
        emails = processor.fetch_emails(
            folder="INBOX",
            limit=10,
            search_criteria="ALL"
        )
        print(f"✅ Fetched {len(emails)} emails\n")
        
        if emails:
            print("Step 4: Preview of first email:")
            first_email = emails[0]
            print(f"   From: {first_email.get('from', 'N/A')}")
            print(f"   Subject: {first_email.get('subject', 'N/A')}")
            print(f"   Date: {first_email.get('date', 'N/A')}")
            print(f"   Has Attachments: {first_email.get('has_attachments', False)}\n")
            
            print("Step 5: Exporting to Excel...")
            exporter = ExcelExporter()
            filename = exporter.create_excel_from_emails(emails, "example_export.xlsx")
            print(f"✅ Exported to: {filename}\n")
            
            print("Step 6: Creating Excel with summary...")
            summary_file = exporter.create_summary_sheet(emails, "example_with_summary.xlsx")
            print(f"✅ Created summary file: {summary_file}\n")
        
        print("Step 7: Disconnecting...")
        processor.disconnect()
        print("✅ Disconnected\n")
        
        print("=" * 60)
        print("✅ Example completed successfully!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        print("\nMake sure to:")
        print("1. Replace EMAIL and PASSWORD with your credentials")
        print("2. For Gmail, use an App Password")
        print("3. Ensure IMAP is enabled in your email settings")


def example_data_structure():
    """
    Example: Email data structure
    
    This shows the format of email data used by the agent
    """
    
    print("\n" + "=" * 60)
    print("Example: Email Data Structure")
    print("=" * 60 + "\n")
    
    sample_email = {
        "id": "12345",
        "subject": "Meeting Tomorrow",
        "from": "colleague@company.com",
        "to": "you@company.com",
        "date": datetime.now(),
        "body": "Hi, let's meet tomorrow at 2 PM to discuss the project.",
        "attachments": ["presentation.pdf", "budget.xlsx"],
        "has_attachments": True
    }
    
    print("Email object structure:")
    print("-" * 60)
    for key, value in sample_email.items():
        print(f"{key:20s}: {value}")
    print("-" * 60)


def example_excel_export_only():
    """
    Example: Export sample data to Excel without connecting to email
    
    This demonstrates the Excel export functionality with sample data
    """
    
    print("\n" + "=" * 60)
    print("Example: Excel Export with Sample Data")
    print("=" * 60 + "\n")
    
    sample_emails = [
        {
            "id": "1",
            "subject": "Project Update",
            "from": "manager@company.com",
            "to": "team@company.com",
            "date": datetime(2025, 10, 1, 10, 30),
            "body": "Here's the latest update on our project progress.",
            "attachments": ["report.pdf"],
            "has_attachments": True
        },
        {
            "id": "2",
            "subject": "Team Meeting Notes",
            "from": "secretary@company.com",
            "to": "team@company.com",
            "date": datetime(2025, 10, 15, 14, 0),
            "body": "Please find attached the notes from today's meeting.",
            "attachments": ["notes.docx"],
            "has_attachments": True
        },
        {
            "id": "3",
            "subject": "Quick Question",
            "from": "colleague@company.com",
            "to": "you@company.com",
            "date": datetime(2025, 10, 20, 9, 15),
            "body": "Do you have a moment to discuss the budget?",
            "attachments": [],
            "has_attachments": False
        }
    ]
    
    try:
        print("Creating Excel file from sample data...")
        exporter = ExcelExporter()
        
        filename = exporter.create_excel_from_emails(
            sample_emails,
            "sample_export.xlsx"
        )
        print(f"✅ Created: {filename}")
        
        summary_filename = exporter.create_summary_sheet(
            sample_emails,
            "sample_with_summary.xlsx"
        )
        print(f"✅ Created: {summary_filename}")
        
        csv_filename = exporter.export_to_csv(
            sample_emails,
            "sample_export.csv"
        )
        print(f"✅ Created: {csv_filename}")
        
        print("\n✅ Sample files created successfully!")
        print("\nYou can open these files to see the format:")
        print(f"  - {filename}")
        print(f"  - {summary_filename}")
        print(f"  - {csv_filename}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")


def main():
    """Main function to run examples"""
    
    print("\n" + "=" * 60)
    print("Email to Excel Agent - Usage Examples")
    print("=" * 60)
    
    print("\nAvailable examples:")
    print("1. IMAP connection and export (requires credentials)")
    print("2. Email data structure")
    print("3. Excel export with sample data (no credentials needed)")
    print("\nRunning example 2 and 3 (no credentials required)...\n")
    
    example_data_structure()
    
    example_excel_export_only()
    
    print("\n" + "=" * 60)
    print("To run the full application with UI:")
    print("  streamlit run email_to_excel_app.py")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
