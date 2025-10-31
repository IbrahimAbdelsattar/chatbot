"""
Test script for Email to Excel Agent
This script tests the core functionality without requiring actual email credentials
"""

from email_processor import EmailProcessor, GmailAPIProcessor
from excel_exporter import ExcelExporter
from datetime import datetime
import os


def test_excel_exporter():
    """Test Excel export functionality with sample data"""
    print("Testing Excel Exporter...")
    
    sample_emails = [
        {
            "id": "1",
            "subject": "Test Email 1",
            "from": "sender1@example.com",
            "to": "recipient@example.com",
            "date": datetime.now(),
            "body": "This is a test email body with some content.",
            "attachments": ["document.pdf", "image.jpg"],
            "has_attachments": True
        },
        {
            "id": "2",
            "subject": "Test Email 2",
            "from": "sender2@example.com",
            "to": "recipient@example.com",
            "date": datetime.now(),
            "body": "Another test email with different content.",
            "attachments": [],
            "has_attachments": False
        },
        {
            "id": "3",
            "subject": "Test Email 3 - Long Body",
            "from": "sender3@example.com",
            "to": "recipient@example.com",
            "date": datetime.now(),
            "body": "This is a longer email body. " * 50,
            "attachments": ["spreadsheet.xlsx"],
            "has_attachments": True
        }
    ]
    
    exporter = ExcelExporter()
    
    try:
        print("  ✓ Creating basic Excel file...")
        filename = exporter.create_excel_from_emails(sample_emails, "test_basic_export.xlsx")
        print(f"    Created: {filename}")
        assert os.path.exists(filename), "Excel file was not created"
        
        print("  ✓ Creating Excel with summary...")
        summary_filename = exporter.create_summary_sheet(sample_emails, "test_summary_export.xlsx")
        print(f"    Created: {summary_filename}")
        assert os.path.exists(summary_filename), "Summary Excel file was not created"
        
        print("  ✓ Creating CSV export...")
        csv_filename = exporter.export_to_csv(sample_emails, "test_export.csv")
        print(f"    Created: {csv_filename}")
        assert os.path.exists(csv_filename), "CSV file was not created"
        
        print("  ✓ Creating in-memory buffer...")
        buffer = exporter.create_excel_buffer(sample_emails)
        assert buffer.getbuffer().nbytes > 0, "Buffer is empty"
        print(f"    Buffer size: {buffer.getbuffer().nbytes} bytes")
        
        print("\n✅ Excel Exporter tests passed!")
        return True
        
    except Exception as e:
        print(f"\n❌ Excel Exporter test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def test_email_processor_structure():
    """Test EmailProcessor class structure (without actual connection)"""
    print("\nTesting Email Processor Structure...")
    
    try:
        print("  ✓ Creating EmailProcessor instance...")
        processor = EmailProcessor()
        assert processor is not None
        assert hasattr(processor, 'connect_imap')
        assert hasattr(processor, 'fetch_emails')
        assert hasattr(processor, 'get_folders')
        assert hasattr(processor, 'disconnect')
        
        print("  ✓ Creating GmailAPIProcessor instance...")
        gmail_processor = GmailAPIProcessor()
        assert gmail_processor is not None
        assert hasattr(gmail_processor, 'connect_gmail_api')
        assert hasattr(gmail_processor, 'fetch_emails_api')
        
        print("\n✅ Email Processor structure tests passed!")
        return True
        
    except Exception as e:
        print(f"\n❌ Email Processor test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def test_data_processing():
    """Test data processing and formatting"""
    print("\nTesting Data Processing...")
    
    try:
        from excel_exporter import ExcelExporter
        import pandas as pd
        
        exporter = ExcelExporter()
        
        test_emails = [
            {
                "subject": "Test with special chars: <>&\"'",
                "from": "test@example.com",
                "to": "recipient@example.com",
                "date": datetime.now(),
                "body": "Body with\nnewlines\nand\ttabs",
                "attachments": [],
                "has_attachments": False
            }
        ]
        
        print("  ✓ Testing DataFrame conversion...")
        df = exporter._emails_to_dataframe(test_emails)
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 1
        assert 'Subject' in df.columns
        assert 'From' in df.columns
        assert 'Body' in df.columns
        
        print("  ✓ Testing special character handling...")
        assert df.iloc[0]['Subject'] == "Test with special chars: <>&\"'"
        
        print("\n✅ Data processing tests passed!")
        return True
        
    except Exception as e:
        print(f"\n❌ Data processing test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def cleanup_test_files():
    """Clean up test files"""
    print("\nCleaning up test files...")
    test_files = [
        "test_basic_export.xlsx",
        "test_summary_export.xlsx",
        "test_export.csv"
    ]
    
    for file in test_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"  Removed: {file}")


def main():
    """Run all tests"""
    print("=" * 60)
    print("Email to Excel Agent - Test Suite")
    print("=" * 60)
    
    results = []
    
    results.append(("Email Processor Structure", test_email_processor_structure()))
    results.append(("Data Processing", test_data_processing()))
    results.append(("Excel Exporter", test_excel_exporter()))
    
    print("\n" + "=" * 60)
    print("Test Results Summary")
    print("=" * 60)
    
    for test_name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{test_name}: {status}")
    
    all_passed = all(result for _, result in results)
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 All tests passed successfully!")
    else:
        print("⚠️  Some tests failed. Please review the output above.")
    print("=" * 60)
    
    cleanup_test_files()
    
    return all_passed


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
