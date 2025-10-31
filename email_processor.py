import imaplib
import email
from email.header import decode_header
from datetime import datetime
from typing import List, Dict, Optional
import base64
import re


class EmailProcessor:
    """Process emails from IMAP or Gmail API"""
    
    def __init__(self):
        self.connection = None
        self.email_address = None
    
    def connect_imap(self, email_address: str, password: str, imap_server: str = "imap.gmail.com") -> bool:
        """Connect to email server using IMAP"""
        try:
            self.connection = imaplib.IMAP4_SSL(imap_server)
            self.connection.login(email_address, password)
            self.email_address = email_address
            return True
        except Exception as e:
            raise Exception(f"Failed to connect: {str(e)}")
    
    def get_folders(self) -> List[str]:
        """Get list of available email folders"""
        if not self.connection:
            raise Exception("Not connected to email server")
        
        try:
            status, folders = self.connection.list()
            folder_list = []
            for folder in folders:
                folder_name = folder.decode().split('"')[-2]
                folder_list.append(folder_name)
            return folder_list
        except Exception as e:
            raise Exception(f"Failed to get folders: {str(e)}")
    
    def fetch_emails(self, folder: str = "INBOX", limit: Optional[int] = None, 
                     search_criteria: str = "ALL") -> List[Dict]:
        """Fetch emails from specified folder"""
        if not self.connection:
            raise Exception("Not connected to email server")
        
        try:
            self.connection.select(folder)
            status, messages = self.connection.search(None, search_criteria)
            
            if status != "OK":
                raise Exception("Failed to search emails")
            
            email_ids = messages[0].split()
            
            if limit:
                email_ids = email_ids[-limit:]
            
            emails = []
            for email_id in email_ids:
                email_data = self._fetch_email_by_id(email_id)
                if email_data:
                    emails.append(email_data)
            
            return emails
        except Exception as e:
            raise Exception(f"Failed to fetch emails: {str(e)}")
    
    def _fetch_email_by_id(self, email_id: bytes) -> Optional[Dict]:
        """Fetch and parse a single email by ID"""
        try:
            status, msg_data = self.connection.fetch(email_id, "(RFC822)")
            
            if status != "OK":
                return None
            
            email_body = msg_data[0][1]
            email_message = email.message_from_bytes(email_body)
            
            subject = self._decode_header(email_message["Subject"])
            from_addr = self._decode_header(email_message["From"])
            to_addr = self._decode_header(email_message["To"])
            date_str = email_message["Date"]
            
            date_obj = self._parse_date(date_str)
            
            body, attachments = self._extract_body_and_attachments(email_message)
            
            return {
                "id": email_id.decode(),
                "subject": subject,
                "from": from_addr,
                "to": to_addr,
                "date": date_obj,
                "body": body,
                "attachments": attachments,
                "has_attachments": len(attachments) > 0
            }
        except Exception as e:
            print(f"Error processing email {email_id}: {str(e)}")
            return None
    
    def _decode_header(self, header: str) -> str:
        """Decode email header"""
        if not header:
            return ""
        
        decoded_parts = decode_header(header)
        decoded_string = ""
        
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                try:
                    decoded_string += part.decode(encoding or "utf-8")
                except:
                    decoded_string += part.decode("utf-8", errors="ignore")
            else:
                decoded_string += part
        
        return decoded_string
    
    def _parse_date(self, date_str: str) -> datetime:
        """Parse email date string"""
        try:
            return email.utils.parsedate_to_datetime(date_str)
        except:
            return datetime.now()
    
    def _extract_body_and_attachments(self, email_message) -> tuple:
        """Extract email body and attachments"""
        body = ""
        attachments = []
        
        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        attachments.append(self._decode_header(filename))
                elif content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body += part.get_payload(decode=True).decode()
                    except:
                        pass
                elif content_type == "text/html" and not body:
                    try:
                        html_body = part.get_payload(decode=True).decode()
                        body = self._html_to_text(html_body)
                    except:
                        pass
        else:
            try:
                body = email_message.get_payload(decode=True).decode()
            except:
                body = str(email_message.get_payload())
        
        return body.strip(), attachments
    
    def _html_to_text(self, html: str) -> str:
        """Convert HTML to plain text"""
        text = re.sub('<[^<]+?>', '', html)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def disconnect(self):
        """Disconnect from email server"""
        if self.connection:
            try:
                self.connection.close()
                self.connection.logout()
            except:
                pass
            self.connection = None


class GmailAPIProcessor:
    """Process emails using Gmail API"""
    
    def __init__(self):
        self.service = None
    
    def connect_gmail_api(self, credentials_path: str):
        """Connect to Gmail using API credentials"""
        try:
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from google.auth.transport.requests import Request
            from googleapiclient.discovery import build
            import os.path
            import pickle
            
            SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
            creds = None
            
            if os.path.exists('token.pickle'):
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        credentials_path, SCOPES)
                    creds = flow.run_local_server(port=0)
                
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
            
            self.service = build('gmail', 'v1', credentials=creds)
            return True
        except Exception as e:
            raise Exception(f"Failed to connect to Gmail API: {str(e)}")
    
    def fetch_emails_api(self, max_results: int = 100, query: str = "") -> List[Dict]:
        """Fetch emails using Gmail API"""
        if not self.service:
            raise Exception("Not connected to Gmail API")
        
        try:
            results = self.service.users().messages().list(
                userId='me', maxResults=max_results, q=query
            ).execute()
            
            messages = results.get('messages', [])
            emails = []
            
            for message in messages:
                email_data = self._fetch_email_by_id_api(message['id'])
                if email_data:
                    emails.append(email_data)
            
            return emails
        except Exception as e:
            raise Exception(f"Failed to fetch emails: {str(e)}")
    
    def _fetch_email_by_id_api(self, msg_id: str) -> Optional[Dict]:
        """Fetch single email using Gmail API"""
        try:
            message = self.service.users().messages().get(
                userId='me', id=msg_id, format='full'
            ).execute()
            
            headers = message['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
            from_addr = next((h['value'] for h in headers if h['name'] == 'From'), '')
            to_addr = next((h['value'] for h in headers if h['name'] == 'To'), '')
            date_str = next((h['value'] for h in headers if h['name'] == 'Date'), '')
            
            date_obj = email.utils.parsedate_to_datetime(date_str) if date_str else datetime.now()
            
            body = self._get_body_api(message['payload'])
            
            attachments = []
            if 'parts' in message['payload']:
                for part in message['payload']['parts']:
                    if part.get('filename'):
                        attachments.append(part['filename'])
            
            return {
                "id": msg_id,
                "subject": subject,
                "from": from_addr,
                "to": to_addr,
                "date": date_obj,
                "body": body,
                "attachments": attachments,
                "has_attachments": len(attachments) > 0
            }
        except Exception as e:
            print(f"Error processing email {msg_id}: {str(e)}")
            return None
    
    def _get_body_api(self, payload) -> str:
        """Extract body from Gmail API payload"""
        body = ""
        
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    if 'data' in part['body']:
                        body = base64.urlsafe_b64decode(part['body']['data']).decode()
                        break
                elif part['mimeType'] == 'text/html' and not body:
                    if 'data' in part['body']:
                        html = base64.urlsafe_b64decode(part['body']['data']).decode()
                        body = re.sub('<[^<]+?>', '', html)
        elif 'body' in payload and 'data' in payload['body']:
            body = base64.urlsafe_b64decode(payload['body']['data']).decode()
        
        return body.strip()
