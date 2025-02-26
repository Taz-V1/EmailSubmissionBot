import os
import base64
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify',
]

def get_gmail_service():
    """Gets an authorized Gmail API service instance."""
    try:
        # Load credentials from token.json
        if os.path.exists('token.json'):
            with open('token.json', 'r') as token_file:
                creds_data = json.load(token_file)
                creds = Credentials.from_authorized_user_info(creds_data, SCOPES)
                
                # Check if credentials need refresh
                if creds and creds.expired and creds.refresh_token:
                    print("Refreshing expired credentials...")
                    creds.refresh(Request())
                    
                    # Save refreshed credentials
                    with open('token.json', 'w') as token_file:
                        token_data = {
                            'token': creds.token,
                            'refresh_token': creds.refresh_token,
                            'token_uri': creds.token_uri,
                            'client_id': creds.client_id,
                            'client_secret': creds.client_secret,
                            'scopes': creds.scopes,
                            'universe_domain': 'googleapis.com'
                        }
                        json.dump(token_data, token_file)
                
                print("Successfully loaded credentials")
                return build('gmail', 'v1', credentials=creds)
        else:
            raise FileNotFoundError("token.json not found. Please run generate_token.py first.")
            
    except Exception as e:
        print(f"Error getting Gmail service: {str(e)}")
        raise

def create_test_draft():
    """Creates a test draft email with resume attachment."""
    try:
        # Get Gmail service
        print("Getting Gmail service...")
        service = get_gmail_service()
        
        # Create message
        print("Creating email message...")
        msg = MIMEMultipart()
        msg['From'] = "taz2547@gmail.com"  # Replace with your email
        msg['To'] = "taz2547sub@gmail.com"  # Replace with test recipient
        msg['Subject'] = "Test Draft Email with Resume"
        
        # Add body
        body = """
        Dear Recipient,

        This is a test draft email created via Gmail API with a resume attachment.

        Best regards,
        Nopparuj
        """
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach resume
        resume_path = os.path.join('Resources', 'NopparujVongpatarakulResume.pdf')
        if not os.path.exists(resume_path):
            print(f"Resume not found at: {resume_path}")
            resume_path = os.path.join('resources', 'NopparujVongpatarakulResume.pdf')  # Try lowercase folder
            if not os.path.exists(resume_path):
                raise FileNotFoundError(f"Resume not found in either Resources or resources folder")
        
        print(f"Attaching resume from: {resume_path}")
        with open(resume_path, 'rb') as f:
            resume = MIMEApplication(f.read(), _subtype='pdf')
            resume.add_header('Content-Disposition', 'attachment', 
                            filename='Nopparuj_Vongpatarakul_Resume.pdf')
            msg.attach(resume)
        
        # Encode the message
        print("Encoding message...")
        raw = base64.urlsafe_b64encode(msg.as_string().encode()).decode()
        
        # Create draft with labels
        print("Creating draft...")
        draft = service.users().drafts().create(
            userId='me',
            body={
                'message': {
                    'raw': raw,
                    'labelIds': ['DRAFT']
                }
            }
        ).execute()
        
        print(f"Draft created successfully! Draft ID: {draft['id']}")
        print("You can now find this draft in your Gmail account")
        return True
        
    except FileNotFoundError as e:
        print(f"File error: {str(e)}")
        return False
    except Exception as e:
        print(f"Error creating draft: {str(e)}")
        print(f"Error type: {type(e)}")
        if hasattr(e, 'error_details'):
            print(f"Error details: {e.error_details}")
        return False

if __name__ == "__main__":
    print("Starting test draft creation with resume attachment...")
    success = create_test_draft()
    print(f"\nFinal result: {'Success' if success else 'Failed'}") 