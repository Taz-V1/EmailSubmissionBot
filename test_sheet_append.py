from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import datetime

# Constants
SPREADSHEET_ID = '13AK7mmpuRHUc7yyP9l7VwEJKPLJ26MPja5o05Ib-J88'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def test_append_to_sheet():
    try:
        # Initialize sheets service
        print("Initializing sheets service...")
        creds = Credentials.from_service_account_file(
            'credentials.json', 
            scopes=SCOPES
        )
        sheet_service = build('sheets', 'v4', credentials=creds)
        
        # Test data
        test_data = {
            'link': 'https://linkedin.com/test-profile',
            'name': 'Test User',
            'company': 'Test Company',
            'position': 'Software Engineer',
            'email': 'test@example.com',
            'education': 'Test University',
            'same_college': True,
            'same_highschool': False,
            'thai': True
        }
        
        # Get current date
        current_date = datetime.datetime.now().strftime("%d/%m/%Y")
        
        # Build connections string
        connections = []
        if test_data['same_college']: connections.append("Same College")
        if test_data['same_highschool']: connections.append("Same Highschool")
        if test_data['thai']: connections.append("Thai")
        
        # Prepare row values
        values = [[
            "N/A",                                    # A: Offered help on
            "Not Yet",                               # B: Status
            test_data['name'],                       # C: Name
            test_data['company'],                    # D: Company
            "",                                      # E: Location
            "",                                      # F: Industry
            test_data['education'],                  # G: Education
            test_data['position'],                   # H: Position
            ", ".join(connections) if connections else "N/A",  # I: Connections
            test_data['email'],                      # J: Email
            test_data['link'],                       # K: LinkedIn URL
            "No",                                    # L: Coffee Chat Invite Sent?
            "No",                                    # M: Coffee Chat Yet?
            "No",                                    # N: Thank You Message Sent?
            "No",                                    # O: Connect on LinkedIn?
            "No",                                    # P: Follow Up 1
            "",                                      # Q: Coffee Chat Notes
            "Process",                               # R: Process with Bot?
            current_date                             # S: Last Updated Date
        ]]

        print("\nAttempting to append row with values:")
        for idx, val in enumerate(values[0]):
            print(f"Column {chr(65+idx)}: {val}")

        # Append to sheet
        print("\nSending append request...")
        result = sheet_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range="Sheet1!A:S",
            valueInputOption="USER_ENTERED",
            body={'values': values}
        ).execute()

        print("\nSuccess! Response:", result)
        
        # Verify the append
        print("\nVerifying append...")
        result = sheet_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Sheet1!A:S"
        ).execute()
        
        last_row = result.get('values', [])[-1]
        print("Last row in sheet:", last_row)

    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        raise

if __name__ == "__main__":
    test_append_to_sheet() 