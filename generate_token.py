from google_auth_oauthlib.flow import InstalledAppFlow
import json

# Same scopes as in main.py
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.settings.basic'
]

def generate_token():
    """Generates a new token.json file with the required scopes."""
    try:
        # Create the flow using the client secrets file
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret.json',  # Make sure this file exists
            scopes=SCOPES
        )

        # Run the OAuth flow
        creds = flow.run_local_server(port=0)

        # Save the credentials to token.json
        token_data = {
            'token': creds.token,
            'refresh_token': creds.refresh_token,
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': creds.scopes,
            'universe_domain': 'googleapis.com'
        }

        with open('token.json', 'w') as token_file:
            json.dump(token_data, token_file)
            
        print("Successfully created token.json")
        
    except Exception as e:
        print(f"Error generating token: {str(e)}")

if __name__ == '__main__':
    generate_token() 