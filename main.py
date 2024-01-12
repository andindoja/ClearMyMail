
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle
import os
import pandas as pd

from email.utils import parseaddr

from google_auth_httplib2 import Request
from httplib2 import Http  # Import Http from httplib2

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify']

def gmail_authenticate():
    print("Authenticating to Gmail...")
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request(Http()))  # Create an HTTP client and pass it
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "C:\\Users\\User\\PycharmProjects\\gmailReader\\credentials.json", SCOPES)
            creds = flow.run_local_server(port=8080)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('gmail', 'v1', credentials=creds)
    print("Authentication successful.")
    return service

def list_messages(service, user_id='me', query=''):
    print(f"Listing messages with query: '{query}'")
    try:
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])
        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, q=query, pageToken=page_token).execute()
            messages.extend(response['messages'])
        return messages
    except Exception as e:
        print(f'An error occurred while listing messages: {e}')
        return None

def get_message(service, user_id, msg_id):
    print(f"Getting message with ID: {msg_id}")
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='metadata', metadataHeaders=['From']).execute()
        return message
    except Exception as e:
        print(f'An error occurred while getting a message: {e}')
        return None

def list_email_addresses(service, user_id='me', query=''):
    print("Listing email addresses...")
    messages = list_messages(service, user_id, query)
    if not messages:
        print("No messages found.")
        return set()
    email_addresses = set()  # Using a set to avoid duplicates
    for msg in messages:
        message = get_message(service, user_id, msg['id'])
        headers = message['payload']['headers']
        for header in headers:
            if header['name'] == 'From':
                email_addresses.add(header['value'])
                break
    return email_addresses

def get_messages_batch(service, user_id, message_ids):
    def callback(request_id, response, exception):
        if exception is not None:
            print(f'An error occurred: {exception}')
        else:
            headers = response['payload']['headers']
            for header in headers:
                if header['name'] == 'From':
                    # Parse the email address using the email module
                    email_address = parseaddr(header['value'])[1]
                    email_addresses.add(email_address)

    email_addresses = set()
    batch = service.new_batch_http_request(callback=callback)
    for msg_id in message_ids:
        batch.add(service.users().messages().get(userId=user_id, id=msg_id, format='metadata', metadataHeaders=['From']))

    batch.execute()
    return email_addresses

def read_email_list(file_path):
    try:
        df = pd.read_excel(file_path)
        return df['Email Addresses'].tolist()  # Assuming the column name is 'Email Addresses'
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def read_email_list(file_path):
    try:
        df = pd.read_excel(file_path)
        email_list = df['Email Addresses'].tolist()  # Assuming the column name is 'Email Addresses'
        print("Email addresses to be deleted:", email_list)
        return email_list
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def delete_emails_from_sender(service, user_id, email_address):
    print(f"Attempting to delete emails from {email_address}")
    try:
        query = f"from:{email_address}"
        response = service.users().messages().list(userId=user_id, q=query).execute()
        messages_to_delete = response.get('messages', [])

        for message in messages_to_delete:
            service.users().messages().delete(userId=user_id, id=message['id']).execute()
            print(f"Deleted email from {email_address}, ID: {message['id']}")
    except Exception as e:
        print(f"Error deleting emails from {email_address}: {e}")

def main() :

    if os.path.exists('token.pickle'):
        os.remove('token.pickle')
        print("Old token removed, re-authentication required.")


    print("Starting main process...")
    service = gmail_authenticate()
    if not service :
        print("Authentication failed.")
        return

    # Get all message IDs (without limiting the max_results)
    messages = list_messages(service)
    if not messages :
        print("No messages found.")
        return

    # Fetch the message details in a batch to get email addresses
    BATCH_SIZE = 100  # Gmail API batch limit is 1000, but we use 100 to be safe
    email_addresses = set()

    # Process messages in batches
    for i in range(0, len(messages), BATCH_SIZE) :
        batch_message_ids = [msg['id'] for msg in messages[i :i + BATCH_SIZE]]
        batch_emails = get_messages_batch(service, 'me', batch_message_ids)
        email_addresses.update(batch_emails)
        print(f"Processed {i + len(batch_message_ids)} messages so far...")

    if email_addresses :
        print(f"Found {len(email_addresses)} unique email addresses:")
        # Convert the set of email addresses to a pandas DataFrame
        email_df = pd.DataFrame(list(email_addresses), columns=['Email Addresses'])
        # Save the DataFrame to an Excel file
        email_df.to_excel('email_addresses.xlsx', index=False)
        print("Email addresses have been saved to email_addresses.xlsx")
    else :
        print("No email addresses found.")

    input_path = input("Enter the path to the Excel file with email addresses to delete: ")
    # Automatically remove any surrounding quotation marks
    input_path = input_path.strip('"')

    email_list_to_delete = read_email_list(input_path)

    if not email_list_to_delete:
        print("No email addresses found in the provided file.")
        return

    for email_address in email_list_to_delete:
        delete_emails_from_sender(service, 'me', email_address)

    print("Email deletion process complete.")


if __name__ == '__main__' :
    main()