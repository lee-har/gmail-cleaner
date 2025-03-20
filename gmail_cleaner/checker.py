import re
import csv
from collections import Counter
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

def get_inbox_senders(max_results=100):
    creds = Credentials.from_authorized_user_file('token.json')
    service = build('gmail', 'v1', credentials=creds)
    
    sender_count = Counter()
    next_page_token = None
    fetched = 0

    while fetched < max_results:
        batch_size = min(500, max_results - fetched)  # Gmail API allows up to 500 per request
        results = service.users().messages().list(
            userId='me',
            labelIds=['INBOX'],
            maxResults=batch_size,
            pageToken=next_page_token
        ).execute()

        messages = results.get('messages', [])
        next_page_token = results.get('nextPageToken')

        if not messages:
            break

        for msg in messages:
            msg_detail = service.users().messages().get(
                userId='me',
                id=msg['id'],
                format='metadata',
                metadataHeaders=['From']
            ).execute()
            for header in msg_detail['payload']['headers']:
                if header['name'] == 'From':
                    match = re.search(r'<(.+?)>', header['value'])
                    email = match.group(1) if match else header['value']
                    sender_count[email] += 1
        fetched += len(messages)

        if not next_page_token:
            break

    sorted_senders = sorted(sender_count.items(), key=lambda x: x[1], reverse=True)

    print(f"\n📬 Top senders (out of {fetched} emails):")
    for sender, count in sorted_senders:
        print(f"{sender}: {count} emails")

    export_to_csv(sorted_senders)

def export_to_csv(senders_count, filename='senders_report.csv'):
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Sender Email', 'Email Count'])  # Header row
        
        for sender, count in senders_count:
            writer.writerow([sender, count])
    
    print(f"\n✅ Exported results to {filename}")

if __name__ == "__main__":
    user_input = input("How many email senders would you like to fetch? (e.g., 500): ").strip()
    if user_input.isdigit():
        get_inbox_senders(int(user_input))
    else:
        print("Invalid input. Please enter a valid number.")
