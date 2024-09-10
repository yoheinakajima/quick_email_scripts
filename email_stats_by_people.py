import imaplib
import email
from email.header import decode_header
import re
import os
from collections import defaultdict
from datetime import datetime
from email.utils import parsedate_to_datetime
import pandas as pd

# Email credentials (fetched from environment variables)
EMAIL = os.environ['GMAIL_ADDRESS']
PASSWORD = os.environ['GMAIL_APP_PASSWORD']
IMAP_SERVER = "imap.gmail.com"

# Limit for how many emails to fetch at a time
MAX_EMAILS_TO_FETCH = 100000
BATCH_SIZE = 50  # Number of emails to fetch at once

def connect_imap():
    print("Connecting to Gmail IMAP server...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    try:
        mail.login(EMAIL, PASSWORD)
        print("Logged in successfully.")
    except imaplib.IMAP4.error as e:
        print(f"Login failed. Check your credentials. Error: {e}")
        return None
    return mail

def fetch_emails_by_address(mail, email_addresses, max_emails):
    try:
        mail.select('"[Gmail]/All Mail"')
    except imaplib.IMAP4.abort as e:
        print(f"Error selecting All Mail folder: {e}")
        return defaultdict()

    email_stats = defaultdict(lambda: {
        'first_email_date': None,
        'last_email_date': None,
        'total_sent': 0,
        'total_received': 0,
        'total_emails': 0
    })

    for email_address in email_addresses:
        print(f"\nSearching for emails to/from: {email_address}")

        # Define search criteria for TO or FROM the email address
        criteria = '(OR TO "{}" FROM "{}")'.format(email_address, email_address)

        try:
            status, messages = mail.search(None, criteria)
        except imaplib.IMAP4.abort as e:
            print(f"Error searching emails for address {email_address}: {e}")
            continue

        if status != 'OK':
            print(f"No emails found for {email_address}.")
            continue

        email_ids = messages[0].split()
        email_ids = email_ids[-max_emails:]  # Limit to most recent `max_emails`

        print(f"Found {len(email_ids)} emails for {email_address}. Processing now in batches...")

        for i in range(0, len(email_ids), BATCH_SIZE):
            batch = email_ids[i:i + BATCH_SIZE]
            batch_str = ",".join(batch.decode() for batch in batch)

            try:
                # Fetch only the headers (FROM, TO, DATE)
                status, msg_data = mail.fetch(batch_str, "(BODY[HEADER.FIELDS (FROM TO DATE)])")
                if status != 'OK':
                    print(f"Failed to fetch batch starting with ID {batch[0]}")
                    continue
            except imaplib.IMAP4.abort as e:
                print(f"Error fetching batch starting with email ID {batch[0]}: {e}")
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    try:
                        msg = email.message_from_bytes(response_part[1])

                        # Decode date and handle potential errors
                        date = msg.get("Date")
                        email_date = parsedate_to_datetime(date)

                        # Update first/last email dates
                        if email_stats[email_address]['first_email_date'] is None or email_date < email_stats[email_address]['first_email_date']:
                            email_stats[email_address]['first_email_date'] = email_date
                        if email_stats[email_address]['last_email_date'] is None or email_date > email_stats[email_address]['last_email_date']:
                            email_stats[email_address]['last_email_date'] = email_date

                        from_ = msg.get("From")
                        to_ = msg.get("To")

                        # Extract sender's email
                        from_email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', from_)
                        to_emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', to_)

                        # Process the FROM field (emails sent by the user)
                        if email_address in from_email:
                            email_stats[email_address]['total_sent'] += 1
                            email_stats[email_address]['total_emails'] += 1

                        # Process the TO field (emails received by the user)
                        if email_address in to_emails:
                            email_stats[email_address]['total_received'] += 1
                            email_stats[email_address]['total_emails'] += 1

                    except Exception as e:
                        print(f"Error processing email in batch: {e}")
                        continue

    return email_stats

def summarize_stats(email_stats):
    sorted_emails = sorted(email_stats.items(), key=lambda x: x[1]['total_emails'], reverse=True)
    for email_address, stats in sorted_emails:
        print(f"\nSummary for email: {email_address}")
        print(f"  Total emails sent: {stats['total_sent']}")
        print(f"  Total emails received: {stats['total_received']}")
        print(f"  Total emails exchanged: {stats['total_emails']}")
        print(f"  First email: {stats['first_email_date'].strftime('%Y-%m-%d') if stats['first_email_date'] else 'N/A'}")
        print(f"  Last email: {stats['last_email_date'].strftime('%Y-%m-%d') if stats['last_email_date'] else 'N/A'}")

def save_stats_to_csv(email_stats, filename="email_stats.csv"):
    data = []
    for email_address, stats in email_stats.items():
        data.append({
            "Email Address": email_address,
            "Sent": stats['total_sent'],
            "Received": stats['total_received'],
            "Total": stats['total_emails'],
            "First Email": stats['first_email_date'].strftime('%Y-%m-%d') if stats['first_email_date'] else None,
            "Last Email": stats['last_email_date'].strftime('%Y-%m-%d') if stats['last_email_date'] else None
        })

    # Create a pandas DataFrame
    df = pd.DataFrame(data)

    # Save to a CSV file
    df.to_csv(filename, index=False)
    print(f"CSV file saved to {filename}")

def main():
    email_addresses = ["example@email.com","example2@email.com"]

    mail = connect_imap()
    if mail:
        email_stats = fetch_emails_by_address(mail, email_addresses, MAX_EMAILS_TO_FETCH)
        mail.logout()

        # Summarize and save stats
        summarize_stats(email_stats)
        save_stats_to_csv(email_stats, filename="email_stats.csv")

if __name__ == "__main__":
    main()
