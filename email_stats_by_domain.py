import imaplib
import email
from email.header import decode_header
import re
import os
from collections import defaultdict
from datetime import datetime
from email.utils import parsedate_to_datetime

# Email credentials (fetched from environment variables)
EMAIL = os.environ['GMAIL_ADDRESS']
PASSWORD = os.environ['GMAIL_APP_PASSWORD']
IMAP_SERVER = "imap.gmail.com"

# Limit for how many emails to fetch at a time (if you want to limit results, adjust this)
MAX_EMAILS_TO_FETCH = 100000  # Adjust as needed
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

def fetch_emails_by_domain(mail, domains, max_emails):
    try:
        mail.select('"[Gmail]/All Mail"')
    except imaplib.IMAP4.abort as e:
        print(f"Error selecting All Mail folder: {e}")
        return defaultdict()

    domain_stats = defaultdict(lambda: {
        'people': defaultdict(lambda: {'to': 0, 'from': 0, 'total': 0}),
        'first_email_date': None,
        'last_email_date': None,
        'total_sent': 0,
        'total_received': 0,
        'total_emails': 0
    })

    for domain in domains:
        print(f"\nSearching for emails to/from domain: {domain}")

        # Define search criteria for TO or FROM the domain
        criteria = '(OR TO "@{}" FROM "@{}")'.format(domain, domain)

        try:
            status, messages = mail.search(None, criteria)
        except imaplib.IMAP4.abort as e:
            print(f"Error searching emails for domain {domain}: {e}")
            continue

        if status != 'OK':
            print(f"No emails found for {domain}.")
            continue

        email_ids = messages[0].split()
        email_ids = email_ids[-max_emails:]  # Limit to most recent `max_emails`

        print(f"Found {len(email_ids)} emails for domain {domain}. Processing now in batches...")

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

                        # Update first/last email dates for the domain
                        if domain_stats[domain]['first_email_date'] is None or email_date < domain_stats[domain]['first_email_date']:
                            domain_stats[domain]['first_email_date'] = email_date
                        if domain_stats[domain]['last_email_date'] is None or email_date > domain_stats[domain]['last_email_date']:
                            domain_stats[domain]['last_email_date'] = email_date

                        from_ = msg.get("From")
                        to_ = msg.get("To")

                        # Extract sender's email
                        from_email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', from_)
                        to_emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', to_)

                        # Process the FROM field
                        for email_addr in from_email:
                            if domain in email_addr:
                                name = email_addr
                                domain_stats[domain]['people'][name]['from'] += 1
                                domain_stats[domain]['people'][name]['total'] += 1
                                domain_stats[domain]['total_received'] += 1
                                domain_stats[domain]['total_emails'] += 1

                        # Process the TO field
                        for email_addr in to_emails:
                            if domain in email_addr:
                                name = email_addr
                                domain_stats[domain]['people'][name]['to'] += 1
                                domain_stats[domain]['people'][name]['total'] += 1
                                domain_stats[domain]['total_sent'] += 1
                                domain_stats[domain]['total_emails'] += 1
                    except Exception as e:
                        print(f"Error processing email ID in batch: {e}")
                        continue

    return domain_stats

def summarize_stats(domain_stats):
    sorted_domains = sorted(domain_stats.items(), key=lambda x: x[1]['total_emails'], reverse=True)
    for domain, stats in sorted_domains:
        print(f"\nSummary for domain: {domain}")
        print(f"  Total emails sent: {stats['total_sent']}")
        print(f"  Total emails received: {stats['total_received']}")
        print(f"  Total emails exchanged: {stats['total_emails']}")
        print(f"  First email: {stats['first_email_date'].strftime('%Y-%m-%d')}")
        print(f"  Last email: {stats['last_email_date'].strftime('%Y-%m-%d')}")

        sorted_people = sorted(stats['people'].items(), key=lambda x: x[1]['total'], reverse=True)
        for person, person_stats in sorted_people:
            print(f"  {person} - sent: {person_stats['to']}, received: {person_stats['from']}, total: {person_stats['total']}")

import pandas as pd

def save_stats_to_spreadsheet(domain_stats, filename="email_stats.csv"):
    data = []
    for domain, stats in domain_stats.items():
        for person, person_stats in stats['people'].items():
            data.append({
                "Domain": domain,
                "Person": person,
                "Sent": person_stats['to'],
                "Received": person_stats['from'],
                "Total": person_stats['total'],
                "First Email": stats['first_email_date'].strftime('%Y-%m-%d') if stats['first_email_date'] else None,
                "Last Email": stats['last_email_date'].strftime('%Y-%m-%d') if stats['last_email_date'] else None,
                "Total Emails Sent (Domain)": stats['total_sent'],
                "Total Emails Received (Domain)": stats['total_received'],
                "Total Emails Exchanged (Domain)": stats['total_emails']
            })

    # Create a pandas DataFrame
    df = pd.DataFrame(data)

    # Save to an Excel file
    df.to_csv(filename, index=False)
    print(f"CSV file saved to {filename}")


def main():
    domains = [
      "example.com",
      "example2.com"
    ]


    mail = connect_imap()
    if mail:
        domain_stats = fetch_emails_by_domain(mail, domains, MAX_EMAILS_TO_FETCH)
        mail.logout()

        summarize_stats(domain_stats)
        save_stats_to_spreadsheet(domain_stats, filename="email_stats.csv")

if __name__ == "__main__":
    main()
