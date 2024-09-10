# Quick Email Scripts

A collection of simple scripts for email analysis and management.

## Current Scripts

### 1. `email_stats_by_domain.py`

Analyzes your Gmail account and provides statistics on email interactions by domain.

Features:
- Counts emails sent and received per domain
- Identifies most frequent contacts within each domain
- Generates a CSV report with detailed statistics

### 2. `email_stats_by_people.py`

Analyzes your Gmail account and provides statistics on email interactions with specific email addresses.

Features:
- Counts emails sent and received per email address
- Tracks first and last email dates for each contact
- Generates a CSV report with detailed statistics

## Usage

1. Clone the repository
2. Set up your Gmail credentials as environment variables:
   - `GMAIL_ADDRESS`: Your Gmail address
   - `GMAIL_APP_PASSWORD`: Your Gmail app password
3. Update the script you want to use:
   - For `email_stats_by_domain.py`: Edit the `domains` list in the `main()` function
   - For `email_stats_by_people.py`: Edit the `email_addresses` list in the `main()` function
   ```python
   email_addresses = [
     "example1@example.com",
     "example2@example.com"
     # Add or modify email addresses as needed
   ]
   ```
4. Run the desired script:
   ```
   python email_stats_by_domain.py
   ```
   or
   ```
   python email_stats_by_people.py
   ```
5. Check the generated `email_stats.csv` file for results

## Note

This is a personal project shared for convenience. While you're welcome to use and modify the scripts, I'm not actively seeking contributions or maintaining this as a collaborative project.

## License

This project is open source and available under the MIT License.

## Background

Created to quickly analyze email interactions with various contacts, organizations, and domains. Inspired by the need to efficiently assess communication history with potential business partners or investors.
