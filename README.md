# Quick Email Scripts

A collection of simple scripts for email analysis and management.

## Current Scripts

### `email_stats_by_domain.py`

This script analyzes your Gmail account and provides statistics on email interactions by domain.

Features:
- Counts emails sent and received per domain
- Identifies most frequent contacts within each domain
- Generates a CSV report with detailed statistics

Useful for quickly assessing communication patterns with various organizations or companies.

## Usage

1. Clone the repository
2. Set up your Gmail credentials as environment variables:
   - `GMAIL_ADDRESS`: Your Gmail address
   - `GMAIL_APP_PASSWORD`: Your Gmail app password
3. Open `email_stats_by_domain.py` and update the `domains` list in the `main()` function:
   ```python
   domains = [
     "example.com",
     "example2.com"
     # Add or modify domains as needed
   ]
   ```
4. Run the script:
   ```
   python email_stats_by_domain.py
   ```
5. Check the generated `email_stats.csv` file for results

## Note

This is a personal project shared for convenience. While you're welcome to use and modify the scripts, I'm not actively seeking contributions or maintaining this as a collaborative project.

## License

This project is open source and available under the MIT License.

## Background

Created to quickly analyze email interactions with various contacts and organizations. Inspired by the need to efficiently assess communication history with potential business partners or investors.
