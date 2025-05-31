Incoming Payments ETL Script

This script performs the following steps:
1. Connects to a PostgreSQL database.
2. Retrieves partner information into a DataFrame.
3. Downloads today's "Incoming Payments" Excel attachments from Outlook.
4. Cleans up yesterday's files from the network share.
5. Loads and cleans the latest attachments.
6. Enriches them with partner metadata.
7. Exports the processed data to an Excel file on the network share.

All configurable parameters (DB creds, network paths, Outlook filter) are
provided via environment variables to avoid hard‑coding sensitive information.


