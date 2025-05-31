
import os
from datetime import date, datetime, timedelta
import pandas as pd
from sqlalchemy import create_engine, text
import win32com.client

# Configuration (read from environment variables) 

DB_USER = os.getenv("DB_USER", "your_user")
DB_PASSWORD = os.getenv("DB_PASSWORD", "your_password")
DB_HOST = os.getenv("DB_HOST", "127.0.0.1")
DB_NAME = os.getenv("DB_NAME", "your_db")

NETWORK_SHARE = os.getenv(
    "NETWORK_SHARE",
    r"\\<network_share>\Operational_Reports\Sales_and_Orders\Service_Files\Incoming_Payments",
)

OUTLOOK_SUBJECT_KEYWORD = os.getenv("OUTLOOK_SUBJECT_KEYWORD", "Incoming payments")
ATTACHMENT_PREFIX = "incoming_payments"

# Database 

connection_string = f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}/{DB_NAME}"
engine = create_engine(connection_string)

with engine.connect() as connection:
    query = """
        SELECT DISTINCT
               Контрагент_ИНН   AS "tax_id",
               Клиент           AS "dictionary_partner",
               Подразделение    AS "channel"
        FROM stg.sales_sources
    """
    partners_dict = connection.execute(text(query)).fetchall()

columns = ["tax_id", "dictionary_partner", "channel"]
agents = pd.DataFrame(partners_dict, columns=columns)
agents["tax_id"] = agents["tax_id"].astype(str).str.split(".").str[0]
agents["channel"] = agents["channel"].replace(
    {"ОРП": "RD", "ОЭП": "ED", "ОГП": "HD", "СМО": "RD"}
)
agents = agents.groupby(["tax_id", "dictionary_partner"], as_index=False).first()

# House‑keeping: remove outdated files 

today = date.today()

for fn in os.listdir(NETWORK_SHARE):
    if ATTACHMENT_PREFIX in fn and fn.lower().endswith((".xls", ".xlsx")):
        fpath = os.path.join(NETWORK_SHARE, fn)
        if date.fromtimestamp(os.path.getmtime(fpath)) < today:
            os.remove(fpath)
            print(f"Removed outdated file: {fn}")

# Fetch today's e‑mails and save attachments

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 == Inbox
messages_today = [
    msg
    for msg in inbox.Items
    if hasattr(msg, "ReceivedTime") and msg.ReceivedTime.date() == today
]

target_msgs = [
    msg for msg in messages_today if OUTLOOK_SUBJECT_KEYWORD.lower() in msg.Subject.lower()
]

attachments = []
for msg in target_msgs:
    for att in msg.Attachments:
        if att.FileName.lower().endswith((".xls", ".xlsx")):
            attachments.append(att)

for idx, att in enumerate(attachments, start=1):
    fname = f"{ATTACHMENT_PREFIX}_{idx}{os.path.splitext(att.FileName)[1]}"
    save_path = os.path.join(NETWORK_SHARE, fname)
    att.SaveAsFile(save_path)
    print(f"Saved attachment -> {fname}")

if not attachments:
    print("No matching attachments found today.")

# Helpers

def load_dataframe(path: str, header_keyword: str = "Оборот по кредиту") -> pd.DataFrame:
    """Read the Excel file and promote the row containing *header_keyword* to header."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    df = pd.read_excel(path, header=None)

    header_loc = df[df.apply(lambda x: x.astype(str).str.contains(header_keyword).any(), axis=1)].index
    if header_loc.empty:
        raise ValueError(f"Header '{header_keyword}' not found in {path}")

    df.columns = df.iloc[header_loc[0]]
    df = df.iloc[header_loc[0] + 1 :].reset_index(drop=True)
    return df


def strip_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows containing 'RUB' or numeric-only totals."""
    cleaned = df[~df["Оборот по кредиту"].str.contains("RUB", na=False)].copy()
    mask_numeric_total = cleaned["Оборот по кредиту"].apply(lambda x: isinstance(x, (int, float))) & cleaned[
        ["ИНН/Наименование контрагента"]
    ].isnull().all(axis=1)
    return cleaned[~mask_numeric_total]

# Load and process attachments 

file_1 = os.path.join(NETWORK_SHARE, f"{ATTACHMENT_PREFIX}_1.xlsx")
file_2 = os.path.join(NETWORK_SHARE, f"{ATTACHMENT_PREFIX}_2.xlsx")

merged_df = pd.DataFrame()

if os.path.exists(file_1):
    df1 = strip_totals(load_dataframe(file_1))
    merged_df = df1[["Оборот по кредиту", "ИНН/Наименование контрагента"]].copy()

if os.path.exists(file_2):
    df2 = pd.read_excel(file_2)
    new_rows = []
    for _, row in df2.iterrows():
        numeric = row[row.apply(lambda x: isinstance(x, (int, float)))]
        if not numeric.empty:
            new_rows.append(
                {
                    "Оборот по кредиту": numeric.sum(),
                    "ИНН/Наименование контрагента": "undefined",
                }
            )
    merged_df = pd.concat([merged_df, pd.DataFrame(new_rows)], ignore_index=True).drop_duplicates()
else:
    print("Second attachment not found; skipping.")

# Enrich with agent metadata 

merged_df["tax_id"] = merged_df["ИНН/Наименование контрагента"].apply(
    lambda x: x.split()[0] if isinstance(x, str) else None
)

merged_df = merged_df.merge(agents, on="tax_id", how="left")

merged_df["dictionary_partner"].fillna("undefined", inplace=True)
merged_df["channel"].fillna("RD", inplace=True)
merged_df["ИНН/Наименование контрагента"].fillna("undefined", inplace=True)

grouped = merged_df.groupby(["tax_id", "dictionary_partner", "channel"], as_index=False)["Оборот по кредиту"].sum()

# Yesterday's date (accounting for weekends)
now = datetime.now()
yesterday = now - timedelta(days=3 if now.weekday() == 0 else 1)

result = grouped.rename(columns={"Оборот по кредиту": "credit_turnover"})
result["date"] = yesterday.date()
result = result[["date", "dictionary_partner", "credit_turnover", "channel"]]

# Save result
output_path = os.path.join(NETWORK_SHARE, "processed_incoming_payments.xlsx")
result.to_excel(output_path, index=False)
print(f"Processing finished. Data saved to: {output_path}")
