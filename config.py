import os
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, '.env'))

TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL")
EXCEL_FILE_NAME = os.getenv("EXCEL_FILE", "appDeadlineSample.xlsx")

# Ensure the excel file path is correctly resolved relative to the project directory
if os.path.isabs(EXCEL_FILE_NAME):
    EXCEL_FILE = EXCEL_FILE_NAME
else:
    EXCEL_FILE = os.path.join(BASE_DIR, EXCEL_FILE_NAME)

SHEET_NAME = os.getenv("SHEET_NAME")
DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"
ENABLE_DEDUPLICATION = os.getenv("ENABLE_DEDUPLICATION", "true").lower() == "true"
REMINDER_DAYS_BEFORE = int(os.getenv("REMINDER_DAYS_BEFORE", 2))
TIMEZONE = os.getenv("TIMEZONE", "Asia/Kolkata")
