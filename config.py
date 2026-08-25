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

# Centralized Business Logic Schema
ACTIONABLE_TASKS = {
    'Dev R-Task', 'Dev OUR', 'Dev Service Request', 'Dev File R-Task',
    'QA R-Task', 'QA OUR', 'QA Service Request', 'QA File R-Task',
    'Prod R-Task', 'Prod OUR', 'Prod Service Request', 'Prod File R-Task'
}

MILESTONE_COLUMNS = {
    'Sprint Start', 'Sprint End', 'Release Date',
    'Effective Sprint Start', 'Effective Sprint End', 'Effective Release Date',
    'App Ready By', 'Sprint Readiness Date', 'CAB Date',
    'Code Review Date', 'Code Review Invite'
}

PURE_MILESTONES = {
    'Sprint Start', 'Sprint End', 'Release Date',
    'Effective Sprint Start', 'Effective Sprint End', 'Effective Release Date',
    'App Ready By'
}

METADATA_COLUMNS = {
    'Next Action Date', 'Start Date', 'End Date', 'Application', 'Responsible PM',
    'Sprint', 'Overall Status', 'Next Action', 'Days Overdue', 'Validation / Notes',
    'Email/Teams', 'Name', 'APP', 'file_type', 'frequency',
    'File Type (all envs)', 'Frequency (all envs)'
}

