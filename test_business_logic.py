import pytest
from message_builder import build_consolidated_message, get_status_priority
import main
from excel_reader import read_tasks, read_config_params
from main import find_status_column, get_pending_activities
from unittest.mock import patch
import datetime
import os
import json

def test_get_status_priority():
    assert get_status_priority("In Progress") == 1
    assert get_status_priority("Blocked") == 2
    assert get_status_priority("Delayed") == 2
    assert get_status_priority("Not Started") == 3
    assert get_status_priority("Completed") == 4
    assert get_status_priority("Unknown") == 5

def test_build_consolidated_message():
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek",
                "pm_missing": False
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12)), ("QA R-Task", datetime.date(2026, 8, 15))]
        },
        {
            "app": "App2",
            "row": {
                "Overall Status": "Not Started",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 15),
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek",
                "pm_missing": False
            },
            "reminders": [("END_DATE", datetime.date(2026, 8, 15))],
            "pending_activities": [("QA OUR", datetime.date(2026, 8, 17))]
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    assert isinstance(cards, list)
    card = cards[0]
    
    assert card["type"] == "AdaptiveCard"
    assert card["body"][0]["text"] == "URGENT & IMPORTANT | ACCESS REQUESTS"
    
    attention_text = card["body"][1]["text"]
    assert "<at>Abishek</at>" in attention_text
    
    entities = card["msteams"]["entities"]
    assert len(entities) == 1
    assert entities[0]["mentioned"]["id"] == "amuralidharan@randomtrees.com"
    
    assert len(card["body"]) == 6
    assert "<at>Abishek</at>" in card["body"][3]["text"]
    
    app1_container = str(card["body"][4])
    app2_container = str(card["body"][5])
    
    assert "App1" in app1_container
    assert "Dev OUR — 12-Aug-2026" in app1_container
    assert "QA R-Task — 15-Aug-2026" in app1_container
    
    assert "App2" in app2_container
    assert "QA OUR — 17-Aug-2026" in app2_container
    assert "Dev OUR" not in app2_container

def test_deduplication_mention():
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek",
                "pm_missing": False
            },
            "reminders": [("START_DATE", None)],
            "pending_activities": []
        },
        {
            "app": "App2",
            "row": {
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek",
                "pm_missing": False
            },
            "reminders": [("START_DATE", None)],
            "pending_activities": []
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    card = cards[0]
    entities = card["msteams"]["entities"]
    assert len(entities) == 1

@patch('main.get_today')
def test_pending_activities_10_day_filter(mock_get_today):
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    
    base_row = {
        'Dev R-Task Status': 'Not Started',
        'Dev OUR Status': 'Not Started',
        'QA R-Task Status': 'Not Started',
        'Prod R-Task Status': 'Not Started',
        'Input Audit Status': 'Not Started',
        'Ready Status': 'Not Started'
    }

    row1 = dict(base_row)
    row1['Dev R-Task'] = datetime.date(2026, 8, 15)
    row1['QA R-Task'] = datetime.date(2026, 8, 25)
    pending1, _ = main.get_pending_activities(row1)
    act_names = [a[0] for a in pending1]
    assert act_names == ['Dev R-Task']

    row2 = dict(base_row)
    row2['Dev R-Task'] = datetime.date(2026, 8, 12)
    row2['QA R-Task'] = datetime.date(2026, 8, 19)
    row2['Prod R-Task'] = datetime.date(2026, 8, 25)
    pending2, _ = main.get_pending_activities(row2)
    act_names2 = [a[0] for a in pending2]
    assert set(act_names2) == {'Dev R-Task', 'QA R-Task'}

    row3 = dict(base_row)
    row3['Input Audit'] = datetime.date(2026, 8, 18)
    row3['Ready'] = datetime.date(2026, 8, 28)
    pending3, _ = main.get_pending_activities(row3)
    act_names3 = [a[0] for a in pending3]
    assert 'Input Audit' in act_names3

    row4 = dict(base_row)
    row4['Dev R-Task'] = datetime.date(2026, 8, 21)
    row4['QA R-Task'] = datetime.date(2026, 8, 30)
    pending4, _ = main.get_pending_activities(row4)
    assert pending4 == []

    row5 = dict(base_row)
    row5['Dev R-Task'] = datetime.date(2026, 8, 20)
    pending5, _ = main.get_pending_activities(row5)
    act_names5 = [a[0] for a in pending5]
    assert act_names5 == ['Dev R-Task']

    row6 = dict(base_row)
    row6['Sprint Start'] = datetime.date(2026, 8, 12)
    row6['Sprint End'] = datetime.date(2026, 8, 25)
    pending6, _ = main.get_pending_activities(row6)
    act_names6 = [a[0] for a in pending6]
    assert 'Sprint Start' in act_names6
    assert 'Sprint End' not in act_names6

# ==================== NEW WORKBOOK TESTS ====================

def test_read_config_params():
    file_path = "Application_Access_Release_Tracker 2.xlsx"
    if os.path.exists(file_path):
        params = read_config_params(file_path)
        assert params['SprintWarningDays'] == 10
        assert params['ReadyOffsetWD'] == -1
        assert params['SROffsetFromStartWD'] == -5
        assert params['OUROffsetFromSR_WD'] == -1
        assert params['RTSKOffsetFromOUR_WD'] == -2
        assert params['CABOffsetFromReleaseWD'] == -1
        assert params['CodeReviewCalendarDays'] == 2
        assert params['CodeReviewInviteCalendarDays'] == 4

def test_new_workbook_loading():
    file_path = "Application_Access_Release_Tracker 2.xlsx"
    if os.path.exists(file_path):
        tasks = read_tasks(file_path)
        assert len(tasks) == 68
        assert all(row['pm_missing'] is True for row in tasks)
        assert all(row['Responsible PM'] == "" for row in tasks)
        
        first_row = tasks[0]
        assert first_row['Application'] == 'ALLOCATIONS'
        assert first_row['APP'] == 'ALLOCATIONS'
        assert first_row['Sprint'] == 'Sprint 17'
        assert first_row['Overall Status'] == 'Yet to start'
        from date_checker import parse_date
        assert parse_date(first_row['Start Date']) == datetime.date(2026, 7, 29)
        assert parse_date(first_row['End Date']) == datetime.date(2026, 8, 11)

def test_find_status_column():
    group_cols = ['Dev R-Task', 'Raised On', 'Dev R-Task Status', 'Dev OUR', 'Dev OUR Status', 'Dev Service Request', 'Dev SR Status']
    
    assert find_status_column('Dev R-Task', group_cols) == 'Dev R-Task Status'
    assert find_status_column('Dev OUR', group_cols) == 'Dev OUR Status'
    assert find_status_column('Dev Service Request', group_cols) == 'Dev SR Status'
    
    group_cols2 = ['CAB Date', 'CAB Status', 'Code Review Date', 'Code Review Status', 'Code Review Invite', 'CR Invite Status']
    assert find_status_column('CAB Date', group_cols2) == 'CAB Status'
    assert find_status_column('Code Review Date', group_cols2) == 'Code Review Status'
    assert find_status_column('Code Review Invite', group_cols2) == 'CR Invite Status'
    
    assert find_status_column('Sprint Start', ['Sprint Start', 'Sprint End']) is None

@patch('main.get_today')
def test_get_pending_activities_with_multiindex(mock_get_today):
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    
    config_params = {'SprintWarningDays': 10}
    
    mock_row = {
        ('DEV DB REQUESTS', 'Dev R-Task'): datetime.date(2026, 8, 15),
        ('DEV DB REQUESTS', 'Raised On'): datetime.date(2026, 8, 5),
        ('DEV DB REQUESTS', 'Dev R-Task Status'): 'Not Started',
        
        ('DEV DB REQUESTS', 'Dev OUR'): datetime.date(2026, 8, 12),
        ('DEV DB REQUESTS', 'Dev OUR Status'): 'Completed',
        
        ('DEV DB REQUESTS', 'Dev Service Request'): datetime.date(2026, 8, 25),
        ('DEV DB REQUESTS', 'Dev SR Status'): 'Not Started',
        
        ('READY / READINESS', 'App Ready By'): datetime.date(2026, 8, 18),
        
        ('RELEASE GOVERNANCE', 'CAB Date'): datetime.date(2026, 8, 12),
        ('RELEASE GOVERNANCE', 'CAB Status'): 'Not Started',
        
        ('RELEASE GOVERNANCE', 'Code Review Invite'): datetime.date(2026, 8, 14),
        ('RELEASE GOVERNANCE', 'CR Invite Status'): 'Completed',
    }
    
    pending, all_statuses = get_pending_activities(mock_row, config_params)
    act_names = [p[0] for p in pending]
    
    assert 'Dev R-Task' in act_names
    assert 'Dev OUR' not in act_names
    assert 'Dev Service Request' not in act_names
    assert 'App Ready By' in act_names
    assert 'CAB Date' in act_names
    assert 'Code Review Invite' not in act_names

def test_message_builder_with_missing_pm():
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": True
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev R-Task", datetime.date(2026, 8, 12))]
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    card = cards[0]
    
    assert "Unassigned PM / Requires Review" in card["body"][1]["text"]
    assert "Unassigned PM / Requires Review" in card["body"][3]["text"]
    
    app_container = card["body"][4]
    warning_block = app_container["items"][1]
    assert warning_block["text"] == "⚠️ No PM assigned in tracker (Requires Review)"
    assert warning_block["color"] == "Attention"

def test_card_size_limit():
    qualifying_apps = []
    for i in range(30):
        qualifying_apps.append({
            "app": f"LargeApp_{i}",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": True
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [
                ("Dev R-Task", datetime.date(2026, 8, 12)),
                ("Dev OUR", datetime.date(2026, 8, 13)),
                ("Dev Service Request", datetime.date(2026, 8, 14)),
                ("Dev File R-Task", datetime.date(2026, 8, 12)),
                ("QA R-Task", datetime.date(2026, 8, 12)),
                ("QA OUR", datetime.date(2026, 8, 13)),
                ("QA Service Request", datetime.date(2026, 8, 14)),
                ("QA File R-Task", datetime.date(2026, 8, 12)),
                ("Prod R-Task", datetime.date(2026, 8, 12)),
                ("Prod OUR", datetime.date(2026, 8, 13)),
                ("Prod Service Request", datetime.date(2026, 8, 14)),
                ("Prod File R-Task", datetime.date(2026, 8, 12)),
            ]
        })
        
    cards = build_consolidated_message(qualifying_apps, max_bytes=15000)
    assert len(cards) > 1
    for card in cards:
        card_size = len(json.dumps(card).encode('utf-8'))
        assert card_size <= 15000
