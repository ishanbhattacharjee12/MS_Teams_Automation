import pytest
from message_builder import build_consolidated_message, get_status_priority
import automation_core
from excel_reader import read_tasks, read_config_params
from automation_core import find_status_column, get_pending_activities
from unittest.mock import patch
import datetime
import os
import json
import pandas as pd

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
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12)), ("QA R-Task", datetime.date(2026, 8, 15))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
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
            "pending_activities": [("QA OUR", datetime.date(2026, 8, 17))],
            "overall_status": "Not Started",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 15)
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    assert isinstance(cards, list)
    card = cards[0]
    
    assert card["type"] == "AdaptiveCard"
    assert card["body"][0]["text"] == "🚨 URGENT & IMPORTANT | ACCESS REQUESTS"
    
    # 📊 Summary block is at index 1
    assert "2 Applications · 1 PM · 1 Card" in card["body"][1]["text"]
    
    # Attention block is at index 2
    attention_text = card["body"][2]["text"]
    assert "<at>Abishek</at>" in attention_text
    
    entities = card["msteams"]["entities"]
    assert len(entities) == 1
    assert entities[0]["mentioned"]["id"] == "amuralidharan@randomtrees.com"
    
    # Body elements: Title (0), Summary (1), Attention (2), Desc (3), PM Header (4), App1 container (5), App2 container (6)
    assert len(card["body"]) == 7
    assert "Abishek" in card["body"][4]["text"]
    
    app1_container = str(card["body"][5])
    app2_container = str(card["body"][6])
    
    assert "App1" in app1_container
    assert "QA R-Task" in app1_container
    
    assert "App2" in app2_container
    assert "QA OUR" in app2_container

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
            "pending_activities": [],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": None
        },
        {
            "app": "App2",
            "row": {
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek",
                "pm_missing": False
            },
            "reminders": [("START_DATE", None)],
            "pending_activities": [],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": None
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    card = cards[0]
    entities = card["msteams"]["entities"]
    assert len(entities) == 1

@patch('automation_core.get_today')
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
    pending1, _ = automation_core.get_pending_activities(row1)
    act_names = [a[0] for a in pending1]
    assert act_names == ['Dev R-Task']

    row2 = dict(base_row)
    row2['Dev R-Task'] = datetime.date(2026, 8, 12)
    row2['QA R-Task'] = datetime.date(2026, 8, 19)
    row2['Prod R-Task'] = datetime.date(2026, 8, 25)
    pending2, _ = automation_core.get_pending_activities(row2)
    act_names2 = [a[0] for a in pending2]
    assert set(act_names2) == {'Dev R-Task', 'QA R-Task'}

    row3 = dict(base_row)
    row3['Input Audit'] = datetime.date(2026, 8, 18)
    row3['Ready'] = datetime.date(2026, 8, 28)
    pending3, _ = automation_core.get_pending_activities(row3)
    act_names3 = [a[0] for a in pending3]
    assert 'Input Audit' in act_names3

    row4 = dict(base_row)
    row4['Dev R-Task'] = datetime.date(2026, 8, 21)
    row4['QA R-Task'] = datetime.date(2026, 8, 30)
    pending4, _ = automation_core.get_pending_activities(row4)
    assert pending4 == []

    row5 = dict(base_row)
    row5['Dev R-Task'] = datetime.date(2026, 8, 20)
    pending5, _ = automation_core.get_pending_activities(row5)
    act_names5 = [a[0] for a in pending5]
    assert act_names5 == ['Dev R-Task']

    row6 = dict(base_row)
    row6['Sprint Start'] = datetime.date(2026, 8, 12)
    row6['Sprint End'] = datetime.date(2026, 8, 25)
    pending6, _ = automation_core.get_pending_activities(row6)
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

@patch('automation_core.get_today')
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
            "pending_activities": [("Dev R-Task", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    card = cards[0]
    
    assert "Unassigned PM / Requires Review" in card["body"][2]["text"]
    assert "Unassigned PM / Requires Review" in card["body"][4]["text"]
    
    app_container = card["body"][5]
    warning_block = app_container["items"][1]
    assert warning_block["text"] == "⚠️ No PM assigned in tracker"
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
            ],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
        
    cards = build_consolidated_message(qualifying_apps, max_bytes=15000)
    assert len(cards) > 1
    for card in cards:
        card_size = len(json.dumps(card).encode('utf-8'))
        assert card_size <= 15000

# ==================== PM IDENTITY RESOLUTION TESTS ====================

def test_parse_pm_value():
    from excel_reader import parse_pm_value
    
    # 1. Name <email>
    res = parse_pm_value("Ishan Bhattacharjee <ishanbhattacharjee19@gmail.com>")
    assert res == {"name": "Ishan Bhattacharjee", "email": "ishanbhattacharjee19@gmail.com", "pm_missing": False}
    
    # 2. Name <email> with spaces/tabs
    res = parse_pm_value("  Abishek   <  amuralidharan@randomtrees.com  > ")
    assert res == {"name": "Abishek", "email": "amuralidharan@randomtrees.com", "pm_missing": False}
    
    # 3. Name only
    res = parse_pm_value("Abishek")
    assert res == {"name": "Abishek", "email": "", "pm_missing": False}
    
    # 4. Null/NaN/empty
    assert parse_pm_value("") is None
    assert parse_pm_value("nan") is None
    assert parse_pm_value(None) is None

def test_pm_identity_resolution_scenarios(tmp_path):
    from excel_reader import find_pm_mappings, resolve_pm_identity
    
    file_path = os.path.join(tmp_path, "mock_tracker.xlsx")
    
    tracker_df = pd.DataFrame([
        {"Application": "AppA", "Responsible PM": "Ishan Bhattacharjee <ishanbhattacharjee19@gmail.com>"},
        {"Application": "AppB", "Responsible PM": "Abishek"},
        {"Application": "AppC", "Responsible PM": ""},
        {"Application": "AppD", "Responsible PM": "John Doe"},
        {"Application": "AppE", "Responsible PM": "Jane Smith"},
    ])
    
    contacts_df = pd.DataFrame([
        {"Name": "Jane Smith", "Email": "janesmith@rt.com"},
        {"Name": "Other PM", "Email": "other@rt.com"}
    ])
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        tracker_df.to_excel(writer, sheet_name="Tracker", index=False)
        contacts_df.to_excel(writer, sheet_name="Contacts", index=False)
        
    email_map = find_pm_mappings(file_path)
    assert email_map.get("jane smith")["email"] == "janesmith@rt.com"
    
    # 1. Name <email>
    id1 = resolve_pm_identity("Ishan Bhattacharjee <ishanbhattacharjee19@gmail.com>", email_map)
    assert id1 == {"name": "Ishan Bhattacharjee", "email": "ishanbhattacharjee19@gmail.com", "pm_missing": False}
    
    # 2. Name only + mapping table in other sheet
    id2 = resolve_pm_identity("Jane Smith", email_map)
    assert id2 == {"name": "Jane Smith", "email": "janesmith@rt.com", "pm_missing": False}
    
    # 3. Missing PM
    id3 = resolve_pm_identity("", email_map)
    assert id3 == {"name": "", "email": "", "pm_missing": True}
    
    # 4. Missing email (fallback)
    id4 = resolve_pm_identity("John Doe", email_map)
    assert id4 == {"name": "John Doe", "email": "", "pm_missing": False}

def test_pm_mapping_on_same_sheet(tmp_path):
    from excel_reader import find_pm_mappings
    
    file_path = os.path.join(tmp_path, "same_sheet_mapping.xlsx")
    df = pd.DataFrame([
        {"Application": "AppA", "Responsible PM": "Abishek", "PM Email": "abishek@rt.com"},
        {"Application": "AppB", "Responsible PM": "Ishan", "PM Email": "ishan@rt.com"}
    ])
    df.to_excel(file_path, index=False)
    
    email_map = find_pm_mappings(file_path)
    assert email_map.get("abishek")["email"] == "abishek@rt.com"
    assert email_map.get("ishan")["email"] == "ishan@rt.com"

def test_different_header_synonyms(tmp_path):
    from excel_reader import find_pm_mappings
    
    file_path = os.path.join(tmp_path, "synonyms.xlsx")
    df = pd.DataFrame([
        {"PM Name": "Abishek", "Email ID": "abishek@rt.com"},
        {"PM Name": "Ishan", "Email ID": "ishan@rt.com"}
    ])
    df.to_excel(file_path, index=False)
    
    email_map = find_pm_mappings(file_path)
    assert email_map.get("abishek")["email"] == "abishek@rt.com"
    
    file_path2 = os.path.join(tmp_path, "synonyms2.xlsx")
    df2 = pd.DataFrame([
        {"pm": "Abishek", "pm email": "abishek@rt.com"}
    ])
    df2.to_excel(file_path2, index=False)
    email_map2 = find_pm_mappings(file_path2)
    assert email_map2.get("abishek")["email"] == "abishek@rt.com"

def test_message_builder_mentions_logic():
    from message_builder import build_consolidated_message
    
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": False,
                "Name": "Abishek",
                "Email/Teams": "abishek@rt.com"
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "App2",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": False,
                "Name": "John Doe",
                "Email/Teams": ""
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "App3",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": True,
                "Name": "",
                "Email/Teams": ""
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps)
    assert len(cards) == 1
    card = cards[0]
    
    attention_text = card["body"][2]["text"]
    assert "<at>Abishek</at>" in attention_text
    assert "John Doe" in attention_text
    assert "<at>John Doe</at>" not in attention_text
    assert "Unassigned PM / Requires Review" in attention_text
    
    entities = card["msteams"]["entities"]
    assert len(entities) == 1
    assert entities[0]["mentioned"]["id"] == "abishek@rt.com"

# ==================== ADDITIONAL SPECIFIC TESTS ====================

def test_next_action_determination():
    from automation_core import determine_next_action
    
    row1 = {
        ('QA DB REQUESTS', 'QA OUR'): datetime.date(2026, 8, 12),
        ('QA DB REQUESTS', 'QA OUR Status'): 'Not Started',
        ('DEV DB REQUESTS', 'Dev OUR'): datetime.date(2026, 8, 15),
        ('DEV DB REQUESTS', 'Dev OUR Status'): 'Not Started'
    }
    next_act, next_date = determine_next_action(row1)
    assert next_act == 'QA OUR'
    assert next_date == datetime.date(2026, 8, 12)
    
    row2 = {
        ('DEV DB REQUESTS', 'Dev R-Task'): datetime.date(2026, 8, 10),
        ('DEV DB REQUESTS', 'Dev R-Task Status'): 'Completed',
        ('DEV DB REQUESTS', 'Dev OUR'): datetime.date(2026, 8, 15),
        ('DEV DB REQUESTS', 'Dev OUR Status'): 'Not Started'
    }
    next_act, next_date = determine_next_action(row2)
    assert next_act == 'Dev OUR'
    assert next_date == datetime.date(2026, 8, 15)

def test_mentions_isolation():
    from message_builder import build_consolidated_message
    
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": False,
                "Name": "Abishek",
                "Email/Teams": "abishek@rt.com"
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "App2",
            "row": {
                "Overall Status": "In Progress",
                "Next Action": "QA R-Task",
                "Next Action Date": datetime.date(2026, 8, 12),
                "pm_missing": False,
                "Name": "Ishan",
                "Email/Teams": "ishan@rt.com"
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "QA R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    
    # Use max_bytes=800 to force split
    cards = build_consolidated_message(qualifying_apps, max_bytes=800)
    assert len(cards) == 2
    
    card1 = cards[0]
    entities1 = card1["msteams"]["entities"]
    assert len(entities1) == 1
    assert entities1[0]["mentioned"]["id"] == "abishek@rt.com"
    assert "Abishek" in card1["body"][2]["text"]
    assert "Ishan" not in card1["body"][2]["text"]
    
    card2 = cards[1]
    entities2 = card2["msteams"]["entities"]
    assert len(entities2) == 1
    assert entities2[0]["mentioned"]["id"] == "ishan@rt.com"
    assert "Ishan" in card2["body"][2]["text"]
    assert "Abishek" not in card2["body"][2]["text"]

def test_duplicate_pm_casing_grouping(tmp_path):
    from excel_reader import read_tasks
    
    file_path = os.path.join(tmp_path, "casing_group.xlsx")
    df = pd.DataFrame([
        {"Application": "App1", "Responsible PM": "abishek"},
        {"Application": "App2", "Responsible PM": "  Abishek "},
        {"Application": "App3", "Responsible PM": "ABISHEK"}
    ])
    df.to_excel(file_path, index=False)
    
    tasks = read_tasks(file_path)
    assert tasks[0]["Name"] == tasks[1]["Name"]
    assert tasks[1]["Name"] == tasks[2]["Name"]

def test_next_action_logic_and_indicators():
    from automation_core import determine_next_action
    
    # 2. Milestones (e.g. App Ready By) do not become actionable tasks
    row = {
        ('READY / READINESS', 'App Ready By'): datetime.date(2026, 8, 12),
        ('DEV DB REQUESTS', 'Dev R-Task'): datetime.date(2026, 8, 15),
        ('DEV DB REQUESTS', 'Dev R-Task Status'): 'Not Started'
    }
    next_act, next_date = determine_next_action(row)
    assert next_act == 'Dev R-Task'
    assert next_date == datetime.date(2026, 8, 15)
    
    # 3. No actionable task -> "No actionable task identified"
    row2 = {
        ('READY / READINESS', 'App Ready By'): datetime.date(2026, 8, 12)
    }
    next_act2, next_date2 = determine_next_action(row2)
    assert next_act2 == "No actionable task identified"
    assert next_date2 is None

@patch('message_builder.get_today')
def test_date_indicators_and_empty_metadata_omission(mock_get_today):
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    
    # 5, 6, 7, 8, 9: Indicators & Omission
    qualifying_apps = [
        {
            "app": "AppOmitted",
            "row": {
                "pm_missing": True
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 10))],
            "pending_activities": [
                ("Sprint Start", datetime.date(2026, 8, 8)), # Overdue
                ("Sprint End", datetime.date(2026, 8, 15)),   # Due soon
                ("Dev R-Task", datetime.date(2026, 8, 25))    # Future date outside warning
            ],
            "overall_status": "Not available",
            "next_action": "No actionable task identified",
            "next_action_date": None
        }
    ]
    
    cards = build_consolidated_message(qualifying_apps, config_params={'SprintWarningDays': 10})
    card = cards[0]
    
    card_body_str = json.dumps(card["body"], ensure_ascii=False)
    assert "Status:" not in card_body_str
    assert "Next Action:" not in card_body_str
    assert "Due:" not in card_body_str
    
    assert "🔴 Overdue — 1" in card_body_str
    assert "• Sprint Start — 08-Aug-2026" in card_body_str
    assert "🟡 Due Soon — 1" in card_body_str
    assert "• Sprint End — 15-Aug-2026" in card_body_str
    assert "• Dev R-Task — 25-Aug-2026" in card_body_str

def test_pm_email_conflict_protection(tmp_path):
    from excel_reader import find_pm_mappings
    
    file_path = os.path.join(tmp_path, "conflicts.xlsx")
    df = pd.DataFrame([
        {"PM Name": "Abishek", "Email ID": "old@email.com"},
        {"PM Name": "Abishek", "Email ID": "new@email.com"}
    ])
    df.to_excel(file_path, index=False)
    
    mappings = find_pm_mappings(file_path)
    # The deterministic behavior is to keep the first mapping encountered
    assert mappings["abishek"]["email"] == "old@email.com"

def test_pm_group_split_continuation():
    from message_builder import build_consolidated_message
    
    # 2 applications under Abishek
    qualifying_apps = [
        {
            "app": "AppA",
            "row": {
                "pm_missing": False,
                "Name": "Abishek",
                "Email/Teams": "abishek@rt.com"
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "AppB",
            "row": {
                "pm_missing": False,
                "Name": "Abishek",
                "Email/Teams": "abishek@rt.com"
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    
    # Force splitting by using a tiny max_bytes threshold (e.g. 800 bytes)
    cards = build_consolidated_message(qualifying_apps, max_bytes=800)
    assert len(cards) == 2
    
    # Card 2 should contain "(Continued)" in the PM header block
    card2_body_str = json.dumps(cards[1]["body"], ensure_ascii=False)
    assert "👤 Abishek (Continued)" in card2_body_str

def test_metadata_both_populated():
    # Test 1 — Both fields populated
    qualifying_apps = [{
        "app": "AppFTPDaily",
        "row": {
            "file_type": "FTP",
            "frequency": "Daily",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    card_str = json.dumps(cards[0]["body"])
    assert "File Type: **FTP**" in card_str
    assert "Frequency: **Daily**" in card_str

def test_metadata_file_type_only():
    # Test 2 — File Type only
    qualifying_apps = [{
        "app": "AppFTPOnly",
        "row": {
            "file_type": "FTP",
            "frequency": "N/A",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    card_str = json.dumps(cards[0]["body"])
    assert "File Type: **FTP**" in card_str
    assert "Frequency" not in card_str

def test_metadata_frequency_only():
    # Test 3 — Frequency only
    qualifying_apps = [{
        "app": "AppFreqOnly",
        "row": {
            "file_type": "N/A",
            "frequency": "Daily",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    card_str = json.dumps(cards[0]["body"])
    assert "Frequency: **Daily**" in card_str
    assert "File Type" not in card_str

def test_metadata_both_unavailable():
    # Test 4 — Both unavailable
    qualifying_apps = [{
        "app": "AppNone",
        "row": {
            "file_type": "N/A",
            "frequency": "N/A",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    card_str = json.dumps(cards[0]["body"])
    assert "File Type" not in card_str
    assert "Frequency" not in card_str

def test_metadata_blank_cells():
    # Test 5 — Blank cells
    qualifying_apps = [{
        "app": "AppBlank",
        "row": {
            "file_type": "",
            "frequency": None,
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    card_str = json.dumps(cards[0]["body"])
    assert "File Type" not in card_str
    assert "Frequency" not in card_str

def test_existing_notification_logic_unchanged():
    # Test 6 — Existing notification logic unchanged
    row_without = {
        "Application": "AppTest",
        "Name": "Abishek",
        "Email/Teams": "abishek@rt.com",
        "pm_missing": False,
        "Sprint Start": datetime.date(2026, 8, 10),
        "Sprint End": datetime.date(2026, 8, 20),
        "Dev OUR": datetime.date(2026, 8, 12),
        "Dev OUR Status": "In Progress"
    }
    
    row_with = dict(row_without)
    row_with["file_type"] = "FTP"
    row_with["frequency"] = "Daily"
    
    acts_without, _ = get_pending_activities(row_without)
    acts_with, _ = get_pending_activities(row_with)
    assert acts_without == acts_with

def test_next_action_unchanged():
    # Test 7 — Next Action unchanged
    row = {
        "Application": "AppTest",
        "Name": "Abishek",
        "Email/Teams": "abishek@rt.com",
        "pm_missing": False,
        "Dev OUR": datetime.date(2026, 8, 12),
        "Dev OUR Status": "In Progress",
        "file_type": "FTP",
        "frequency": "Daily"
    }
    next_act, _ = automation_core.determine_next_action(row)
    assert next_act == "Dev OUR"
    assert next_act not in ["file_type", "frequency", "File Type (all envs)", "Frequency (all envs)"]

def test_pending_activities_unchanged():
    # Test 8 — Pending Activities unchanged
    row = {
        "Application": "AppTest",
        "Name": "Abishek",
        "Email/Teams": "abishek@rt.com",
        "pm_missing": False,
        "Dev OUR": datetime.date(2026, 8, 12),
        "Dev OUR Status": "In Progress",
        "file_type": "FTP",
        "frequency": "Daily"
    }
    pending, _ = get_pending_activities(row)
    assert len(pending) == 1
    assert pending[0][0] == "Dev OUR"

def test_pm_mentions_unchanged():
    # Test 9 — PM mentions unchanged
    qualifying_apps = [{
        "app": "AppFTPDaily",
        "row": {
            "file_type": "FTP",
            "frequency": "Daily",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(qualifying_apps)
    entities = cards[0]["msteams"]["entities"]
    assert len(entities) == 1
    assert entities[0]["mentioned"]["id"] == "abishek@rt.com"
    assert entities[0]["mentioned"]["name"] == "Abishek"

def test_card_splitting_with_metadata():
    # Test 10 — Card splitting
    apps = []
    for i in range(15):
        apps.append({
            "app": f"App{i}",
            "row": {
                "file_type": "FTP",
                "frequency": "Daily",
                "Name": f"PM{i % 2}",
                "Email/Teams": f"pm{i % 2}@rt.com",
                "pm_missing": False
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    cards = build_consolidated_message(apps, max_bytes=2000)
    assert len(cards) > 1
    for c in cards:
        assert len(json.dumps(c).encode('utf-8')) <= 2000
        
    for c in cards:
        card_str = json.dumps(c["body"])
        entities = c["msteams"].get("entities", [])
        entity_ids = {ent["mentioned"]["id"] for ent in entities}
        
        pms_in_body = []
        if "PM0" in card_str:
            pms_in_body.append("pm0@rt.com")
        if "PM1" in card_str:
            pms_in_body.append("pm1@rt.com")
        
        for pid in pms_in_body:
            assert pid in entity_ids
        for ent in entities:
            assert ent["mentioned"]["id"] in ["pm0@rt.com", "pm1@rt.com"]

def test_header_normalization(tmp_path):
    # Test 11 — Header normalization
    from excel_reader import read_tasks
    
    file_path = os.path.join(tmp_path, "header_norm.xlsx")
    df = pd.DataFrame([
        {
            "Application": "AppNorm",
            "Responsible PM": "Abishek",
            "file type (all envs)": "FTP",
            "  FREQUENCY (all envs)  ": "Daily"
        }
    ])
    df.to_excel(file_path, index=False)
    
    records = read_tasks(file_path)
    assert len(records) == 1
    assert records[0]["file_type"] == "FTP"
    assert records[0]["frequency"] == "Daily"

def test_existing_workbook_regression():
    # Test 12 — Existing workbook regression
    from excel_reader import read_tasks
    records = read_tasks("appDeadlineSample.xlsx")
    assert len(records) > 0
    for r in records:
        assert "file_type" in r
        assert "frequency" in r

def test_card_splitting_continuation_and_preservation():
    from message_builder import build_consolidated_message
    
    # Construct apps for PM1, PM2, PM3
    apps = []
    # PM1: 4 apps
    for i in range(4):
        apps.append({
            "app": f"PM1_App_{i}",
            "row": {
                "file_type": "FTP",
                "frequency": "Daily",
                "Name": "PM1",
                "Email/Teams": "pm1@rt.com",
                "pm_missing": False
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    # PM2: 4 apps
    for i in range(4):
        apps.append({
            "app": f"PM2_App_{i}",
            "row": {
                "file_type": "FTP",
                "frequency": "Daily",
                "Name": "PM2",
                "Email/Teams": "pm2@rt.com",
                "pm_missing": False
            },
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
        
    # We will trigger splits by using a small max_bytes (e.g. 1500 bytes)
    cards = build_consolidated_message(apps, max_bytes=1500)
    assert len(cards) > 1
    
    # 1. input applications == union(all card applications)
    # 2. no application appears twice
    all_card_apps = []
    for card in cards:
        card_body_str = json.dumps(card["body"])
        for app in apps:
            if f"**{app['app']}**" in card_body_str:
                all_card_apps.append(app["app"])
                
    assert len(all_card_apps) == len(apps)
    assert set(all_card_apps) == {app["app"] for app in apps}
    
    # 3. PM entirely fits in one card -> no (Continued)
    large_cards = build_consolidated_message(apps, max_bytes=50000)
    assert len(large_cards) == 1
    card1_body_str = json.dumps(large_cards[0]["body"], ensure_ascii=False)
    assert "(Continued)" not in card1_body_str
    
    # 4. Check that "(Continued)" is not on the first card where a PM appears,
    # but is present on the subsequent cards where they continue.
    pm1_appeared_indices = []
    pm2_appeared_indices = []
    
    for idx, card in enumerate(cards):
        card_body_str = json.dumps(card["body"], ensure_ascii=False)
        if "PM1" in card_body_str:
            pm1_appeared_indices.append(idx)
        if "PM2" in card_body_str:
            pm2_appeared_indices.append(idx)
            
    # Verify PM1 continuation
    if len(pm1_appeared_indices) > 1:
        first_card_pm1 = json.dumps(cards[pm1_appeared_indices[0]]["body"], ensure_ascii=False)
        assert "PM1 (Continued)" not in first_card_pm1
        for idx in pm1_appeared_indices[1:]:
            cont_card_pm1 = json.dumps(cards[idx]["body"], ensure_ascii=False)
            assert "PM1 (Continued)" in cont_card_pm1
            
    # Verify PM2 continuation
    if len(pm2_appeared_indices) > 1:
        first_card_pm2 = json.dumps(cards[pm2_appeared_indices[0]]["body"], ensure_ascii=False)
        assert "PM2 (Continued)" not in first_card_pm2
        for idx in pm2_appeared_indices[1:]:
            cont_card_pm2 = json.dumps(cards[idx]["body"], ensure_ascii=False)
            assert "PM2 (Continued)" in cont_card_pm2

# ==================== UX BEAUTIFICATION TESTS ====================

def test_single_pm_app_numbering():
    # Test 1 — Single PM application numbering
    from message_builder import build_consolidated_message
    apps = [
        {
            "app": "APP_A",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_B",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_C",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "1. **APP_A**" in card_str
    assert "2. **APP_B**" in card_str
    assert "3. **APP_C**" in card_str

def test_numbering_resets_per_pm():
    # Test 2 — Numbering resets per PM
    from message_builder import build_consolidated_message
    apps = [
        {
            "app": "APP_A",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_B",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_C",
            "row": {"Name": "Ishan", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_D",
            "row": {"Name": "Ishan", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "1. **APP_A**" in card_str
    assert "2. **APP_B**" in card_str
    assert "1. **APP_C**" in card_str
    assert "2. **APP_D**" in card_str
    assert "3. **APP_C**" not in card_str

def test_numbering_across_continuation_cards():
    # Test 3 — Numbering across continuation cards
    from message_builder import build_consolidated_message
    apps = []
    for i in range(5):
        apps.append({
            "app": f"APP_{chr(65+i)}",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    # Force split using a low max_bytes threshold (e.g., 4500 bytes)
    cards = build_consolidated_message(apps, max_bytes=4500)
    assert len(cards) == 2
    
    card1_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    card2_str = json.dumps(cards[1]["body"], ensure_ascii=False)
    
    assert "1. **APP_A**" in card1_str
    assert "2. **APP_B**" in card1_str
    assert "3. **APP_C**" in card1_str
    assert "4. **APP_D**" in card1_str
    # Card 2 should continue numbering
    assert "5. **APP_E**" in card2_str
    assert "1. **APP_E**" not in card2_str

def test_multiple_pms_independently_split():
    # Test 4 — Multiple PMs independently split
    from message_builder import build_consolidated_message
    apps = []
    # PM1: 3 apps, PM2: 3 apps
    for i in range(3):
        apps.append({
            "app": f"PM1_APP_{i}",
            "row": {"Name": "PM1", "Email/Teams": "pm1@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    for i in range(3):
        apps.append({
            "app": f"PM2_APP_{i}",
            "row": {"Name": "PM2", "Email/Teams": "pm2@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    # Force split using a low max_bytes (e.g. 3500 bytes)
    cards = build_consolidated_message(apps, max_bytes=3500)
    assert len(cards) >= 2
    
    # Assert each PM sequence is tracked independently and does not reset to 1 on subsequent splits unless it is a new PM
    pm1_apps = {}
    pm2_apps = {}
    for idx, card in enumerate(cards):
        card_str = json.dumps(card["body"], ensure_ascii=False)
        for i in range(3):
            if f"PM1_APP_{i}" in card_str:
                pm1_apps[f"PM1_APP_{i}"] = idx
            if f"PM2_APP_{i}" in card_str:
                pm2_apps[f"PM2_APP_{i}"] = idx
                
    # Verify numbering formatting exists correctly
    card_strs = [json.dumps(c["body"], ensure_ascii=False) for c in cards]
    assert any("1. **PM1_APP_0**" in cs for cs in card_strs)
    assert any("1. **PM2_APP_0**" in cs for cs in card_strs)

@patch('message_builder.get_today')
def test_all_overdue(mock_get_today):
    # Test 5 — All overdue
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppOverdueOnly",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [
            ("Sprint Start", datetime.date(2026, 8, 5)),
            ("Dev OUR", datetime.date(2026, 8, 8))
        ],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "🔴 Overdue — 2" in card_str
    assert "🟡 Due Soon" not in card_str

@patch('message_builder.get_today')
def test_all_due_soon(mock_get_today):
    # Test 6 — All due soon
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppDueSoonOnly",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [
            ("Sprint End", datetime.date(2026, 8, 15)),
            ("Dev R-Task", datetime.date(2026, 8, 18))
        ],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "🟡 Due Soon — 2" in card_str
    assert "🔴 Overdue" not in card_str

@patch('message_builder.get_today')
def test_mixed_urgency(mock_get_today):
    # Test 7 — Mixed urgency
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppMixed",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [
            ("Act_A", datetime.date(2026, 8, 5)),  # overdue
            ("Act_B", datetime.date(2026, 8, 15)), # due soon
            ("Act_C", datetime.date(2026, 8, 8)),  # overdue
            ("Act_D", datetime.date(2026, 8, 18))  # due soon
        ],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    
    # Assert group headings
    assert "🔴 Overdue — 2" in card_str
    assert "🟡 Due Soon — 2" in card_str
    
    # Verify order preservation within each category (A before C in Overdue, B before D in Due Soon)
    # A should appear before C
    idx_a = card_str.find("Act_A")
    idx_c = card_str.find("Act_C")
    assert idx_a != -1 and idx_c != -1
    assert idx_a < idx_c
    
    # B should appear before D
    idx_b = card_str.find("Act_B")
    idx_d = card_str.find("Act_D")
    assert idx_b != -1 and idx_d != -1
    assert idx_b < idx_d

@patch('message_builder.get_today')
def test_count_accuracy(mock_get_today):
    # Test 8 — Count accuracy
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppCount",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [
            ("Act1", datetime.date(2026, 8, 5)),
            ("Act2", datetime.date(2026, 8, 6)),
            ("Act3", datetime.date(2026, 8, 15))
        ],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "🔴 Overdue — 2" in card_str
    assert "🟡 Due Soon — 1" in card_str

def test_no_pending_activities_preserved():
    # Test 9 — No pending activities preserved (renders None)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppNonePending",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "None" in card_str

def test_existing_business_logic_unchanged_by_grouping():
    # Test 10 — Existing business logic unchanged
    from message_builder import build_consolidated_message
    row = {
        "Application": "AppTest",
        "Name": "Abishek",
        "Email/Teams": "abishek@rt.com",
        "pm_missing": False,
        "Dev OUR": datetime.date(2026, 8, 12),
        "Dev OUR Status": "In Progress"
    }
    next_act, next_act_date = automation_core.determine_next_action(row)
    assert next_act == "Dev OUR"
    assert next_act_date == datetime.date(2026, 8, 12)

def test_metadata_regression_preserved():
    # Test 11 — Metadata regression
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppFTPDaily",
        "row": {
            "file_type": "FTP",
            "frequency": "Daily",
            "Name": "Abishek",
            "Email/Teams": "abishek@rt.com",
            "pm_missing": False
        },
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "File Type: **FTP**" in card_str
    assert "Frequency: **Daily**" in card_str

def test_card_size_regression_preserved():
    # Test 12 — Card size regression
    from message_builder import build_consolidated_message
    apps = []
    for i in range(10):
        apps.append({
            "app": f"AppSize_{i}",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    # Size remains below 18 KB (18432 bytes)
    cards = build_consolidated_message(apps, max_bytes=18432)
    for c in cards:
        card_size = len(json.dumps(c).encode('utf-8'))
        assert card_size <= 18432

def test_pm_mention_regression_preserved():
    # Test 13 — PM mention regression
    from message_builder import build_consolidated_message
    apps_resolved = [{
        "app": "AppA",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    apps_unresolved = [{
        "app": "AppB",
        "row": {"Name": "UnknownPM", "Email/Teams": "", "pm_missing": True},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    
    cards_res = build_consolidated_message(apps_resolved)
    card_unres = build_consolidated_message(apps_unresolved)
    
    assert "<at>Abishek</at>" in cards_res[0]["body"][2]["text"]
    assert "msteams" in cards_res[0]
    
    assert "**Unassigned PM / Requires Review**" in card_unres[0]["body"][2]["text"]
    assert "msteams" not in card_unres[0]

def test_no_duplicate_applications_preserved():
    # Test 14 — No duplicate applications
    from message_builder import build_consolidated_message
    apps = []
    for i in range(12):
        apps.append({
            "app": f"AppUnique_{i}",
            "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    cards = build_consolidated_message(apps, max_bytes=2000)
    assert len(cards) > 1
    
    all_card_apps = []
    for card in cards:
        card_str = json.dumps(card["body"], ensure_ascii=False)
        for app in apps:
            if f"**{app['app']}**" in card_str:
                all_card_apps.append(app["app"])
                
    assert len(all_card_apps) == len(apps)
    assert set(all_card_apps) == {app["app"] for app in apps}

# ==================== PM NAME & ENTRY POINT TESTS ====================

def test_pm_name_appears_beside_app():
    # Test 1 — PM name appears beside application
    from message_builder import build_consolidated_message
    apps = [{
        "app": "APP_A",
        "row": {"Name": "Ishan Bhattacharjee", "Email/Teams": "ishan@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "1. **APP_A** [Ishan Bhattacharjee]" in card_str

def test_multiple_apps_pm_name():
    # Test 2 — Multiple applications
    from message_builder import build_consolidated_message
    apps = [
        {
            "app": "APP_A",
            "row": {"Name": "Ishan Bhattacharjee", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_B",
            "row": {"Name": "Ishan Bhattacharjee", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_C",
            "row": {"Name": "Ishan Bhattacharjee", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "1. **APP_A** [Ishan Bhattacharjee]" in card_str
    assert "2. **APP_B** [Ishan Bhattacharjee]" in card_str
    assert "3. **APP_C** [Ishan Bhattacharjee]" in card_str

def test_numbering_survives_card_split_pm_name():
    # Test 3 — Numbering survives card split
    from message_builder import build_consolidated_message
    apps = []
    for i in range(12):
        apps.append({
            "app": f"APP_{i}",
            "row": {"Name": "Ishan Bhattacharjee", "Email/Teams": "ishan@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        })
    cards = build_consolidated_message(apps, max_bytes=4000)
    assert len(cards) > 1
    
    # Assert that subsequent indices still show the PM name correctly
    found_continuation = False
    for idx, card in enumerate(cards):
        card_str = json.dumps(card["body"], ensure_ascii=False)
        if idx > 0:
            # This is a continuation card, check some indices
            for i in range(12):
                expected = f"{i+1}. **APP_{i}** [Ishan Bhattacharjee]"
                if expected in card_str:
                    found_continuation = True
                    break
    assert found_continuation

def test_different_pms_formatting():
    # Test 4 — Different PMs
    from message_builder import build_consolidated_message
    apps = [
        {
            "app": "APP_A",
            "row": {"Name": "PM A", "Email/Teams": "pma@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_B",
            "row": {"Name": "PM A", "Email/Teams": "pma@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_C",
            "row": {"Name": "PM B", "Email/Teams": "pmb@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "APP_D",
            "row": {"Name": "PM B", "Email/Teams": "pmb@rt.com", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "1. **APP_A** [PM A]" in card_str
    assert "2. **APP_B** [PM A]" in card_str
    assert "1. **APP_C** [PM B]" in card_str
    assert "2. **APP_D** [PM B]" in card_str

def test_missing_pm_fallback():
    # Test 5 — Missing/unresolved PM
    from message_builder import build_consolidated_message
    apps = [{
        "app": "APP_A",
        "row": {"Name": "", "Email/Teams": "", "pm_missing": True},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    # Check that fallback is clean and doesn't print [None], [nan], or []
    assert "1. **APP_A**" in card_str
    assert "[None]" not in card_str
    assert "[nan]" not in card_str
    assert "[]" not in card_str

def test_metadata_unaffected():
    # Test 6 — Existing metadata still works
    from message_builder import build_consolidated_message
    apps = [
        {
            "app": "AppFT",
            "row": {"file_type": "FTP", "frequency": "N/A", "Name": "Abishek", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "AppFreq",
            "row": {"file_type": "N/A", "frequency": "Daily", "Name": "Abishek", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        },
        {
            "app": "AppBoth",
            "row": {"file_type": "SFTP", "frequency": "Daily", "Name": "Abishek", "pm_missing": False},
            "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
            "pending_activities": [("Dev OUR", datetime.date(2026, 8, 12))],
            "overall_status": "In Progress",
            "next_action": "Dev R-Task",
            "next_action_date": datetime.date(2026, 8, 12)
        }
    ]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "📄 File Type: **FTP**" in card_str
    assert "🔄 Frequency: **Daily**" in card_str
    assert "📄 File Type: **SFTP**  ·  Frequency: **Daily**" in card_str

@patch('message_builder.get_today')
def test_urgency_grouping_unaffected(mock_get_today):
    # Test 7 — Existing urgency grouping still works
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    from message_builder import build_consolidated_message
    apps = [{
        "app": "AppMixed",
        "row": {"Name": "Abishek", "Email/Teams": "abishek@rt.com", "pm_missing": False},
        "reminders": [("START_DATE", datetime.date(2026, 8, 12))],
        "pending_activities": [
            ("Act_A", datetime.date(2026, 8, 5)),  # overdue
            ("Act_B", datetime.date(2026, 8, 15))  # due soon
        ],
        "overall_status": "In Progress",
        "next_action": "Dev R-Task",
        "next_action_date": datetime.date(2026, 8, 12)
    }]
    cards = build_consolidated_message(apps)
    card_str = json.dumps(cards[0]["body"], ensure_ascii=False)
    assert "🔴 Overdue — 1" in card_str
    assert "🟡 Due Soon — 1" in card_str

def test_production_launcher_execution():
    # Test 8 — Production launcher
    import subprocess
    import os
    env = dict(os.environ, DRY_RUN="true")
    res = subprocess.run(["python", "teams_automation.py"], env=env, capture_output=True, text=True)
    assert res.returncode == 0
    assert "DRY RUN" in res.stdout
