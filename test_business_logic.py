import pytest
from message_builder import build_consolidated_message, get_status_priority
import main
from unittest.mock import patch
import datetime

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
                "Name": "Abishek"
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
                "Name": "Abishek"
            },
            "reminders": [("END_DATE", datetime.date(2026, 8, 15))],
            "pending_activities": [("QA OUR", datetime.date(2026, 8, 17))]
        }
    ]
    
    card = build_consolidated_message(qualifying_apps)
    
    assert card["type"] == "AdaptiveCard"
    
    # Check severity
    assert card["body"][0]["text"] == "URGENT & IMPORTANT | ACCESS REQUESTS"
    
    # Check mentions aggregation
    attention_text = card["body"][1]["text"]
    assert "<at>Abishek</at>" in attention_text
    
    entities = card["msteams"]["entities"]
    assert len(entities) == 1
    assert entities[0]["mentioned"]["id"] == "amuralidharan@randomtrees.com"
    
    # We should have:
    # 0: Title
    # 1: Attention required
    # 2: The following applications...
    # 3: Mention for Abishek
    # 4: Container for App1
    # 5: Container for App2
    assert len(card["body"]) == 6
    
    # Check PM header block
    assert "<at>Abishek</at>" in card["body"][3]["text"]
    
    app1_container = str(card["body"][4])
    app2_container = str(card["body"][5])
    
    # App 1 properties
    assert "App1" in app1_container
    assert "Dev OUR — 12-Aug-2026" in app1_container
    assert "QA R-Task — 15-Aug-2026" in app1_container
    
    # App 2 properties (Activities should not merge)
    assert "App2" in app2_container
    assert "QA OUR — 17-Aug-2026" in app2_container
    assert "Dev OUR" not in app2_container

def test_deduplication_mention():
    qualifying_apps = [
        {
            "app": "App1",
            "row": {
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek"
            },
            "reminders": [("START_DATE", None)],
            "pending_activities": []
        },
        {
            "app": "App2",
            "row": {
                "Email/Teams": "amuralidharan@randomtrees.com",
                "Name": "Abishek"
            },
            "reminders": [("START_DATE", None)],
            "pending_activities": []
        }
    ]
    
    card = build_consolidated_message(qualifying_apps)
    entities = card["msteams"]["entities"]
    assert len(entities) == 1 # Duplicate removed

@patch('main.get_today')
def test_pending_activities_10_day_filter(mock_get_today):
    # Setup today's date
    mock_get_today.return_value = datetime.date(2026, 8, 10)
    
    # Base row
    base_row = {
        'Dev R-Task Status': 'Not Started',
        'Dev OUR Status': 'Not Started',
        'QA R-Task Status': 'Not Started',
        'Prod R-Task Status': 'Not Started',
        'Input Audit Status': 'Not Started',
        'Ready Status': 'Not Started'
    }

    # Case 1: Only DEV is due within 10 days
    row1 = dict(base_row)
    row1['Dev R-Task'] = datetime.date(2026, 8, 15) # 5 days
    row1['QA R-Task'] = datetime.date(2026, 8, 25) # 15 days
    pending1, _ = main.get_pending_activities(row1)
    act_names = [a[0] for a in pending1]
    assert act_names == ['Dev R-Task']

    # Case 2: DEV and QA due within 10 days, PROD is not
    row2 = dict(base_row)
    row2['Dev R-Task'] = datetime.date(2026, 8, 12) # 2 days
    row2['QA R-Task'] = datetime.date(2026, 8, 19) # 9 days
    row2['Prod R-Task'] = datetime.date(2026, 8, 25) # 15 days
    pending2, _ = main.get_pending_activities(row2)
    act_names2 = [a[0] for a in pending2]
    assert set(act_names2) == {'Dev R-Task', 'QA R-Task'}

    # Case 3: Input Audit and Ready have different due dates and are filtered independently
    row3 = dict(base_row)
    row3['Input Audit'] = datetime.date(2026, 8, 18) # 8 days (include)
    row3['Ready'] = datetime.date(2026, 8, 28) # 18 days (exclude)
    pending3, _ = main.get_pending_activities(row3)
    act_names3 = [a[0] for a in pending3]
    assert act_names3 == ['Input Audit']

    # Case 4: No activities are due within 10 days
    row4 = dict(base_row)
    row4['Dev R-Task'] = datetime.date(2026, 8, 21) # 11 days
    row4['QA R-Task'] = datetime.date(2026, 8, 30) # 20 days
    pending4, _ = main.get_pending_activities(row4)
    assert pending4 == []

    # Case 5: Activities exactly 10 days away are included
    row5 = dict(base_row)
    row5['Dev R-Task'] = datetime.date(2026, 8, 20) # exactly 10 days
    pending5, _ = main.get_pending_activities(row5)
    act_names5 = [a[0] for a in pending5]
    assert act_names5 == ['Dev R-Task']

    # Case 6: Date-bearing columns without a status column are dynamically identified
    row6 = dict(base_row)
    row6['Sprint Start'] = datetime.date(2026, 8, 12)
    row6['Sprint End'] = datetime.date(2026, 8, 25)
    pending6, _ = main.get_pending_activities(row6)
    act_names6 = [a[0] for a in pending6]
    assert 'Sprint Start' in act_names6
    assert 'Sprint End' not in act_names6
