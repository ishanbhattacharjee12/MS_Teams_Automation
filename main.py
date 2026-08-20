import config
import json
import os
from excel_reader import read_tasks, read_config_params
from message_builder import build_consolidated_message
from teams_webhook import send_notification
from notification_state import should_send_notification, update_state
from date_checker import parse_date, is_in_reminder_window, get_today

def find_status_column(date_col_name, group_columns):
    d_name = str(date_col_name).strip()
    d_name_lower = d_name.lower()
    
    status_cols = [c for c in group_columns if "status" in str(c).lower()]
    
    # 1. Direct Substring Check
    for sc in status_cols:
        if d_name_lower in str(sc).lower():
            return sc
            
    # 2. Known Abbreviation / Mapping Pattern
    abbrev_map = {
        "service request": ["sr status"],
        "code review invite": ["cr invite status", "cr status"]
    }
    for pattern, status_suffixes in abbrev_map.items():
        if pattern in d_name_lower:
            prefix = d_name_lower.split(pattern)[0].strip()
            for sc in status_cols:
                sc_lower = str(sc).lower()
                if prefix in sc_lower and any(suffix in sc_lower for suffix in status_suffixes):
                    return sc
                    
    # 3. Base Name Check
    d_base = d_name_lower.replace("date", "").replace("invite", "").strip()
    for sc in status_cols:
        sc_lower = str(sc).lower()
        sc_base = sc_lower.replace("status", "").strip()
        if d_base and sc_base and (d_base in sc_base or sc_base in d_base):
            return sc
            
    return None

def get_pending_activities(row, config_params=None):
    pending = []
    statuses = []
    
    today = get_today()
    warning_days = config_params.get('SprintWarningDays', 10) if config_params else 10
    
    MILESTONE_COLUMNS = {
        'Sprint Start', 'Sprint End', 'Release Date',
        'Effective Sprint Start', 'Effective Sprint End', 'Effective Release Date',
        'App Ready By'
    }
    
    exclude_cols = {
        'Next Action Date', 'Start Date', 'End Date', 'Application', 'Responsible PM',
        'Sprint', 'Overall Status', 'Next Action', 'Days Overdue', 'Validation / Notes',
        'Email/Teams', 'Name', 'APP'
    }
    
    keys = list(row.keys())
    has_tuples = any(isinstance(k, tuple) for k in keys)
    
    if has_tuples:
        tuple_keys = [k for k in keys if isinstance(k, tuple)]
        for idx, k in enumerate(tuple_keys):
            group, col = k
            col_str = str(col).strip()
            
            if col_str in exclude_cols or col_str == "":
                continue
            if any(term in col_str.lower() for term in ["raised on", "file type", "iam", "ownership identified", "frequency", "unnamed:"]):
                continue
                
            act_date = parse_date(row.get(k))
            if act_date is not None:
                st = ""
                if col_str in MILESTONE_COLUMNS:
                    pass
                else:
                    group_cols = [tk[1] for tk in tuple_keys if tk[0] == group]
                    matched_status_col = find_status_column(col_str, group_cols)
                    if matched_status_col is None:
                        continue
                    status_tuple_key = (group, matched_status_col)
                    st = str(row.get(status_tuple_key, '')).strip().replace('\xa0', '')
                    statuses.append(st)
                    
                if st.lower() != 'completed':
                    days_away = (act_date - today).days
                    if days_away <= warning_days:
                        pending.append((col_str, act_date))
    else:
        for i, col in enumerate(keys):
            col_str = str(col).strip()
            if col_str in exclude_cols or col_str == "":
                continue
            if any(term in col_str.lower() for term in ["raised on", "file type", "iam", "ownership identified", "frequency", "unnamed:"]):
                continue
                
            act_date = parse_date(row.get(col))
            if act_date is not None:
                st = ""
                if col_str in MILESTONE_COLUMNS:
                    pass
                else:
                    matched_status_col = find_status_column(col_str, keys)
                    if matched_status_col is None:
                        continue
                    st = str(row.get(matched_status_col, '')).strip().replace('\xa0', '')
                    statuses.append(st)
                    
                if st.lower() != 'completed':
                    days_away = (act_date - today).days
                    if days_away <= warning_days:
                        pending.append((col_str, act_date))
                        
    return pending, "_".join(statuses)

def main():
    print("[INFO] Loading Excel file...")
    tasks = read_tasks(config.EXCEL_FILE, config.SHEET_NAME)
    config_params = read_config_params(config.EXCEL_FILE)
    print(f"[INFO] Loaded Config parameters: {config_params}")
    
    if not tasks:
        print("[INFO] No tasks found.")
        return
        
    print(f"[INFO] Found {len(tasks)} applications.")
    
    qualifying_apps = []
    warning_days = config_params.get('SprintWarningDays', config.REMINDER_DAYS_BEFORE)
    
    for row in tasks:
        app = row.get('APP', 'Unknown')
        
        pending_acts, all_statuses = get_pending_activities(row, config_params)
        
        if not pending_acts:
            print(f"[INFO] {app}: all requests completed. No notification required.")
            continue
            
        print(f"[INFO] {app}: pending request detected. Checking dates...")
        
        start_date = parse_date(row.get('Start Date'))
        end_date = parse_date(row.get('End Date'))
        
        reminder_types = []
        if start_date and is_in_reminder_window(start_date, warning_days):
            reminder_types.append(('START_DATE', start_date))
        if end_date and is_in_reminder_window(end_date, warning_days):
            reminder_types.append(('END_DATE', end_date))
            
        if not reminder_types:
            print(f"[INFO] {app}: neither Start Date nor End Date are in the reminder window.")
            continue
            
        valid_reminders = []
        for r_type, r_date in reminder_types:
            if config.ENABLE_DEDUPLICATION and not should_send_notification(
                app, r_type, str(r_date), all_statuses, "", "", config.ENABLE_DEDUPLICATION
            ):
                print(f"[INFO] {app}: {r_type} duplicate notification prevented by deduplication logic.")
            else:
                valid_reminders.append((r_type, r_date))
                
        if valid_reminders:
            qualifying_apps.append({
                "app": app,
                "row": row,
                "reminders": valid_reminders,
                "pending_activities": pending_acts,
                "all_statuses": all_statuses
            })
            print(f"[INFO] {app}: qualified for notification.")

    if not qualifying_apps:
        print("[INFO] No applications qualify for notification at this time. Processing complete.")
        return

    cards = build_consolidated_message(qualifying_apps)
    
    print(f"[INFO] Generated {len(cards)} Adaptive Card(s).")
    for idx, c in enumerate(cards):
        size_bytes = len(json.dumps(c).encode('utf-8'))
        print(f"[INFO] Card {idx+1} size: {size_bytes / 1024.0:.2f} KB ({size_bytes} bytes)")
        
    if config.DRY_RUN:
        print("\n--- DRY RUN: WOULD SEND CONSOLIDATED ADAPTIVE CARDS ---")
        for idx, c in enumerate(cards):
            print(f"--- CARD {idx+1} ---")
            print(json.dumps({"card": c}, indent=2))
        print("------------------------------------------------------\n")
    else:
        print("[INFO] Sending consolidated Teams notification(s)...")
        all_success = True
        for idx, c in enumerate(cards):
            success = send_notification(c)
            if success:
                print(f"[INFO] Card {idx+1} sent successfully.")
            else:
                print(f"[ERROR] Failed to send Card {idx+1}.")
                all_success = False
                
        if all_success:
            print("[INFO] All Teams notifications sent successfully.")
            if config.ENABLE_DEDUPLICATION:
                for qa in qualifying_apps:
                    for r_type, r_date in qa["reminders"]:
                        update_state(qa["app"], r_type, str(r_date), qa["all_statuses"], "", "")
        else:
            print("[ERROR] One or more Teams notifications failed.")

    print("[INFO] Processing complete.")

if __name__ == "__main__":
    main()
