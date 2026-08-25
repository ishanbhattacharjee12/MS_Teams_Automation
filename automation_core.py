import config
import json
import os
import re
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
    
    MILESTONE_COLUMNS = config.PURE_MILESTONES
    exclude_cols = config.METADATA_COLUMNS
    
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

def determine_next_action(row, config_params=None):
    ACTIONABLE_TASKS = config.ACTIONABLE_TASKS
    
    keys = list(row.keys())
    has_tuples = any(isinstance(k, tuple) for k in keys)
    
    candidates = []
    
    if has_tuples:
        tuple_keys = [k for k in keys if isinstance(k, tuple)]
        for k in tuple_keys:
            group, col = k
            col_str = str(col).strip()
            
            if col_str not in ACTIONABLE_TASKS:
                continue
                
            act_date = parse_date(row.get(k))
            if act_date is not None:
                group_cols = [tk[1] for tk in tuple_keys if tk[0] == group]
                matched_status_col = find_status_column(col_str, group_cols)
                if matched_status_col is not None:
                    status_tuple_key = (group, matched_status_col)
                    st = str(row.get(status_tuple_key, '')).strip().replace('\xa0', '')
                    if st.lower() != 'completed':
                        candidates.append((col_str, act_date))
    else:
        for col in keys:
            col_str = str(col).strip()
            if col_str not in ACTIONABLE_TASKS:
                continue
                
            act_date = parse_date(row.get(col))
            if act_date is not None:
                matched_status_col = find_status_column(col_str, keys)
                if matched_status_col is not None:
                    st = str(row.get(matched_status_col, '')).strip().replace('\xa0', '')
                    if st.lower() != 'completed':
                        candidates.append((col_str, act_date))
                        
    if candidates:
        # Sort candidates by date
        candidates.sort(key=lambda x: x[1])
        return candidates[0][0], candidates[0][1]
        
    return "No actionable task identified", None

def main():
    tasks = read_tasks(config.EXCEL_FILE, config.SHEET_NAME)
    print(f"[INFO] Workbook loaded: {len(tasks)} applications")
    
    config_params = read_config_params(config.EXCEL_FILE)
    
    # Calculate resolved unique PMs count
    unique_pms = set()
    for r in tasks:
        pm_name = str(r.get('Name', '')).strip().lower()
        if pm_name and not r.get('pm_missing'):
            unique_pms.add(pm_name)
    print(f"[INFO] PM identities resolved: {len(unique_pms)}")
    
    if not tasks:
        print("[INFO] No tasks found.")
        return
        
    qualifying_apps = []
    warning_days = config_params.get('SprintWarningDays', config.REMINDER_DAYS_BEFORE)
    
    for row in tasks:
        app = row.get('APP', 'Unknown')
        if not app:
            continue
            
        pending_acts, all_statuses = get_pending_activities(row, config_params)
        
        if not pending_acts:
            continue
            
        start_date = parse_date(row.get('Start Date'))
        end_date = parse_date(row.get('End Date'))
        
        reminder_types = []
        if start_date and is_in_reminder_window(start_date, warning_days):
            reminder_types.append(('START_DATE', start_date))
        if end_date and is_in_reminder_window(end_date, warning_days):
            reminder_types.append(('END_DATE', end_date))
            
        if not reminder_types:
            continue
            
        valid_reminders = []
        for r_type, r_date in reminder_types:
            if config.ENABLE_DEDUPLICATION and not should_send_notification(
                app, r_type, str(r_date), all_statuses, "", "", config.ENABLE_DEDUPLICATION
            ):
                pass
            else:
                valid_reminders.append((r_type, r_date))
                
        if valid_reminders:
            next_act, next_act_date = determine_next_action(row, config_params)
            
            overall_status = str(row.get('Overall Status', '')).strip()
            if not overall_status:
                overall_status = "Not available"
                
            qualifying_apps.append({
                "app": app,
                "row": row,
                "reminders": valid_reminders,
                "pending_activities": pending_acts,
                "all_statuses": all_statuses,
                "overall_status": overall_status,
                "next_action": next_act,
                "next_action_date": next_act_date
            })

    # Stats for logging
    num_unassigned = sum(1 for qa in qualifying_apps if qa["row"].get("pm_missing"))
    num_unresolved_emails = sum(1 for qa in qualifying_apps if not qa["row"].get("pm_missing") and not qa["row"].get("Email/Teams"))
    
    print(f"[INFO] Applications requiring notification: {len(qualifying_apps)}")
    print(f"[INFO] Unassigned PM applications: {num_unassigned}")
    print(f"[INFO] Applications with unresolved PM emails: {num_unresolved_emails}")

    if not qualifying_apps:
        print("[INFO] No applications qualify for notification at this time. Processing complete.")
        return

    cards = build_consolidated_message(qualifying_apps, config_params=config_params)
    print(f"[INFO] Adaptive Cards generated: {len(cards)}")
    
    for idx, c in enumerate(cards):
        size_bytes = len(json.dumps(c).encode('utf-8'))
        print(f"[INFO] Card {idx+1}: {size_bytes} bytes")
        
    if config.DRY_RUN:
        print("\n--- DRY RUN: WOULD SEND CONSOLIDATED ADAPTIVE CARDS ---")
        for idx, c in enumerate(cards):
            print(f"--- CARD {idx+1} ---")
            print(json.dumps({"card": c}, indent=2))
        print("------------------------------------------------------\n")
    else:
        all_success = True
        for idx, c in enumerate(cards):
            print(f"[INFO] Sending card {idx+1}/{len(cards)}...")
            success, status_code, err_msg = send_notification(c)
            if success:
                pass
            else:
                print(f"[ERROR] Teams notification failed for card {idx+1}/{len(cards)}")
                if status_code is not None:
                    print(f"[ERROR] HTTP {status_code} - {err_msg}")
                else:
                    print(f"[ERROR] Connection Error - {err_msg}")
                
                size_bytes = len(json.dumps(c).encode('utf-8'))
                if size_bytes < 18000:
                    print(f"[WARNING] Card size ({size_bytes} bytes) is below the safety threshold. The failure is likely due to downstream Power Automate issues (invalid credentials, disabled flow, or connector limit).")
                all_success = False
                
        if all_success:
            if config.ENABLE_DEDUPLICATION:
                for qa in qualifying_apps:
                    for r_type, r_date in qa["reminders"]:
                        update_state(qa["app"], r_type, str(r_date), qa["all_statuses"], "", "")
            print("[INFO] Processing complete")
        else:
            print("[ERROR] One or more Teams notifications failed.")

if __name__ == "__main__":
    main()
