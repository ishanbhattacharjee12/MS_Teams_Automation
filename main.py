import config
import json
from excel_reader import read_tasks
from message_builder import build_consolidated_message
from teams_webhook import send_notification
from notification_state import should_send_notification, update_state
from date_checker import parse_date, is_in_reminder_window, get_today

def get_pending_activities(row):
    pending = []
    statuses = []
    
    today = get_today()
    
    # Exclude internal/control dates
    exclude_cols = {'Next Action Date', 'Start Date', 'End Date'}
    
    keys = list(row.keys())
    for i, col in enumerate(keys):
        if col in exclude_cols:
            continue
            
        act_date = parse_date(row.get(col))
        if act_date is not None:
            # We found a date-bearing column!
            st = ""
            if i + 1 < len(keys):
                next_col = keys[i+1]
                if "status" in next_col.lower():
                    st = str(row.get(next_col, '')).strip().replace('\xa0', '')
                    statuses.append(st)
                    
            if st.lower() != 'completed':
                days_away = (act_date - today).days
                if days_away <= 10:
                    pending.append((col, act_date))
                    
    return pending, "_".join(statuses)

def main():
    print("[INFO] Loading Excel file...")
    tasks = read_tasks(config.EXCEL_FILE, config.SHEET_NAME)
    
    if not tasks:
        print("[INFO] No tasks found.")
        return
        
    print(f"[INFO] Found {len(tasks)} applications.")
    
    qualifying_apps = []
    
    for row in tasks:
        app = row.get('APP', 'Unknown')
        
        pending_acts, all_statuses = get_pending_activities(row)
        
        if not pending_acts:
            print(f"[INFO] {app}: all requests completed. No notification required.")
            continue
            
        print(f"[INFO] {app}: pending request detected. Checking dates...")
        
        start_date = parse_date(row.get('Start Date'))
        end_date = parse_date(row.get('End Date'))
        
        reminder_types = []
        if start_date and is_in_reminder_window(start_date, config.REMINDER_DAYS_BEFORE):
            reminder_types.append(('START_DATE', start_date))
        if end_date and is_in_reminder_window(end_date, config.REMINDER_DAYS_BEFORE):
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

    # Build ONE consolidated message
    card = build_consolidated_message(qualifying_apps)
    
    if config.DRY_RUN:
        print("\n--- DRY RUN: WOULD SEND CONSOLIDATED ADAPTIVE CARD ---")
        print(json.dumps({"card": card}, indent=2))
        print("------------------------------------------------------\n")
    else:
        print("[INFO] Sending consolidated Teams notification...")
        success = send_notification(card)
        if success:
            print("[INFO] Teams notification sent successfully.")
            if config.ENABLE_DEDUPLICATION:
                for qa in qualifying_apps:
                    for r_type, r_date in qa["reminders"]:
                        update_state(qa["app"], r_type, str(r_date), qa["all_statuses"], "", "")
        else:
            print("[ERROR] Failed to send Teams notification.")

    print("[INFO] Processing complete.")

if __name__ == "__main__":
    main()
