import json
import os

STATE_FILE = "notification_state.json"

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_state(state):
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(state, f, indent=4)
    except Exception as e:
        print(f"[ERROR] Failed to save notification state: {e}")

def get_state_key(app, reminder_type, reminder_date, dev_req, qa_req, prod_req):
    return f"{app}_{reminder_type}_{reminder_date}_{dev_req}_{qa_req}_{prod_req}"

def should_send_notification(app, reminder_type, reminder_date, dev_req, qa_req, prod_req, enable_dedup):
    if not enable_dedup:
        return True
    
    state = load_state()
    current_key = get_state_key(app, reminder_type, reminder_date, dev_req, qa_req, prod_req)
    state_key_prefix = f"{app}_{reminder_type}"
    
    if state_key_prefix in state:
        if state[state_key_prefix] == current_key:
            return False
            
    return True

def update_state(app, reminder_type, reminder_date, dev_req, qa_req, prod_req):
    state = load_state()
    state_key_prefix = f"{app}_{reminder_type}"
    state[state_key_prefix] = get_state_key(app, reminder_type, reminder_date, dev_req, qa_req, prod_req)
    save_state(state)
