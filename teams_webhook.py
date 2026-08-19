import requests
import config

def send_notification(card_dict):
    if not config.TEAMS_WEBHOOK_URL:
        print("[ERROR] TEAMS_WEBHOOK_URL is not set in environment.")
        return False
        
    payload = {
        "card": card_dict
    }
    
    try:
        response = requests.post(
            config.TEAMS_WEBHOOK_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        response.raise_for_status()
        return True
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Teams webhook failed: {e}")
        return False
