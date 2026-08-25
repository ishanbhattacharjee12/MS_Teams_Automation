import requests
import config

def sanitize_url(url):
    if not url:
        return ""
    if "?" in url:
        return url.split("?")[0] + "?..."
    return url

def send_notification(card_dict):
    if not config.TEAMS_WEBHOOK_URL:
        print("[ERROR] TEAMS_WEBHOOK_URL is not set in environment.")
        return False, None, "TEAMS_WEBHOOK_URL is not set"
        
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
        if response.status_code >= 400:
            return False, response.status_code, response.text
        return True, response.status_code, ""
    except requests.exceptions.RequestException as e:
        err_msg = str(e)
        if config.TEAMS_WEBHOOK_URL and config.TEAMS_WEBHOOK_URL in err_msg:
            err_msg = err_msg.replace(config.TEAMS_WEBHOOK_URL, sanitize_url(config.TEAMS_WEBHOOK_URL))
        return False, None, err_msg
