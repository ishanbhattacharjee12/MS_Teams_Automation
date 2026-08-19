import datetime
import pandas as pd

def format_date_str(date_val):
    if isinstance(date_val, datetime.date) and not isinstance(date_val, datetime.datetime):
        dt = date_val
    elif hasattr(date_val, 'date'):
        dt = date_val.date()
    elif isinstance(date_val, str) and date_val:
        try:
            dt = datetime.datetime.strptime(date_val.split()[0], "%Y-%m-%d").date()
        except Exception:
            return str(date_val).strip()
    else:
        return str(date_val).strip()
        
    return dt.strftime("%d-%b-%Y")

def get_status_priority(status):
    s = str(status).strip().lower()
    if s == "in progress": return 1
    if "blocked" in s or "delayed" in s: return 2
    if s == "not started": return 3
    if s == "completed": return 4
    return 5

def build_consolidated_message(qualifying_apps):
    """
    qualifying_apps: list of dicts:
    {
        "app": "appName",
        "row": { ... excel row ... },
        "reminders": [("START_DATE", date_obj), ...],
        "pending_activities": ["Dev OUR", ...],
        "all_statuses": "..."
    }
    """
    has_start = any("START_DATE" in [r[0] for r in qa["reminders"]] for qa in qualifying_apps)
    has_end = any("END_DATE" in [r[0] for r in qa["reminders"]] for qa in qualifying_apps)
    
    if has_end and has_start:
        title = "URGENT & IMPORTANT | ACCESS REQUESTS"
        title_color = "Attention"
    elif has_end:
        title = "URGENT | ACCESS REQUESTS"
        title_color = "Attention"
    else:
        title = "IMPORTANT | ACCESS REQUESTS"
        title_color = "Warning"
        
    # Group by Responsible PM
    pm_groups = {}
    mentions_data = []
    unique_mentions = set()
    
    for qa in qualifying_apps:
        row = qa["row"]
        email = str(row.get('Email/Teams', '')).strip()
        name = str(row.get('Name', '')).strip()
        if not name:
            name = email if email else "Responsible Person"
            
        user_key = email if email else name
        
        if user_key not in pm_groups:
            pm_groups[user_key] = {
                "name": name,
                "email": email,
                "apps": []
            }
            
        app_data = {
            "name": qa["app"],
            "status": str(row.get('Overall Status', '')).strip(),
            "next_action": str(row.get('Next Action', '')).strip(),
            "next_action_date": row.get('Next Action Date'),
            "pending_activities": qa["pending_activities"] # List of (act_name, date_obj)
        }
        pm_groups[user_key]["apps"].append(app_data)
    
    # Process groups and build header components
    mention_strings = []
    
    for user_key, group in pm_groups.items():
        name = group["name"]
        email = group["email"]
        is_valid_upn = bool(email and "@" in email)
        
        # Build mention
        if is_valid_upn:
            mention_str = f"<at>{name}</at>"
            if user_key not in unique_mentions:
                unique_mentions.add(user_key)
                mention_strings.append(mention_str)
                mentions_data.append({
                    "type": "mention",
                    "text": mention_str,
                    "mentioned": {
                        "id": email,
                        "name": name
                    }
                })
        else:
            mention_str = f"**{name}**"
            if user_key not in unique_mentions:
                unique_mentions.add(user_key)
                mention_strings.append(mention_str)

    attention_text = "Attention required from:\n" + " · ".join(mention_strings)

    body = [
        {
            "type": "TextBlock",
            "text": title,
            "weight": "Bolder",
            "size": "Medium",
            "color": title_color
        },
        {
            "type": "TextBlock",
            "text": attention_text,
            "wrap": True,
            "spacing": "Medium"
        },
        {
            "type": "TextBlock",
            "text": "The following applications have incomplete access requests:",
            "wrap": True,
            "spacing": "Medium"
        }
    ]
    
    for user_key, group in pm_groups.items():
        name = group["name"]
        email = group["email"]
        is_valid_upn = bool(email and "@" in email)
        mention_str = f"<at>{name}</at>" if is_valid_upn else f"**{name}**"
        
        # Add a section for EACH application
        for i, app_data in enumerate(group["apps"]):
            # Format status color
            status_color = "Default"
            if "in progress" in app_data["status"].lower():
                status_color = "Warning"
            elif "completed" in app_data["status"].lower():
                status_color = "Good"
                
            # Date format
            date_str = format_date_str(app_data["next_action_date"])
            
            # Format pending activities as bullet points
            if app_data["pending_activities"]:
                formatted_acts = []
                for act_name, act_date in app_data["pending_activities"]:
                    formatted_acts.append(f"- {act_name} — {format_date_str(act_date)}")
                pending_str = "\n".join(formatted_acts)
            else:
                pending_str = "None"
                
            app_container = {
                "type": "Container",
                "spacing": "Medium",
                "separator": True,
                "items": [
                    {
                        "type": "TextBlock",
                        "text": f"**Applications:** {app_data['name']}",
                        "wrap": True
                    },
                    {
                        "type": "TextBlock",
                        "text": f"**Overall Status:** {app_data['status'] if app_data['status'] else 'N/A'}",
                        "wrap": True,
                        "color": status_color
                    },
                    {
                        "type": "TextBlock",
                        "text": f"**Next Action:** {app_data['next_action'] if app_data['next_action'] else 'N/A'}",
                        "wrap": True
                    },
                    {
                        "type": "TextBlock",
                        "text": f"**Next Action Date:** {date_str}",
                        "wrap": True
                    },
                    {
                        "type": "TextBlock",
                        "text": "**Pending Activities:**\n" + pending_str,
                        "wrap": True,
                        "spacing": "Small"
                    }
                ]
            }
            
            if i == 0:
                body.append({
                    "type": "TextBlock",
                    "text": mention_str,
                    "wrap": True,
                    "size": "Medium",
                    "weight": "Bolder",
                    "spacing": "Large",
                    "separator": True
                })
                app_container["separator"] = False
                app_container["spacing"] = "Small"
                
            body.append(app_container)

    card = {
        "type": "AdaptiveCard",
        "version": "1.2",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "body": body
    }
    
    if mentions_data:
        card["msteams"] = {
            "entities": mentions_data
        }
        
    return card
