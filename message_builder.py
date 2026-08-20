import datetime
import pandas as pd
import json

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

def build_consolidated_message(qualifying_apps, max_bytes=18000):
    """
    qualifying_apps: list of dicts:
    {
        "app": "appName",
        "row": { ... excel row ... },
        "reminders": [("START_DATE", date_obj), ...],
        "pending_activities": [("Dev R-Task", date_obj), ...],
        "all_statuses": "..."
    }
    
    Returns a list of Adaptive Card dictionaries split based on serialized payload size.
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
        pm_missing = row.get('pm_missing', False)
        
        if pm_missing or not name:
            user_key = "Unassigned PM"
            name = "Unassigned PM / Requires Review"
            email = ""
        else:
            user_key = email if email else name
            
        if user_key not in pm_groups:
            pm_groups[user_key] = {
                "name": name,
                "email": email,
                "pm_missing": pm_missing or (user_key == "Unassigned PM"),
                "apps": []
            }
            
        app_data = {
            "name": qa["app"],
            "status": str(row.get('Overall Status', '')).strip(),
            "next_action": str(row.get('Next Action', '')).strip(),
            "next_action_date": row.get('Next Action Date'),
            "pending_activities": qa["pending_activities"], # List of (act_name, date_obj)
            "pm_missing": pm_missing or (user_key == "Unassigned PM")
        }
        pm_groups[user_key]["apps"].append(app_data)
    
    # Process groups and build header components
    mention_strings = []
    
    for user_key, group in pm_groups.items():
        name = group["name"]
        email = group["email"]
        pm_missing = group.get("pm_missing", False)
        is_valid_upn = bool(email and "@" in email)
        
        # Build mention
        if pm_missing:
            mention_str = f"**{name}**"
            if user_key not in unique_mentions:
                unique_mentions.add(user_key)
                mention_strings.append(mention_str)
        elif is_valid_upn:
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

    cards = []
    
    def create_fresh_card():
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
        
        card = {
            "type": "AdaptiveCard",
            "version": "1.2",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "body": body
        }
        
        if mentions_data:
            card["msteams"] = {
                "entities": list(mentions_data)
            }
        return card

    current_card = create_fresh_card()
    apps_in_current_card = 0
    
    for user_key, group in pm_groups.items():
        name = group["name"]
        email = group["email"]
        pm_missing = group.get("pm_missing", False)
        is_valid_upn = bool(email and "@" in email)
        
        if pm_missing:
            mention_str = f"**{name}**"
        else:
            mention_str = f"<at>{name}</at>" if is_valid_upn else f"**{name}**"
            
        pm_header_added = False
        
        for i, app_data in enumerate(group["apps"]):
            status_color = "Default"
            if "in progress" in app_data["status"].lower():
                status_color = "Warning"
            elif "completed" in app_data["status"].lower():
                status_color = "Good"
                
            date_str = format_date_str(app_data["next_action_date"])
            
            if app_data["pending_activities"]:
                formatted_acts = []
                for act_name, act_date in app_data["pending_activities"]:
                    formatted_acts.append(f"- {act_name} — {format_date_str(act_date)}")
                pending_str = "\n".join(formatted_acts)
            else:
                pending_str = "None"
                
            app_items = [
                {
                    "type": "TextBlock",
                    "text": f"**Applications:** {app_data['name']}",
                    "wrap": True
                }
            ]
            
            if app_data.get("pm_missing"):
                app_items.append({
                    "type": "TextBlock",
                    "text": "⚠️ No PM assigned in tracker (Requires Review)",
                    "wrap": True,
                    "color": "Attention",
                    "weight": "Bolder"
                })
                
            app_items.extend([
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
            ])
            
            app_container = {
                "type": "Container",
                "spacing": "Medium",
                "separator": True,
                "items": app_items
            }
            
            temp_blocks = []
            add_pm_header = not pm_header_added
            
            if add_pm_header:
                pm_header_block = {
                    "type": "TextBlock",
                    "text": mention_str,
                    "wrap": True,
                    "size": "Medium",
                    "weight": "Bolder",
                    "spacing": "Large",
                    "separator": True
                }
                temp_blocks.append(pm_header_block)
                app_container_temp = dict(app_container)
                app_container_temp["separator"] = False
                app_container_temp["spacing"] = "Small"
                temp_blocks.append(app_container_temp)
            else:
                temp_blocks.append(app_container)
                
            test_card = {
                "type": "AdaptiveCard",
                "version": "1.2",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "body": current_card["body"] + temp_blocks
            }
            if mentions_data:
                test_card["msteams"] = current_card["msteams"]
                
            test_size = len(json.dumps(test_card).encode('utf-8'))
            
            if test_size > max_bytes and apps_in_current_card > 0:
                cards.append(current_card)
                
                current_card = create_fresh_card()
                apps_in_current_card = 0
                
                pm_header_block = {
                    "type": "TextBlock",
                    "text": mention_str,
                    "wrap": True,
                    "size": "Medium",
                    "weight": "Bolder",
                    "spacing": "Large",
                    "separator": True
                }
                app_container["separator"] = False
                app_container["spacing"] = "Small"
                
                current_card["body"].append(pm_header_block)
                current_card["body"].append(app_container)
                apps_in_current_card += 1
                pm_header_added = True
            else:
                if add_pm_header:
                    app_container["separator"] = False
                    app_container["spacing"] = "Small"
                    current_card["body"].append(pm_header_block)
                    pm_header_added = True
                    
                current_card["body"].append(app_container)
                apps_in_current_card += 1
                
    if apps_in_current_card > 0:
        cards.append(current_card)
        
    if not cards:
        cards.append(create_fresh_card())
        
    return cards
