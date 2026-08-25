import datetime
import pandas as pd
import json
import re
from date_checker import get_today, parse_date, get_urgency_info

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

def is_valid_metadata(val):
    if val is None or pd.isna(val):
        return False
    val_str = str(val).strip()
    if val_str.lower() in ["", "none", "nan", "n/a", "na"]:
        return False
    return True

def get_status_priority(status):
    s = str(status).strip().lower()
    if s == "in progress": return 1
    if "blocked" in s or "delayed" in s: return 2
    if s == "not started" or s == "yet to start": return 3
    if s == "completed": return 4
    return 5

def get_status_color(status):
    s = str(status).strip().lower()
    if "blocked" in s or "delayed" in s:
        return "Attention"
    if "yet to start" in s or "not started" in s:
        return "Warning"
    if "in progress" in s:
        return "Accent"
    if "completed" in s:
        return "Good"
    return "Default"

def build_consolidated_message(qualifying_apps, max_bytes=18000, config_params=None):
    """
    qualifying_apps: list of dicts:
    {
        "app": "appName",
        "row": { ... excel row ... },
        "reminders": [("START_DATE", date_obj), ...],
        "pending_activities": [("Dev R-Task", date_obj), ...],
        "all_statuses": "...",
        "overall_status": "...",
        "next_action": "...",
        "next_action_date": date_obj
    }
    
    Returns a list of Adaptive Card dictionaries split based on serialized payload size.
    """
    today = get_today()
    warning_days = config_params.get('SprintWarningDays', 10) if config_params else 10
    
    has_start = any("START_DATE" in [r[0] for r in qa["reminders"]] for qa in qualifying_apps)
    has_end = any("END_DATE" in [r[0] for r in qa["reminders"]] for qa in qualifying_apps)
    
    if has_end and has_start:
        title = "🚨 URGENT & IMPORTANT | ACCESS REQUESTS"
        title_color = "Attention"
    elif has_end:
        title = "🚨 URGENT | ACCESS REQUESTS"
        title_color = "Attention"
    else:
        title = "⚠️ IMPORTANT | ACCESS REQUESTS"
        title_color = "Warning"
        
    # Group by Responsible PM (standardizing casing/whitespace)
    pm_groups = {}
    for qa in qualifying_apps:
        row = qa["row"]
        email = str(row.get('Email/Teams', '')).strip()
        name = str(row.get('Name', '')).strip()
        pm_missing = row.get('pm_missing', False)
        
        if pm_missing or not name:
            user_key = "unassigned"
            group_name = "Unassigned PM / Requires Review"
            group_email = ""
        else:
            user_key = name.lower().strip()
            group_name = name.strip()
            group_email = email
            
        if user_key not in pm_groups:
            pm_groups[user_key] = {
                "name": group_name,
                "email": group_email,
                "pm_missing": pm_missing or (user_key == "unassigned"),
                "apps": []
            }
            
        pm_groups[user_key]["apps"].append(qa)
        
    sorted_keys = sorted([k for k in pm_groups.keys() if k != "unassigned"])
    if "unassigned" in pm_groups:
        sorted_keys.append("unassigned")

    def create_fresh_card_template():
        body = [
            {
                "type": "TextBlock",
                "text": title,
                "weight": "Bolder",
                "size": "Medium",
                "color": title_color
            },
            # Index 1: Placeholder for global summary
            {
                "type": "TextBlock",
                "text": "",
                "wrap": True
            },
            # Index 2: Placeholder for attention text
            {
                "type": "TextBlock",
                "text": "",
                "wrap": True,
                "spacing": "Medium"
            },
            # Index 3: Description TextBlock
            {
                "type": "TextBlock",
                "text": "The following applications require attention:",
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
        return card

    cards = []
    block_stream = []
    
    for user_key in sorted_keys:
        group = pm_groups[user_key]
        apps = group["apps"]
        
        pm_header_text = f"👤 {group['name']} ({len(apps)} application{'s' if len(apps) > 1 else ''} requiring attention)"
        pm_header_block = {
            "type": "TextBlock",
            "text": pm_header_text,
            "wrap": True,
            "size": "Medium",
            "weight": "Bolder",
            "spacing": "Large",
            "separator": True
        }
        
        block_stream.append({
            "type": "PM_HEADER",
            "user_key": user_key,
            "block": pm_header_block,
            "pm_info": {
                "name": group["name"],
                "email": group["email"],
                "pm_missing": group["pm_missing"]
            }
        })
        
        for app_idx, qa in enumerate(apps, 1):
            status_color = get_status_color(qa["overall_status"])
            
            # Format next action and due date using the centralized get_urgency_info
            next_action_date = qa["next_action_date"]
            urgency_label, urgency_badge, urgency_suffix = get_urgency_info(next_action_date, warning_days, today)
            
            if next_action_date:
                due_str = f"{format_date_str(next_action_date)}{urgency_suffix}"
            else:
                due_str = "Not available"
                
            pm_suffix = f" [{group['name']}]" if not group["pm_missing"] and group["name"] else ""
            app_items = [
                {
                    "type": "TextBlock",
                    "text": f"{app_idx}. **{qa['app']}**{pm_suffix}",
                    "size": "Medium",
                    "weight": "Bolder"
                }
            ]
            
            if qa.get("row", {}).get("pm_missing"):
                app_items.append({
                    "type": "TextBlock",
                    "text": "⚠️ No PM assigned in tracker",
                    "wrap": True,
                    "color": "Attention",
                    "weight": "Bolder"
                })
                
            file_type = qa.get("row", {}).get("file_type", "")
            frequency = qa.get("row", {}).get("frequency", "")
            has_ft = is_valid_metadata(file_type)
            has_freq = is_valid_metadata(frequency)
            
            if has_ft and has_freq:
                meta_text = f"📄 File Type: **{file_type}**  ·  Frequency: **{frequency}**"
                app_items.append({
                    "type": "TextBlock",
                    "text": meta_text,
                    "wrap": True,
                    "spacing": "Small"
                })
            elif has_ft:
                meta_text = f"📄 File Type: **{file_type}**"
                app_items.append({
                    "type": "TextBlock",
                    "text": meta_text,
                    "wrap": True,
                    "spacing": "Small"
                })
            elif has_freq:
                meta_text = f"🔄 Frequency: **{frequency}**"
                app_items.append({
                    "type": "TextBlock",
                    "text": meta_text,
                    "wrap": True,
                    "spacing": "Small"
                })
                
            # Clean up empty/unavailable metadata: only show what is available!
            if qa["overall_status"] != "Not available":
                app_items.append({
                    "type": "TextBlock",
                    "text": f"Status: **{qa['overall_status']}**",
                    "wrap": True,
                    "color": status_color
                })
                
            if qa["next_action"] != "No actionable task identified":
                app_items.extend([
                    {
                        "type": "TextBlock",
                        "text": f"Next Action: **{urgency_badge}{qa['next_action']}**",
                        "wrap": True
                    },
                    {
                        "type": "TextBlock",
                        "text": f"Due: **{due_str}**",
                        "wrap": True
                    }
                ])
            
            # Format pending activities using get_urgency_info
            overdue_acts = []
            due_soon_acts = []
            other_acts = []
            
            for act_name, act_date in qa["pending_activities"]:
                urg_label, _, _ = get_urgency_info(act_date, warning_days, today)
                formatted_item = f"• {act_name} — {format_date_str(act_date)}"
                if urg_label == 'overdue':
                    overdue_acts.append(formatted_item)
                elif urg_label == 'due_soon':
                    due_soon_acts.append(formatted_item)
                else:
                    other_acts.append(formatted_item)
                    
            parts = []
            if overdue_acts:
                parts.append(f"🔴 Overdue — {len(overdue_acts)}")
                parts.extend(overdue_acts)
            if due_soon_acts:
                if parts:
                    parts.append("")
                parts.append(f"🟡 Due Soon — {len(due_soon_acts)}")
                parts.extend(due_soon_acts)
            if other_acts:
                if parts:
                    parts.append("")
                parts.extend(other_acts)
                
            pending_str = "\n".join(parts) if parts else "None"
            
            app_items.append({
                "type": "TextBlock",
                "text": "**Pending Activities**\n" + pending_str,
                "wrap": True,
                "spacing": "Small"
            })
            
            app_container = {
                "type": "Container",
                "spacing": "Medium",
                "separator": True,
                "items": app_items
            }
            
            block_stream.append({
                "type": "APP",
                "user_key": user_key,
                "block": app_container,
                "pm_info": {
                    "name": group["name"],
                    "email": group["email"],
                    "pm_missing": group["pm_missing"]
                }
            })

    current_card = create_fresh_card_template()
    card_pms_seen = []
    card_pm_keys = set()
    apps_added_to_card = 0
    global_pm_keys_committed = set()
    
    idx = 0
    while idx < len(block_stream):
        item = block_stream[idx]
        pm_key = item["user_key"]
        need_pm_header = (pm_key not in card_pm_keys)
        
        temp_blocks = []
        if need_pm_header:
            if pm_key in global_pm_keys_committed:
                pm_header_text = f"👤 {item['pm_info']['name']} (Continued)"
            else:
                pm_header_text = f"👤 {item['pm_info']['name']}\n{len(pm_groups[pm_key]['apps'])} application{'s' if len(pm_groups[pm_key]['apps']) > 1 else ''} requiring attention"
            
            pm_header_block = {
                "type": "TextBlock",
                "text": pm_header_text,
                "wrap": True,
                "size": "Medium",
                "weight": "Bolder",
                "spacing": "Large",
                "separator": True
            }
            temp_blocks.append(pm_header_block)
            
        if item["type"] == "APP":
            app_block_to_add = dict(item["block"])
            if need_pm_header:
                app_block_to_add["separator"] = False
                app_block_to_add["spacing"] = "Small"
            temp_blocks.append(app_block_to_add)
            
        if not temp_blocks:
            idx += 1
            continue
            
        # Test size
        test_card = {
            "type": "AdaptiveCard",
            "version": "1.2",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "body": current_card["body"] + temp_blocks
        }
        
        test_pms = list(card_pms_seen)
        test_pm_keys = set(card_pm_keys)
        if pm_key not in test_pm_keys:
            test_pm_keys.add(pm_key)
            test_pms.append(item["pm_info"])
            
        test_mentions_data = []
        for pm in test_pms:
            if not pm["pm_missing"] and pm["email"] and "@" in pm["email"]:
                mention_str = f"<at>{pm['name']}</at>"
                test_mentions_data.append({
                    "type": "mention",
                    "text": mention_str,
                    "mentioned": {
                        "id": pm["email"],
                        "name": pm["name"]
                    }
                })
        if test_mentions_data:
            test_card["msteams"] = {"entities": test_mentions_data}
            
        test_size = len(json.dumps(test_card).encode('utf-8'))
        
        if test_size > (max_bytes - 1000) and apps_added_to_card > 0:
            cards.append(current_card)
            global_pm_keys_committed.update(card_pm_keys)
            current_card = create_fresh_card_template()
            card_pms_seen = []
            card_pm_keys = set()
            apps_added_to_card = 0
        else:
            for block in temp_blocks:
                current_card["body"].append(block)
            if pm_key not in card_pm_keys:
                card_pm_keys.add(pm_key)
                card_pms_seen.append(item["pm_info"])
            if item["type"] == "APP":
                apps_added_to_card += 1
            idx += 1
            
    if apps_added_to_card > 0:
        cards.append(current_card)
    elif not cards:
        cards.append(current_card)
        
    # Finalize card headers and mentions
    total_cards = len(cards)
    global_apps_count = len(qualifying_apps)
    global_pms_count = len([k for k in pm_groups.keys() if k != "unassigned"])
    
    for card in cards:
        # 1. Update Global Summary Block
        summary_text = f"📊 {global_apps_count} Application{'s' if global_apps_count > 1 else ''} · {global_pms_count} PM{'s' if global_pms_count > 1 else ''} · {total_cards} Card{'s' if total_cards > 1 else ''}"
        card["body"][1] = {
            "type": "TextBlock",
            "text": summary_text,
            "wrap": True,
            "weight": "Bolder"
        }
        
        # 2. Extract PMs actually present on this card
        card_pms_on_this_card = []
        card_pm_names_seen = set()
        for block in card["body"]:
            if block["type"] == "TextBlock" and block["text"].startswith("👤 "):
                header_text = block["text"]
                pm_name = header_text.split("👤 ")[1].split("\n")[0].split(" (")[0].split(" —")[0].strip()
                for group in pm_groups.values():
                    if group["name"] == pm_name and pm_name not in card_pm_names_seen:
                        card_pm_names_seen.add(pm_name)
                        card_pms_on_this_card.append(group)
                        break
                        
        # 3. Build Attention Block & Mentions
        card_mention_strings = []
        card_mentions_data = []
        for pm in card_pms_on_this_card:
            if pm["pm_missing"]:
                card_mention_strings.append(f"**{pm['name']}**")
            elif pm["email"] and "@" in pm["email"]:
                mention_str = f"<at>{pm['name']}</at>"
                card_mention_strings.append(mention_str)
                card_mentions_data.append({
                    "type": "mention",
                    "text": mention_str,
                    "mentioned": {
                        "id": pm["email"],
                        "name": pm["name"]
                    }
                })
            else:
                card_mention_strings.append(f"**{pm['name']}**")
                
        if card_mention_strings:
            attention_text = "Attention required from:\n" + " · ".join(card_mention_strings)
        else:
            attention_text = "No attention required."
            
        card["body"][2] = {
            "type": "TextBlock",
            "text": attention_text,
            "wrap": True,
            "spacing": "Medium"
        }
        
        if card_mentions_data:
            card["msteams"] = {"entities": card_mentions_data}
        else:
            if "msteams" in card:
                del card["msteams"]
                
    return cards
