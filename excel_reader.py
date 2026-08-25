import pandas as pd
import os
import sys
import shutil
import tempfile
import re

def normalize_header(header_val):
    s = str(header_val).strip().lower()
    s = re.sub(r'[^a-z0-9]+', '', s)
    return s

NORM_MAP = {
    'application': 'Application',
    'app': 'Application',
    
    'responsiblepm': 'Responsible PM',
    'pmname': 'Responsible PM',
    'pm': 'Responsible PM',
    'name': 'Responsible PM',
    
    'sprint': 'Sprint',
    
    'overallstatus': 'Overall Status',
    'status': 'Overall Status',
    
    'nextaction': 'Next Action',
    
    'nextactiondate': 'Next Action Date',
    
    'sprintstart': 'Sprint Start',
    'effectivesprintstart': 'Effective Sprint Start',
    
    'sprintend': 'Sprint End',
    'effectivesprintend': 'Effective Sprint End',
    
    'releasedate': 'Release Date',
    'effectivereleasedate': 'Effective Release Date',
    
    'appreadyby': 'App Ready By',
    'sprintreadinessdate': 'Sprint Readiness Date',
    'sprintreadinessstatus': 'Sprint Readiness Status',
    'validationnotes': 'Validation / Notes',
    'daysoverdue': 'Days Overdue',
    
    'filetypeallenvs': 'file_type',
    'filetype': 'file_type',
    'frequencyallenvs': 'frequency',
    'frequency': 'frequency'
}

def parse_pm_value(pm_val):
    if pm_val is None or pd.isna(pm_val):
        return None
    pm_val = str(pm_val).strip()
    pm_val = pm_val.replace('\xa0', ' ')
    if not pm_val or pm_val.lower() == 'nan' or pm_val.lower() == 'none':
        return None
        
    # Match Name <email>
    match = re.match(r"^(.*?)\s*<([^>@]+@[^>]+)>$", pm_val)
    if match:
        name = match.group(1).strip()
        email = match.group(2).strip()
        return {"name": name, "email": email, "pm_missing": False}
        
    return {"name": pm_val, "email": "", "pm_missing": False}

def find_pm_mappings(file_path):
    mappings = {}
    if not os.path.exists(file_path):
        return mappings
        
    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)
        xl = pd.ExcelFile(temp_path)
        sheet_names = xl.sheet_names
        xl.close()
        
        # Scan all sheets
        for sheet in sheet_names:
            try:
                df = pd.read_excel(temp_path, sheet_name=sheet, header=None)
                if df.empty:
                    continue
                    
                header_row_idx = None
                name_col_idx = None
                email_col_idx = None
                
                # Scan first 10 rows for name and email headers
                for r_idx in range(min(10, len(df))):
                    row_vals = [str(x).strip().lower() if pd.notna(x) else "" for x in df.iloc[r_idx]]
                    temp_name_idx = None
                    temp_email_idx = None
                    
                    for c_idx, val in enumerate(row_vals):
                        val_clean = re.sub(r'[^a-z]+', '', val)
                        if val_clean in ['name', 'responsiblepm', 'pmname', 'pm']:
                            temp_name_idx = c_idx
                        elif val_clean in ['email', 'emailid', 'pmemail', 'teamsemail', 'emailteams']:
                            temp_email_idx = c_idx
                            
                    if temp_name_idx is not None and temp_email_idx is not None:
                        header_row_idx = r_idx
                        name_col_idx = temp_name_idx
                        email_col_idx = temp_email_idx
                        break
                        
                if header_row_idx is not None:
                    for r_idx in range(header_row_idx + 1, len(df)):
                        name_val = df.iloc[r_idx, name_col_idx]
                        email_val = df.iloc[r_idx, email_col_idx]
                        if pd.notna(name_val) and pd.notna(email_val):
                            name_str = str(name_val).strip().replace('\xa0', ' ')
                            email_str = str(email_val).strip()
                            if name_str and email_str and "@" in email_str:
                                name_lower = name_str.lower()
                                if name_lower in mappings:
                                    existing_email = mappings[name_lower]["email"]
                                    if existing_email.lower() != email_str.lower():
                                        print(f"[WARNING] Conflicting PM mappings found for '{name_str}': '{existing_email}' vs '{email_str}'. Keeping the first mapping '{existing_email}'.")
                                else:
                                    mappings[name_lower] = {
                                        "name": name_str,
                                        "email": email_str
                                    }
            except Exception as e:
                print(f"[WARNING] Could not scan sheet '{sheet}' for PM mappings: {e}")
    except Exception as e:
        print(f"[WARNING] Could not read mappings from Excel file: {e}")
    finally:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
            
    return mappings

def resolve_pm_identity(pm_val, mappings, seen_pms=None):
    parsed = parse_pm_value(pm_val)
    if not parsed:
        return {"name": "", "email": "", "pm_missing": True}
        
    name = parsed["name"]
    email = parsed["email"]
    
    name_lower = name.lower()
    
    if name_lower in mappings:
        standard_name = mappings[name_lower]["name"]
        resolved_email = mappings[name_lower]["email"]
    else:
        if seen_pms is not None:
            if name_lower not in seen_pms:
                seen_pms[name_lower] = name
            standard_name = seen_pms[name_lower]
        else:
            standard_name = name
        resolved_email = email
        
    if email:
        resolved_email = email
        
    if standard_name and not resolved_email:
        print(f"[WARNING] Could not resolve email for PM: '{standard_name}'")
        
    return {"name": standard_name, "email": resolved_email, "pm_missing": False}

def read_config_params(file_path):
    params = {
        'SprintWorkingDays': 10,
        'ReadyOffsetWD': -1,
        'SROffsetFromStartWD': -5,
        'OUROffsetFromSR_WD': -1,
        'RTSKOffsetFromOUR_WD': -2,
        'CABOffsetFromReleaseWD': -1,
        'CodeReviewCalendarDays': 2,
        'CodeReviewInviteCalendarDays': 4,
        'SprintWarningDays': 10,
        'SprintReadinessOffsetWD': -2
    }
    
    if not os.path.exists(file_path):
        return params
        
    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)
        xl = pd.ExcelFile(temp_path)
        has_config = 'Config' in xl.sheet_names
        xl.close()
        
        if has_config:
            df = pd.read_excel(temp_path, sheet_name='Config', header=None)
            for row in df.values:
                if len(row) > 1:
                    param_name = str(row[0]).strip()
                    if param_name in params:
                        val = row[1]
                        if pd.notna(val):
                            try:
                                params[param_name] = int(val)
                            except ValueError:
                                pass
    except Exception as e:
        print(f"[WARNING] Could not read Config sheet: {e}")
    finally:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
            
    return params

def read_tasks(file_path, sheet_name=None):
    if not os.path.exists(file_path):
        print(f"[ERROR] Excel file not found: {file_path}")
        sys.exit(1)
        
    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)
        
        xl = pd.ExcelFile(temp_path)
        selected_sheet = sheet_name
        if not selected_sheet:
            if 'Tracker' in xl.sheet_names:
                selected_sheet = 'Tracker'
            elif 'Sheet1' in xl.sheet_names:
                selected_sheet = 'Sheet1'
            else:
                selected_sheet = xl.sheet_names[0]
        xl.close()
        
        # Build PM mappings
        mappings = find_pm_mappings(temp_path)
        
        # Determine header structure
        df_preview = pd.read_excel(temp_path, sheet_name=selected_sheet, header=None, nrows=2)
        row0_col0 = str(df_preview.iloc[0, 0]).strip().lower() if len(df_preview) > 0 and pd.notna(df_preview.iloc[0, 0]) else ""
        row1_col0 = str(df_preview.iloc[1, 0]).strip().lower() if len(df_preview) > 1 and pd.notna(df_preview.iloc[1, 0]) else ""
        
        row0_norm = normalize_header(row0_col0)
        row1_norm = normalize_header(row1_col0)
        
        is_multi_index = False
        header_idx = 0
        
        if row1_norm in ['application', 'app']:
            is_multi_index = True
            header_idx = [0, 1]
        elif row0_norm in ['application', 'app']:
            is_multi_index = False
            header_idx = 0
        else:
            is_multi_index = False
            header_idx = 0
            
        df = pd.read_excel(temp_path, sheet_name=selected_sheet, header=header_idx)
        
        if df.empty:
            print(f"[ERROR] Excel file is empty: {file_path}")
            sys.exit(1)
            
        records = []
        df_records = df.to_dict('records')
        seen_pms = {}
        
        for row in df_records:
            row_dict = {}
            for col in df.columns:
                val = row[col]
                row_dict[col] = val
                
                col_str = col[1] if isinstance(col, tuple) else col
                col_str_clean = str(col_str).strip()
                
                norm_key = normalize_header(col_str_clean)
                if norm_key in NORM_MAP:
                    standard_key = NORM_MAP[norm_key]
                    row_dict[standard_key] = val
                
                if col_str_clean not in row_dict:
                    row_dict[col_str_clean] = val
            
            app_val = row_dict.get('Application', '')
            row_dict['Application'] = app_val
            row_dict['APP'] = app_val
            
            pm_val = row_dict.get('Responsible PM', '')
            identity = resolve_pm_identity(pm_val, mappings, seen_pms)
            row_dict['Responsible PM'] = identity['name']
            row_dict['Name'] = identity['name']
            row_dict['pm_missing'] = identity['pm_missing']
            row_dict['Email/Teams'] = identity['email']
                
            if 'Start Date' not in row_dict:
                if 'Sprint Start' in row_dict:
                    row_dict['Start Date'] = row_dict['Sprint Start']
                elif 'Effective Sprint Start' in row_dict:
                    row_dict['Start Date'] = row_dict['Effective Sprint Start']
                    
            if 'End Date' not in row_dict:
                if 'Sprint End' in row_dict:
                    row_dict['End Date'] = row_dict['Sprint End']
                elif 'Effective Sprint End' in row_dict:
                    row_dict['End Date'] = row_dict['Effective Sprint End']
            
            # Ensure file_type and frequency keys exist in row_dict
            row_dict['file_type'] = row_dict.get('file_type', '')
            row_dict['frequency'] = row_dict.get('frequency', '')
            
            for k in list(row_dict.keys()):
                if pd.isna(row_dict[k]):
                    row_dict[k] = ""
            records.append(row_dict)
            
        return records
    except Exception as e:
        print(f"[ERROR] Could not read Excel file: {e}")
        sys.exit(1)
    finally:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception as e:
            print(f"[WARNING] Could not remove temp file {temp_path}: {e}")
