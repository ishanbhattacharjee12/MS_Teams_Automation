import pandas as pd
import os
import sys
import shutil
import tempfile

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
        if 'Config' in xl.sheet_names:
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
        
        # Read Sheet2 for emails if it exists
        email_map = {}
        if 'Sheet2' in xl.sheet_names:
            try:
                df_emails = pd.read_excel(temp_path, sheet_name="Sheet2")
                df_emails.columns = df_emails.columns.str.strip()
                df_emails['Name'] = df_emails['Name'].astype(str).str.replace(r'\xa0', ' ', regex=True).str.strip()
                email_map = dict(zip(df_emails['Name'], df_emails['Email']))
            except Exception as e:
                print(f"[WARNING] Could not read Sheet2 for emails: {e}")
        
        # Determine header structure
        df_preview = pd.read_excel(temp_path, sheet_name=selected_sheet, header=None, nrows=2)
        row0_col0 = str(df_preview.iloc[0, 0]).strip().lower() if len(df_preview) > 0 and pd.notna(df_preview.iloc[0, 0]) else ""
        row1_col0 = str(df_preview.iloc[1, 0]).strip().lower() if len(df_preview) > 1 and pd.notna(df_preview.iloc[1, 0]) else ""
        
        is_multi_index = False
        header_idx = 0
        
        if "application" in row1_col0 or "app" in row1_col0:
            is_multi_index = True
            header_idx = [0, 1]
        elif "application" in row0_col0 or "app" in row0_col0:
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
        
        for row in df_records:
            row_dict = {}
            if is_multi_index:
                # Populate tuple keys
                for col_tuple in df.columns:
                    group, col_name = col_tuple
                    val = row[col_tuple]
                    row_dict[col_tuple] = val
                    
                    col_name_clean = str(col_name).strip()
                    if col_name_clean not in row_dict:
                        row_dict[col_name_clean] = val
                
                app_val = row_dict.get('Application', '')
                row_dict['Application'] = app_val
                row_dict['APP'] = app_val
                
                pm_val = str(row_dict.get('Responsible PM', '')).strip()
                if pm_val == "nan" or pm_val == "":
                    row_dict['Responsible PM'] = ""
                    row_dict['Name'] = ""
                    row_dict['pm_missing'] = True
                    row_dict['Email/Teams'] = ""
                else:
                    row_dict['Responsible PM'] = pm_val
                    row_dict['Name'] = pm_val
                    row_dict['pm_missing'] = False
                    row_dict['Email/Teams'] = email_map.get(pm_val, "")
                    
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
            else:
                for col_name in df.columns:
                    col_name_clean = str(col_name).strip()
                    row_dict[col_name_clean] = row[col_name]
                
                if 'APP' not in row_dict and 'Application' in row_dict:
                    row_dict['APP'] = row_dict['Application']
                    
                pm_val = str(row_dict.get('Responsible PM', '')).strip()
                if pm_val == "nan" or pm_val == "":
                    row_dict['Responsible PM'] = ""
                    row_dict['Name'] = ""
                    row_dict['pm_missing'] = True
                    row_dict['Email/Teams'] = ""
                else:
                    row_dict['Responsible PM'] = pm_val
                    row_dict['Name'] = pm_val
                    row_dict['pm_missing'] = False
                    row_dict['Email/Teams'] = email_map.get(pm_val, "")
                    
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
