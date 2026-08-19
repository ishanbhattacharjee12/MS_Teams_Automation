import pandas as pd
import os
import sys
import shutil
import tempfile

def read_tasks(file_path, sheet_name=None):
    if not os.path.exists(file_path):
        print(f"[ERROR] Excel file not found: {file_path}")
        sys.exit(1)
        
    # Copy file to a temporary location to avoid PermissionError if open in Excel/OneDrive
    fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)
        
        # Check header
        df_preview = pd.read_excel(temp_path, sheet_name=sheet_name if sheet_name else 0, nrows=2)
        has_unnamed = any("Unnamed:" in str(col) for col in df_preview.columns)
        header_idx = 1 if has_unnamed else 0
        
        if sheet_name:
            df = pd.read_excel(temp_path, sheet_name=sheet_name, header=header_idx)
        else:
            df = pd.read_excel(temp_path, sheet_name=0, header=header_idx)
            
        if df.empty:
            print(f"[ERROR] Excel file is empty: {file_path}")
            sys.exit(1)
            
        df.columns = df.columns.str.strip()
        
        # Mapping old required columns to new format
        if 'Start Date' not in df.columns:
            if 'Sprint Start' in df.columns:
                df['Start Date'] = df['Sprint Start']
            elif 'Effective Sprint Start' in df.columns:
                df['Start Date'] = df['Effective Sprint Start']
                
        if 'End Date' not in df.columns:
            if 'Sprint End' in df.columns:
                df['End Date'] = df['Sprint End']
            elif 'Effective Sprint End' in df.columns:
                df['End Date'] = df['Effective Sprint End']
                
        if 'APP' not in df.columns and 'Application' in df.columns:
            df['APP'] = df['Application']
            
        # Read Sheet2 for emails
        try:
            df_emails = pd.read_excel(temp_path, sheet_name="Sheet2")
            df_emails.columns = df_emails.columns.str.strip()
            df_emails['Name'] = df_emails['Name'].astype(str).str.replace(r'\xa0', ' ', regex=True).str.strip()
            email_map = dict(zip(df_emails['Name'], df_emails['Email']))
        except Exception as e:
            print(f"[WARNING] Could not read Sheet2 for emails: {e}")
            email_map = {}
            
        # Normalize Responsible PM
        if 'Responsible PM' in df.columns:
            df['Responsible PM'] = df['Responsible PM'].astype(str).str.replace(r'\xa0', ' ', regex=True).str.strip()
            df['Email/Teams'] = df['Responsible PM'].map(email_map)
            df['Name'] = df['Responsible PM']
        
        df = df.fillna("")
        
        return df.to_dict('records')
    except Exception as e:
        print(f"[ERROR] Could not read Excel file: {e}")
        sys.exit(1)
    finally:
        # Clean up temporary file
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception as e:
            print(f"[WARNING] Could not remove temp file {temp_path}: {e}")
