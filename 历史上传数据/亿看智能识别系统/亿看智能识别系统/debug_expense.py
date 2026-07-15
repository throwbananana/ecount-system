
import pandas as pd
import os
import re
from datetime import datetime

path = r"C:\Users\123\Downloads\亿看智能识别系统\基础资料\费用1-11月.xlsx"

def _parse_header_date(df):
    if df.empty:
        return None, None
    try:
        header_str = str(df.columns[0])
        if isinstance(df.columns[0], int):
            header_str = str(df.iloc[0, 0])
    except:
        return None, None

    date_pattern = r'(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})\s*~\s*(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})'
    match = re.search(date_pattern, header_str)
    
    if match:
        print(f"Match found: {match.groups()}")
        return True
    return False

print(f"--- Debugging {path} ---")
if not os.path.exists(path):
    print("File not found!")
    exit()

# 1. Peek Header
try:
    df_peek = pd.read_excel(path, header=None, nrows=1)
    print("Row 0 content:", df_peek.iloc[0].values)
    is_date_range = _parse_header_date(df_peek)
    print(f"Header Date Range Detected: {is_date_range}")
except Exception as e:
    print(f"Peek error: {e}")

# 2. Read with header=1
try:
    df = pd.read_excel(path, header=1)
    print("\nColumns with header=1:")
    print(df.columns.tolist())
    
    date_col = None
    for c in df.columns:
        if '日期' in str(c) or 'Date' in str(c):
            date_col = c
            break
    
    print(f"\nIdentified Date Column: {date_col}")
    
    if date_col:
        print("\nFirst 5 values in Date Column:")
        print(df[date_col].head(5).tolist())
        
        # Test Parsing
        def parse_date(val):
            s = str(val).split('-')[0].strip()
            try:
                return pd.to_datetime(s)
            except:
                return None
        
        parsed = df[date_col].apply(parse_date)
        print("\nParsed Dates head:")
        print(parsed.head(5))
        
        print("\nUnique Months found:")
        df['P'] = parsed
        print(df['P'].dt.strftime('%Y-%m').unique())
        
except Exception as e:
    print(f"Read error: {e}")
