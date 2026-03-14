import pandas as pd
import sys
import os

# Add current directory to path
sys.path.append(os.getcwd())

try:
    from bank_parser import BankParser
except ImportError:
    print("Error: Could not import BankParser")
    sys.exit(1)

def fix_mojibake(text):
    try:
        # Try to encode as ISO-8859-1 (Latin1) and decode as UTF-8
        # Common scenario: UTF-8 bytes displayed as Latin1
        # BUT here we saw '鏃ユ湡' which is GBK interpretation of UTF-8 bytes
        # '日期' (E6 97 A5 E6 9C 9F)
        # interpreted as GBK: E697=鏃, A5E6=ユ, 9C9F=湡
        return text.encode('gbk').decode('utf-8')
    except:
        return text

def analyze_file(path, label):
    print(f"\n--- Analyzing {label}: {path} ---")
    if not os.path.exists(path):
        print("File not found.")
        return None

    df = None
    if path.lower().endswith('.pdf'):
        df = BankParser.parse_pdf(path)
    else:
        df = pd.read_excel(path)
    
    if df is None or df.empty:
        print("Empty DataFrame.")
        return None

    # Check Columns for Mojibake
    print("Original Columns:", df.columns.tolist())
    fixed_cols = [fix_mojibake(str(c)) for c in df.columns]
    print("Attempted Fix Columns:", fixed_cols)
    
    # Check Date Range
    date_col = None
    # Guess date col
    for c in df.columns:
        if 'date' in str(c).lower() or '日期' in str(c) or 'fecha' in str(c).lower():
            date_col = c
            break
    
    if not date_col:
        # Try fixed columns
        for c, fc in zip(df.columns, fixed_cols):
            if 'date' in str(fc).lower() or '日期' in str(fc) or 'fecha' in str(fc).lower():
                date_col = c
                break
    
    if date_col:
        print(f"Detected Date Column: {date_col}")
        # Try to parse dates
        dates = pd.to_datetime(df[date_col], errors='coerce')
        valid_dates = dates.dropna()
        if not valid_dates.empty:
            print(f"Date Range: {valid_dates.min()} to {valid_dates.max()}")
            print(f"Sample Dates: {valid_dates.head(3).tolist()}")
        else:
            print("Could not parse dates in column.")
            print("Raw samples:", df[date_col].head(3).tolist())
    else:
        print("No Date Column Detected.")

    return df

analyze_file(r"12yue\1122.xlsx", "Target (1122.xlsx)")
analyze_file(r"12yue\12月st摘要.xlsx", "Source 1 (ST Summary)")
analyze_file(r"12yue\AccountStatementMovementsDetail (1).pdf", "Source 2 (PDF)")
