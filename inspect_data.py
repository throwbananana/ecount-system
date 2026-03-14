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

def inspect_excel(path):
    print(f"\n--- Inspecting {path} ---")
    if not os.path.exists(path):
        print("File not found.")
        return
    try:
        df = pd.read_excel(path)
        print("Columns:", df.columns.tolist())
        print("First 5 rows:")
        print(df.head())
        print("Data Types:")
        print(df.dtypes)
    except Exception as e:
        print(f"Error reading Excel: {e}")

def inspect_pdf(path):
    print(f"\n--- Inspecting {path} ---")
    if not os.path.exists(path):
        print("File not found.")
        return
    try:
        # Use the same logic as the main app
        df = BankParser.parse_pdf(path, use_ocr=False)
        if df is None or df.empty:
            print("PDF Parse Result: Empty or None")
            # Try OCR if standard fails, just to see
            print("Retrying with OCR...")
            df = BankParser.parse_pdf(path, use_ocr=True)
        
        if df is not None and not df.empty:
            print("Columns:", df.columns.tolist())
            print("First 5 rows:")
            print(df.head())
            print("Data Types:")
            print(df.dtypes)
        else:
            print("Still empty after OCR.")
    except Exception as e:
        print(f"Error parsing PDF: {e}")

file1 = r"12yue\1122.xlsx"
file2 = r"12yue\12月st摘要.xlsx"
file3 = r"12yue\AccountStatementMovementsDetail (1).pdf"

inspect_excel(file1)
inspect_excel(file2)
inspect_pdf(file3)