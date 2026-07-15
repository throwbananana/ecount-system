import sys
import os
import pandas as pd

# Add current directory to path so we can import bank_parser
sys.path.append(os.getcwd())

try:
    from bank_parser import BankParser
except ImportError:
    print("Error: Could not import BankParser. Make sure bank_parser.py is in the current directory.")
    sys.exit(1)

pdf_path = r"12yue\AccountStatementMovementsDetail (2).pdf"

if not os.path.exists(pdf_path):
    print(f"Error: File not found: {pdf_path}")
    sys.exit(1)

print(f"Parsing {pdf_path}...")
try:
    df = BankParser.parse_pdf(pdf_path)
    if df is not None and not df.empty:
        print("Parse Successful.")
        print("Columns:", df.columns.tolist())
        print("-" * 30)
        # Print first 10 rows
        print(df[['Date', 'Desc', 'Debit', 'Credit', 'Balance']].head(10).to_string())
        print("-" * 30)
        
        # Check logic
        # Find a row with Debit > 0
        debits = df[df['Debit'] > 0].head(1)
        if not debits.empty:
            print("\nExample Debit (Bank Perspective?):")
            print(debits[['Date', 'Desc', 'Debit', 'Credit', 'Balance']].to_string())
            
        credits = df[df['Credit'] > 0].head(1)
        if not credits.empty:
            print("\nExample Credit (Bank Perspective?):")
            print(credits[['Date', 'Desc', 'Debit', 'Credit', 'Balance']].to_string())
            
    else:
        print("Parse returned empty DataFrame or None.")
except Exception as e:
    print(f"An error occurred: {e}")
    import traceback
    traceback.print_exc()