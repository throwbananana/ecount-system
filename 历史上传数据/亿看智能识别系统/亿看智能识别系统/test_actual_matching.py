
import pandas as pd
from datetime import datetime, date, timedelta
import sys
import os

# Import the necessary classes from the main script
# Since we are in the same directory, we can try to import or mock the logic
from bank_parser import BankParser
from 亿看智能识别系统 import score_similarity

def test_actual_match():
    # 1. Load Data
    pdf_path = r"12yue\AccountStatementMovementsDetail (2).pdf"
    excel_path = r"12yue\12月st摘要.xlsx"
    
    print(f"--- 1. Testing PDF Parsing: {pdf_path} ---")
    df_source = BankParser.parse_pdf(pdf_path)
    if df_source is None or df_source.empty:
        print("Error: PDF parsing failed or returned empty data.")
        return
    
    print(f"PDF parsed: {len(df_source)} rows.")
    print(df_source.head(5))

    print(f"\n--- 2. Loading Target Excel: {excel_path} ---")
    if not os.path.exists(excel_path):
        print(f"Error: Excel file {excel_path} not found.")
        return
    df_target = pd.read_excel(excel_path)
    print(f"Excel loaded: {len(df_target)} rows.")
    print(df_target.head(5))

    # Mock settings
    date_tol_days = 3
    amount_abs_tol = 0.01
    amount_pct_tol = 0.01
    
    # 3. Simulate Matching Logic (Simplified version of _run_summary_match)
    matched_count = 0
    results = []
    
    # Pre-process source
    source_entries = []
    for idx, row in df_source.iterrows():
        # Handle date parsing similar to the main app
        d_val = row['Date']
        try:
            # Bank parser usually returns strings like "12/11/2025" or "12-Dec-2025"
            s_date = pd.to_datetime(d_val).date()
        except:
            continue
            
        source_entries.append({
            "date": s_date,
            "amount": abs(row['Debit']) if row['Debit'] > 0 else abs(row['Credit']),
            "direction": "debit" if row['Debit'] > 0 else "credit",
            "summary": str(row['Desc'])
        })

    print("\n--- 3. Running Match Simulation ---")
    for t_idx, t_row in df_target.iterrows():
        # Guessing target columns based on file content
        t_date_val = t_row.get('日期') or t_row.get('Date')
        t_amt_val = t_row.get('金额') or t_row.get('Amount') or t_row.get('Debit') or t_row.get('Credit')
        t_desc = str(t_row.get('摘要') or t_row.get('Desc') or "")
        
        try:
            t_date = pd.to_datetime(t_date_val).date()
            # If 2-digit year fix is needed (as applied in main app)
            if t_date.year < 100: t_date = t_date.replace(year=t_date.year + 2000)
            
            t_amount = abs(float(t_amt_val))
        except:
            continue
            
        best = None
        for entry in source_entries:
            # Date check
            if abs((t_date - entry['date']).days) > date_tol_days:
                continue
            
            # Amount check
            diff = abs(t_amount - entry['amount'])
            if diff > amount_abs_tol:
                continue
            
            # Similarity
            score = score_similarity(t_desc, entry['summary'], t_desc)
            
            if best is None or score > best['score']:
                best = {"entry": entry, "score": score}
        
        if best and best['score'] > 0.1:
            matched_count += 1
            results.append({
                "Target_Row": t_idx + 2,
                "Date": t_date,
                "Amount": t_amount,
                "Old_Desc": t_desc,
                "New_Desc": best['entry']['summary'],
                "Direction": best['entry']['direction'],
                "Score": best['score']
            })

    print(f"\nMatch Results: Successfully matched {matched_count} out of {len(df_target)} rows.")
    for res in results[:10]:
        print(f"Row {res['Target_Row']}: {res['Date']} | Amt: {res['Amount']} | Dir: {res['Direction']}")
        print(f"  Old: {res['Old_Desc'][:50]}...")
        print(f"  New: {res['New_Desc'][:50]}...")
        print("-" * 30)

if __name__ == "__main__":
    test_actual_match()
