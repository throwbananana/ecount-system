
import pandas as pd
from reconciliation_module import StandardReconciler

def diagnose_mapping_issue():
    print("=== Mapping Diagnosis ===")
    
    try:
        # Load raw file to check content
        df_yikan_raw = pd.read_excel("11月bac对账建议.xlsx")
        print("Raw Columns:", list(df_yikan_raw.columns))
        print("\n--- Raw Data Sample (First 3 rows) ---")
        cols_of_interest = [c for c in df_yikan_raw.columns if '金额' in c or '借' in c or '贷' in c]
        print(df_yikan_raw[cols_of_interest].head(3).to_string())
    except Exception as e:
        print(f"Error loading raw file: {e}")
        return

    # Simulate the User's "Bad" Mapping
    # User Mapped: '金额': '外币借方金额', '借方': '外币借方金额', '贷方': '外币贷方金额'
    mapping = {
        '凭证日期': '日期-号码',
        '序号': '债权债务号码',
        '摘要': '摘要',
        '往来单位编码': '摘要',
        '金额': '外币借方金额', # <--- SUSPECTED ISSUE
        '借方': '外币借方金额',
        '贷方': '外币贷方金额'
    }
    
    print("\n--- Applying User Mapping ---")
    df_mapped = pd.DataFrame()
    for k, v in mapping.items():
        if v in df_yikan_raw.columns:
            df_mapped[k] = df_yikan_raw[v]
            
    print("Mapped DataFrame Sample:")
    print(df_mapped[['金额', '借方', '贷方']].head(3).to_string())
    
    # Parse using Reconciler
    rec = StandardReconciler()
    # We need to hack parse_standard_df slightly or just call it, 
    # but parse_standard_df uses '外币金额' if present. 
    # In user mapping, '外币金额' is NOT mapped. So it uses '金额'.
    
    parsed = rec.parse_standard_df(df_mapped, 'Yikan')
    
    print("\n--- Parsed Result (Internal 'Amount') ---")
    print(parsed[['Debit', 'Credit', 'Amount']].head(5).to_string())
    
    # Check for Credit entries (where Debit should be 0)
    print("\n--- Checking Credit Entries ---")
    credits = parsed[parsed['Credit'] > 0]
    if not credits.empty:
        print(credits[['Debit', 'Credit', 'Amount']].head(3).to_string())
        
        # Verify if Amount matches Credit
        # If User mapped '金额' -> '外币借方金额', then for a Credit row:
        # Raw Debit = 0. Raw Credit = 100.
        # Mapped '金额' = 0.
        # Parsed 'Amount' might take '金额' (0).
        
        row0 = credits.iloc[0]
        if row0['Amount'] == 0:
            print("\n[CRITICAL] Found the bug! Amount is 0 for Credit entries.")
            print("Reason: '金额' was mapped to '外币借方金额', which is 0 for credits.")
            print("Fix: Unmap '金额' (set to Not Used) and only map '借方' and '贷方'.")
    else:
        print("No Credit entries found in sample?")

if __name__ == "__main__":
    diagnose_mapping_issue()
