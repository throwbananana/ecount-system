
import pandas as pd
import os

# Simulate _prepare_recon_dataframe logic for Excel
def test_invert(file_path):
    print(f"Loading {file_path}...")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error: {e}")
        return

    print("Original Head:")
    print(df[['Desc', 'Debit', 'Credit']].head(5).to_string())
    
    # Simulate map_columns_smart (which maps Debit->借方, Credit->贷方)
    # Since inspect_excel showed 'Debit' and 'Credit' columns, let's assume they are present.
    # The reconciler maps them.
    
    # Let's manually rename to simulate what happens before inversion check
    rename_map = {}
    if 'Debit' in df.columns: rename_map['Debit'] = '借方'
    if 'Credit' in df.columns: rename_map['Credit'] = '贷方'
    
    df.rename(columns=rename_map, inplace=True)
    
    print("\nAfter Rename (Simulated):")
    cols = [c for c in ['Desc', '借方', '贷方'] if c in df.columns]
    print(df[cols].head(5).to_string())
    
    # Simulate Invert Logic
    print("\nApplying Invert...")
    if "借方" in df.columns and "贷方" in df.columns:
        df["Temp"] = df["借方"]
        df["借方"] = df["贷方"]
        df["贷方"] = df["Temp"]
        del df["Temp"]
    
    print("\nAfter Invert:")
    print(df[cols].head(5).to_string())
    
    # Check a specific row (e.g. Commission)
    # In original: Commission was Debit 1.05
    # Inverted: Commission should be Credit 1.05 (Company View: Money Out = Credit Bank)
    
    row = df[df['Desc'].str.contains('COMISION', na=False)].head(1)
    if not row.empty:
        print("\nCommission Row (Inverted):")
        print(row[cols].to_string())
        
        d = row['借方'].values[0]
        c = row['贷方'].values[0]
        print(f"Debit: {d}, Credit: {c}")
        if c > 0 and d == 0:
            print("RESULT: Commission is CREDIT (Correct for Company View 'Money Out').")
        else:
            print("RESULT: Commission is NOT Credit (Unexpected).")

test_invert(r"12yue\1122.xlsx")
