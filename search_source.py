import pandas as pd

source_path = r'12yue\12月st摘要.xlsx'
target_path = r'12yue\1122.xlsx'

df_source = pd.read_excel(source_path)
df_target = pd.read_excel(target_path)

# Helper to find columns
def find_col(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in str(col):
                return col
    return None

s_cols = {
    "summary": find_col(df_source, ["DESCRIPCION", "内容"]),
    "debit": find_col(df_source, ["DEPOS", "存款"]), 
    "credit": find_col(df_source, ["PAGAR", "支付"])
}

print(f"Source Columns: {s_cols}")

target_val = 1.05
print(f"Searching for {target_val} in Source...")

for idx, row in df_source.iterrows():
    d = row.get(s_cols["debit"])
    c = row.get(s_cols["credit"])
    try: d = float(d) 
    except: d = 0
    try: c = float(c)
    except: c = 0
    
    if abs(d - target_val) < 0.001 or abs(c - target_val) < 0.001:
        print(f"MATCH FOUND at Row {idx}")
        print(row)
        print("-" * 20)
