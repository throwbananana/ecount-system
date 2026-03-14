import pandas as pd
import os
from datetime import datetime

def fix_mojibake(text):
    if not isinstance(text, str): return text
    try:
        # Reversing the common UTF-8 -> GBK/Latin1 error
        return text.encode('gbk').decode('utf-8')
    except:
        # Manual mapping for common garbled strings found in your files
        mapping = {
            "鏃ユ湡": "日期",
            "鍐呭": "内容",
            "瀛樻": "存款",
            "鏀粯": "支付",
            "缁撳瓨": "结存",
            "鏂囦欢": "文件",
            "鍙楁浜": "收款人",
            "鍘熸憳瑕": "原摘要"
        }
        for k, v in mapping.items():
            if k in text:
                text = text.replace(k, v)
        return text

def standardize_date(df, date_col, dayfirst=False):
    if date_col not in df.columns: return df
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=dayfirst, errors='coerce')
    # Fill remaining NaT with empty string or handle as needed
    return df

# File Paths
source_path = r'12yue\12月st摘要.xlsx'
target_path = r'12yue\1122.xlsx'

source_fixed_path = r'12yue\12月st摘要_已修复.xlsx'
target_fixed_path = r'12yue\1122_已修复.xlsx'

print("--- Starting Repair ---")

# 1. Fix Source File
print(f"Repairing Source: {source_path}")
df_s = pd.read_excel(source_path)
df_s.columns = [fix_mojibake(c) for c in df_s.columns]
# Source file uses DD-MM-YYYY (01-12-2025 is Dec 1)
df_s = standardize_date(df_s, '日期FECHA', dayfirst=True)
df_s.to_excel(source_fixed_path, index=False)

# 2. Fix Target File
print(f"Repairing Target: {target_path}")
df_t = pd.read_excel(target_path)
df_t.columns = [fix_mojibake(c) for c in df_t.columns]
# Target file uses MM/DD/YY (12/11/25 is Dec 11)
df_t = standardize_date(df_t, 'Date', dayfirst=False)
df_t.to_excel(target_fixed_path, index=False)

print("\n--- Repair Complete ---")
print(f"Fixed Source saved to: {source_fixed_path}")
print(f"Fixed Target saved to: {target_fixed_path}")
print("\nInstructions:")
print("1. Open '亿看智能识别系统'")
print("2. Choose '1122_已修复.xlsx' as Target")
print("3. Choose '12月st摘要_已修复.xlsx' as Source in '摘要匹配' tab")
print("4. The columns should now auto-match correctly (日期FECHA, 存款DEPOS, 支付PAGAR).")
