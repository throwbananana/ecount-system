import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import re
import os

from bank_parser import BankParser

def score_similarity(s1, s2, context=""):
    s1, s2 = str(s1).lower(), str(s2).lower()
    if not s1 or not s2: return 0.0
    common = set(s1) & set(s2)
    return len(common) / max(len(s1), len(s2))

def parse_date_fixed(value, dayfirst=True):
    if pd.isna(value): return None
    s_val = str(value).strip()
    try:
        dt = pd.to_datetime(s_val, dayfirst=dayfirst, errors='coerce')
        if pd.notna(dt):
            y = dt.year
            if y < 100: y += 2000
            return date(y, dt.month, dt.day)
    except: pass
    return None

def find_col(df, keywords):
    for col in df.columns:
        c_str = str(col).upper()
        for k in keywords:
            if k.upper() in c_str:
                return col
    return None

def perform_match():
    # 1. 加载 PDF
    pdf_path = r"12yue\AccountStatementMovementsDetail (2).pdf"
    df_source_raw = BankParser.parse_pdf(pdf_path)
    source_entries = []
    for _, row in df_source_raw.iterrows():
        d = parse_date_fixed(row['Date'], dayfirst=True)
        if not d: continue
        amt = row['Debit'] if row['Debit'] > 0 else row['Credit']
        direc = "debit" if row['Debit'] > 0 else "credit"
        source_entries.append({
            "date": d, "amount": amt, "direction": direc, "summary": str(row['Desc'])
        })

    # 2. 加载 Excel
    excel_path = r"12yue\12月st摘要.xlsx"
    df_target = pd.read_excel(excel_path)
    
    col_date = find_col(df_target, ["日期", "FECHA"])
    col_summary = find_col(df_target, ["内容", "DESCRIPCION"])
    col_debit = find_col(df_target, ["支付", "PAGAR"])
    col_credit = find_col(df_target, ["存入", "DEPOSITO"])

    print(f"检测到列: 日期={col_date}, 摘要={col_summary}, 借={col_debit}, 贷={col_credit}")

    matched_count = 0
    df_result = df_target.copy()
    
    for idx, row in df_target.iterrows():
        t_date = parse_date_fixed(row[col_date], dayfirst=True)
        d_val = pd.to_numeric(row[col_debit], errors='coerce') if col_debit else 0
        c_val = pd.to_numeric(row[col_credit], errors='coerce') if col_credit else 0
        
        d_val = 0 if pd.isna(d_val) else d_val
        c_val = 0 if pd.isna(c_val) else c_val
        
        if abs(d_val) > 0.001:
            t_amt, t_dir = abs(d_val), "debit"
        elif abs(c_val) > 0.001:
            t_amt, t_dir = abs(c_val), "credit"
        else: continue

        if not t_date: continue

        best = None
        for s in source_entries:
            if abs((t_date - s['date']).days) > 3: continue
            if abs(t_amt - s['amount']) > 0.01: continue
            
            score = score_similarity(str(row[col_summary]), s['summary'])
            if best is None or score > best['score']:
                best = {"entry": s, "score": score}
        
        if best:
            matched_count += 1
            df_result.at[idx, col_summary] = best['entry']['summary']
            s_dir = best['entry']['direction']
            s_amt = best['entry']['amount']
            if s_dir == "debit" and col_debit:
                df_result.at[idx, col_debit] = s_amt
                if col_credit: df_result.at[idx, col_credit] = 0
            elif s_dir == "credit" and col_credit:
                df_result.at[idx, col_credit] = s_amt
                if col_debit: df_result.at[idx, col_debit] = 0

    output_path = r"12yue\12月st摘要_匹配结果_预览.xlsx"
    df_result.to_excel(output_path, index=False)
    print(f"成功匹配 {matched_count} 行，结果已保存。")

if __name__ == "__main__":
    perform_match()