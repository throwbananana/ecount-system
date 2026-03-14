
def normalize_header(s):
    if s is None: return ""
    return str(s).lower().replace(" ", "")

def guess_column(columns, keywords):
    norm_keywords = [normalize_header(k) for k in keywords]
    for col in columns:
        norm_col = normalize_header(col)
        for nk in norm_keywords:
            if nk and nk in norm_col:
                return col
    return None

# Actual columns from your source file (Mojibake)
source_columns = ['鏃ユ湡FECHA', '鏂囦欢DOCUM', '鍐呭DESCRIPCION', '瀛樻DEPOS', '鏀粯PAGAR', '缁撳瓨BALANCE', '鍙楁浜篜AGARSE ORDEN']

# Keywords from the application code (亿看智能识别系统.py)
amount_keywords = ["金额", "amount", "本币", "金额本币", "原币", "balance"]
debit_keywords = ["借方", "借", "debit"]
credit_keywords = ["贷方", "贷", "credit"]

print("--- Simulation of Auto-Guessing ---")
guessed_amount = guess_column(source_columns, amount_keywords)
guessed_debit = guess_column(source_columns, debit_keywords)
guessed_credit = guess_column(source_columns, credit_keywords)

print(f"Guessed Amount Column: '{guessed_amount}'")
print(f"Guessed Debit Column:  '{guessed_debit}'")
print(f"Guessed Credit Column: '{guessed_credit}'")

if guessed_amount and "BALANCE" in guessed_amount:
    print("\n[CONCLUSION] The system matched 'BALANCE' as the Amount column.")
    print("This means it tried to match your Transaction Amount against the Bank Balance.")
    print("This explains why nothing matched and why Direction (Debit/Credit) was ignored.")
elif not guessed_debit and not guessed_credit:
    print("\n[CONCLUSION] The system failed to identify Debit/Credit columns.")
    print("It likely used the wrong Amount column or found no Amount at all.")
