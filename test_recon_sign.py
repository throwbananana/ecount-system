
import pandas as pd
from reconciliation_module import StandardReconciler

def test_opposite_sign():
    reconciler = StandardReconciler()
    
    # 1. Local (Debit 1.5 -> Amount 1.5)
    df_local = pd.DataFrame({
        '凭证日期': ['2025-11-01'],
        '序号': ['70701'],
        '摘要': ['Test'],
        '往来单位编码': ['BANK'],
        '金额': [1.5],
        '类型': ['3'] # Debit
    })
    
    # 2. Yikan (Credit 1.5 -> Amount -1.5)
    df_yikan = pd.DataFrame({
        '凭证日期': ['2025-11-01'],
        '序号': ['70701'],
        '摘要': ['Test'],
        '往来单位编码': ['BANK'],
        '金额': [-1.5],
        '类型': ['4'] # Credit
    })
    
    parsed_local = reconciler.parse_standard_df(df_local, 'Local')
    parsed_yikan = reconciler.parse_standard_df(df_yikan, 'Yikan')
    
    # Fix mapped code for test
    parsed_local['Mapped_Code'] = 'BANK'
    
    print("Local Amt:", parsed_local.iloc[0]['Amount'])
    print("Yikan Amt:", parsed_yikan.iloc[0]['Amount'])
    
    config = {
        'start_date': None,
        'end_date': None,
        'fuzzy_code': False
    }
    
    print("Running reconcile...")
    results = reconciler.reconcile(df_local, df_yikan, config)
    
    print("Matched Rows:", len(results['matched']))
    if not results['matched'].empty:
        print("Match Reason:", results['matched'].iloc[0]['匹配原因'])

if __name__ == "__main__":
    test_opposite_sign()
