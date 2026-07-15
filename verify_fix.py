import sys, pandas as pd, numpy as np, re
sys.stdout.reconfigure(encoding='utf-8')

files = [
    '2026.5.8博胜 藤井 威普 足步一个整柜 269874839 MRSU58712 ZKP202601090(1).xlsx',
    '2026.5.8博胜 藤井一个整柜269874800 MRSU5133682 ZKP2026009(1).xlsx'
]

for file in files:
    print(f'\n{"="*60}')
    print(f'文件: {file}')
    df = pd.read_excel(file, sheet_name='报关清单')
    df = df.dropna(how='all')

    def _is_valid_container_no(v):
        if not isinstance(v, str):
            return False
        s = v.strip()
        return bool(re.search(r'[A-Za-z]', s) and re.search(r'\d', s))

    df['集装箱号'] = df['集装箱号'].apply(lambda v: v if _is_valid_container_no(v) else None)
    df['集装箱号'] = df['集装箱号'].ffill()

    container_nos = df['集装箱号'].dropna().unique()
    print(f'货柜号列表: {list(container_nos)}')

    for cn in container_nos:
        block = df[df['集装箱号'] == cn]
        cols = list(block.columns)
        
        def find_by_label(label):
            for _, row in block.iterrows():
                for col in cols:
                    val = row[col]
                    if isinstance(val, str) and label in val:
                        start = cols.index(col) + 1
                        for c2 in cols[start:]:
                            v2 = row[c2]
                            if isinstance(v2, (int, float, np.number)) and not pd.isna(v2):
                                return float(v2)
                        for c2 in cols:
                            v2 = row[c2]
                            if isinstance(v2, (int, float, np.number)) and not pd.isna(v2):
                                return float(v2)
            return None

        ins = find_by_label('保费')
        ag = find_by_label('代理费')
        rate = find_by_label('汇率')
        print(f'  货柜 {cn}:')
        print(f'    保费(USD): {ins}')
        print(f'    代理费(RMB): {ag}')
        print(f'    汇率: {rate}')
