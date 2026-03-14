
import pandas as pd

def inspect_report():
    report_path = "智能对账报告_20251220_162440.xlsx"
    print(f"Inspecting Report: {report_path}")
    
    try:
        xls = pd.ExcelFile(report_path)
        print("Sheet Names:", xls.sheet_names)
        
        if '当地未匹配(建议新增)' in xls.sheet_names:
            df_l = pd.read_excel(xls, '当地未匹配(建议新增)')
            print(f"\nUnmatched Local Rows: {len(df_l)}")
            if not df_l.empty:
                # Print relevant columns for matching
                # Standard format usually: 凭证日期, 序号, 摘要, 金额
                cols = [c for c in df_l.columns if c in ['凭证日期', '序号', '摘要', '金额', '借方', '贷方']]
                print(df_l[cols].head(5).to_string())
                
        if '亿看未匹配(建议清理)' in xls.sheet_names:
            df_y = pd.read_excel(xls, '亿看未匹配(建议清理)')
            print(f"\nUnmatched Yikan Rows: {len(df_y)}")
            if not df_y.empty:
                cols = [c for c in df_y.columns if c in ['凭证日期', '序号', '摘要', '金额', '借方', '贷方']]
                print(df_y[cols].head(5).to_string())

    except Exception as e:
        print(f"Error reading report: {e}")

if __name__ == "__main__":
    inspect_report()
