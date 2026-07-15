
import pandas as pd
import os
from datetime import datetime

# Standard Voucher Columns based on 亿看智能识别系统.py
VOUCHER_COLUMNS = [
    "凭证日期", "序号", "会计凭证No.", "摘要编码", "摘要", 
    "类型", "科目编码", "对方科目", "默认账户", 
    "往来单位编码", "往来单位名", "金额", "外币金额", 
    "汇率", "部门", "币种"
]

def format_date(d):
    try:
        if pd.isna(d): return ""
        if isinstance(d, str):
            # Try parsing various formats
            for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%b-%Y"]:
                try:
                    return datetime.strptime(d, fmt).strftime("%Y%m%d")
                except: continue
            return d.replace("-","").replace("/","")[:8]
        return d.strftime("%Y%m%d")
    except:
        return ""

def export_to_voucher(df_add, output_path=None):
    """
    Converts the 'Should Add' dataframe (Local data missing in Yikan) 
    into the standard Voucher Import format.
    """
    if df_add.empty:
        print("No data to export.")
        return

    # Prepare target dataframe
    voucher_data = []
    
    # Iterate through df_add (Source: Local System)
    # Expected cols: '当地日期', '当地单号', '当地编码', '当地映射编码', '当地借方', '当地贷方', '当地摘要'
    # Note: Column names might be different depending on where this is called.
    # In the GUI, they are renamed for the report. 
    # Let's assume the input df_add has the report columns:
    # '当地日期' (Local Date), '当地摘要' (Desc), '当地映射编码' (Code), '当地借方' (Debit), '当地贷方' (Credit)
    
    for _, row in df_add.iterrows():
        # 1. Date
        p_date = format_date(row.get('当地日期', ''))
        
        # 2. Amount & Type
        debit = row.get('当地借方', 0)
        credit = row.get('当地贷方', 0)
        
        # Determine Amount and Direction
        # Type: 1=Out (Credit?), 2=In (Debit?), 3=Debit, 4=Credit
        # Usually: 3=Debit (借), 4=Credit (贷) is safest for general vouchers
        
        amt = 0
        v_type = ""
        
        if debit > 0 and credit == 0:
            amt = debit
            v_type = "3" # Debit
        elif credit > 0 and debit == 0:
            amt = credit
            v_type = "4" # Credit
        elif debit > credit:
            amt = debit - credit
            v_type = "3"
        else:
            amt = credit - debit
            v_type = "4"
            
        if amt == 0: continue # Skip zero rows
        
        # 3. Summary
        desc = str(row.get('当地摘要', ''))
        doc_no = str(row.get('当地单号', ''))
        full_desc = f"{desc} {doc_no}".strip()
        
        # 4. Partner/Subject
        code = str(row.get('当地映射编码', '')).strip()
        
        # Create Row
        v_row = {col: "" for col in VOUCHER_COLUMNS}
        v_row["凭证日期"] = p_date
        v_row["摘要"] = full_desc
        v_row["类型"] = v_type
        v_row["金额"] = amt
        v_row["往来单位编码"] = code # Map Code to Partner Code
        # v_row["科目编码"] = "" # Leave empty for smart recognition or manual fill?
        
        # Defaults
        v_row["序号"] = "1"
        v_row["汇率"] = "1" 
        
        voucher_data.append(v_row)
        
    df_voucher = pd.DataFrame(voucher_data)
    
    # Save
    if not output_path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"补录凭证_{timestamp}.xlsx"
        
    df_voucher.to_excel(output_path, index=False)
    print(f"Exported {len(df_voucher)} rows to {output_path}")
    return output_path

if __name__ == "__main__":
    # Test with dummy data
    data = {
        '当地日期': ['2025-10-01', '2025-10-02'],
        '当地单号': ['DOC001', 'DOC002'],
        '当地映射编码': ['C001', 'C002'],
        '当地借方': [100.50, 0],
        '当地贷方': [0, 200.00],
        '当地摘要': ['Test Debit', 'Test Credit']
    }
    df = pd.DataFrame(data)
    export_to_voucher(df, "test_voucher_export.xlsx")
