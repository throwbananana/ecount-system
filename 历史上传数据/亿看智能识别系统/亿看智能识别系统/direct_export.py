
import sys
import os
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

# 强制加载主程序
import 亿看智能识别系统 as main_app

# 模拟 GUI 环境
class MockGUI:
    def __init__(self):
        self.use_foreign_currency_var = type('obj', (), {'get': lambda *a: True})()
        self.enable_smart_recognition = type('obj', (), {'get': lambda *a: False})()
        self.auto_balance_var = type('obj', (), {'get': lambda *a: True})()
        self.split_amount_var = type('obj', (), {'get': lambda *a: False})()
        self.use_ai_var = type('obj', (), {'get': lambda *a: False})()
        
        # 模拟手动映射借贷列
        self.manual_debit_col_var = type('obj', (), {'get': lambda *a: "Debit"})()
        self.manual_credit_col_var = type('obj', (), {'get': lambda *a: "Credit"})()
        self.manual_dc_col_var = type('obj', (), {'get': lambda *a: ""})()
        
        self.default_values = {
            "汇率": "7.01031",
            "部门": "10001",
            "科目编码": "1002",
            "默认账户往来": ""
        }
        self.field_formats = {}
        self.summary_recognizer = None
        self.base_data_mgr = None

    def log_message(self, msg): print(f"[LOG] {msg}")
    def _debug_log(self, msg): print(f"[DEBUG] {msg}")

def run_fix_export():
    print("🚀 [修正测试 V2] 正在执行通用凭证直接导出...")
    
    source_file = r"12yue\1122.xlsx"
    output_file = r"12yue\1122_修复版导出.xlsx"
    
    # 1. 加载数据
    df = pd.read_excel(source_file)
    gui = MockGUI()
    gui.input_df = df
    gui.input_columns = df.columns.tolist()
    
    # 手动定义通用凭证表头 (按 Template_通用凭证.xlsx 精确顺序)
    headers = ["日期", "序号", "摘要", "科目编码", "对方科目", "默认账户", "往来单位编码", "往来单位名", "金额", "外币金额", "汇率", "部门", "类型", "摘要编码", "会计凭证No."]
    
    # 建立表头到索引的映射
    h_map = {h: i for i, h in enumerate(headers)}
    mapping = {
        "日期": "Date",
        "摘要": "Desc",
        "科目编码": "Code",
        "外币金额": "[综合] 借贷列辅助",
        "汇率": "[综合] 默认汇率",
        "部门": "[综合] 默认部门",
        "类型": "[综合] 借贷列辅助"
    }

    output_rows = []
    serial_map = {}
    next_serial_id = 1
    
    for idx, src_row in df.iterrows():
        # A. 借贷推断 (直接使用主程序逻辑)
        derived_ctx = main_app.ExcelConverterGUI._derive_debit_credit_context(gui, src_row)
        derived_amount = derived_ctx.get("derived_amount")
        derived_type = derived_ctx.get("derived_type")

        # B. 汇率预处理
        rate_str = gui.default_values.get("汇率", "1")
        # 直接按主程序 do_convert 中的精度预处理
        rate_formatted = main_app.convert_value("汇率", rate_str)
        row_calc_rate = float(rate_formatted.replace(",", ""))

        # C. 外币基数 (此时 derived_amount 应该有值了)
        row_f_amt = derived_amount
        
        # D. 行级预算 (金额 = 外币 * 汇率)
        row_theoretical_amount = None
        if row_f_amt is not None and not pd.isna(row_f_amt):
            _d_f = Decimal(str(row_f_amt))
            _d_r = Decimal(str(row_calc_rate))
            row_theoretical_amount = float((_d_f * _d_r).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))

        # E. 构建行数据
        out_row_data = {}
        for h in headers:
            val = None
            if h == "日期": val = src_row.get("Date")
            elif h == "摘要": val = src_row.get("Desc")
            elif h == "科目编码": val = src_row.get("Code") or "1002"
            elif h == "外币金额": val = row_f_amt
            elif h == "汇率": val = rate_str
            elif h == "部门": val = "10001"
            elif h == "类型": val = derived_type
            elif h == "金额": val = row_theoretical_amount
            elif h == "序号":
                s_val = str(src_row.get("Doc") if not pd.isna(src_row.get("Doc")) else (idx // 2 + 1))
                if s_val not in serial_map:
                    serial_map[s_val] = str(next_serial_id)
                    next_serial_id += 1
                val = serial_map[s_val]
            
            out_row_data[h] = main_app.convert_value(h, val)

        # F. 添加主行
        row_list = [out_row_data[h] for h in headers]
        output_rows.append(row_list)
        
        # G. 自动平账行 (只有当计算成功时)
        if out_row_data["类型"] in ["3", "4"] and out_row_data["金额"]:
            balance_row = list(row_list)
            balance_row[12] = "4" if row_list[12] == "3" else "3"
            balance_row[3] = "1122" # 对方科目
            
            # 平账行的外币也要同步计算
            b_amt = float(balance_row[8])
            b_rate = float(balance_row[10])
            if b_rate > 0:
                balance_row[9] = format_number_simple(b_amt / b_rate, 4)
            
            output_rows.append(balance_row)

    # 导出
    res_df = pd.DataFrame(output_rows, columns=headers)
    res_df.to_excel(output_file, index=False)
    print(f"✅ 导出成功: {output_file} (总行数: {len(output_rows)})")

def format_number_simple(val, dec):
    return format(Decimal(str(val)).quantize(Decimal("1." + "0"*dec), rounding=ROUND_HALF_UP), "f")

if __name__ == "__main__":
    run_fix_export()
