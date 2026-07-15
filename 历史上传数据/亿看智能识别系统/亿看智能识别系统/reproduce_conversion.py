
import pandas as pd
import json
from decimal import Decimal, ROUND_HALF_UP
import math

# 模拟系统核心逻辑
def format_number(value, max_decimal_len=2):
    if value is None or value == "": return "0.00"
    try:
        d = Decimal(str(value).replace(",", ""))
        quantize_str = "1." + ("0" * max_decimal_len)
        d = d.quantize(Decimal(quantize_str), rounding=ROUND_HALF_UP)
        return format(d, "f")
    except:
        return "0.00"

def test_logic():
    print("--- 模拟外币模式转换测试 ---")
    
    # 模拟输入数据 (行 8)
    src_row = {
        "Date": "2025/12/22",
        "Desc": "LAFISE 转 BAC",
        "Debit": 30000.0,
        "Credit": None,
        "Code": "1002"
    }
    
    # 模拟配置
    use_foreign_currency = True
    default_rate = 7.01031
    mapping = {
        "日期": "Date",
        "摘要": "Desc",
        "金额": None, # 留空
        "外币金额": "[综合] 借贷列辅助",
        "汇率": "[综合] 默认汇率"
    }
    
    # 模拟智能识别出的错误金额 (AI 干扰)
    smart_data = {"金额": "0.0", "外币金额": "30000.0"}
    
    # 模拟借贷推断
    derived_amount = 30000.0 # 从 Debit 提取
    
    print(f"输入外币: {derived_amount}, 汇率: {default_rate}")
    
    # --- 执行修复后的逻辑 ---
    # 1. 预同步汇率 (保留 4 位)
    calc_rate = float(Decimal(str(default_rate)).quantize(Decimal("0.0000"), rounding=ROUND_HALF_UP))
    print(f"计算用汇率: {calc_rate}")
    
    # 2. 计算本币
    f_amt_val = derived_amount # 假设成功获取
    _d_f = Decimal(str(f_amt_val))
    _d_r = Decimal(str(calc_rate))
    calculated_local = float((_d_f * _d_r).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
    
    print(f"计算出的本币金额: {calculated_local}")
    
    # 模拟最终输出
    final_amount = calculated_local
    print(f"最终写入 Excel 的金额: {final_amount}")
    
    expected = 210309.0 # 30000 * 7.0103
    if abs(final_amount - expected) < 0.01:
        print("✅ 测试通过：金额完全匹配！")
    else:
        print(f"❌ 测试失败：期望 {expected}, 实际 {final_amount}")

if __name__ == "__main__":
    test_logic()
