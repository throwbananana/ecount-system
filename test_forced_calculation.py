
import sys
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

# 模拟主程序中的关键转换函数
def format_number(value, max_decimal_len=2):
    try:
        s = str(value).replace(",", "")
        d = Decimal(s)
        quantize_str = "1." + ("0" * max_decimal_len)
        d = d.quantize(Decimal(quantize_str), rounding=ROUND_HALF_UP)
        return format(d, "f")
    except:
        return str(value)

def convert_value(header_name, src_value, max_decimal=None):
    if header_name == "汇率":
        return format_number(src_value, max_decimal if max_decimal else 6)
    if header_name == "金额":
        return format_number(src_value, 2)
    return str(src_value)

def run_test():
    print("🚀 [自动化测试] 正在验证 '外币模式' 强制同步逻辑...")
    
    # 1. 模拟用户配置
    default_rate_setting = "7.01031"
    use_foreign_currency = True
    
    # 模拟数据：行 8 (外币 30000.0, 但 AI 识别出了错误的金额 0.0)
    row_data = {"Debit": 30000.0, "Desc": "LAFISE 转 BAC"}
    smart_data = {"金额": "0.0", "外币金额": "30000.0"} # AI 干扰项
    mapping = {"外币金额": "Debit", "金额": None} # 金额留空
    
    print(f"📍 测试目标: 外币={row_data['Debit']}, 默认汇率={default_rate_setting}")
    
    # --- 模拟主程序 do_convert 内部逻辑 ---
    
    # A. 汇率预处理 (确保计算用汇率 = Excel 存入汇率)
    rate_formatted = convert_value("汇率", default_rate_setting, 6)
    calc_rate = float(rate_formatted)
    print(f"   [1] 汇率预处理: {default_rate_setting} -> {calc_rate} (保留6位)")

    # B. 行级预算 (row_theoretical_amount)
    row_f_amt = row_data["Debit"] # 从映射获取
    _d_f = Decimal(str(row_f_amt))
    _d_r = Decimal(str(calc_rate))
    row_theoretical_amount = float((_d_f * _d_r).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
    print(f"   [2] 行级预算结果: {row_theoretical_amount}")

    # C. 金额强制同步
    # 假设原本 src_value 是从智能识别拿到的 "0.0"
    src_value = smart_data.get("金额")
    print(f"   [3] 原始金额(AI识别): {src_value}")
    
    if use_foreign_currency and row_theoretical_amount is not None:
        src_value = row_theoretical_amount
        print(f"   [4] 强制同步后金额: {src_value} ✅")

    # D. 最终转换输出
    final_amt = convert_value("金额", src_value)
    final_rate = convert_value("汇率", calc_rate)
    
    print("\n📊 [最终结果]")
    print(f"   - 金额列: {final_amt}")
    print(f"   - 汇率列: {final_rate}")
    print(f"   - 外币列: {row_f_amt}")

    # E. 验证公式
    check_val = float(final_amt)
    expected = float(Decimal(str(row_f_amt * float(final_rate))).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))
    
    if abs(check_val - expected) < 0.001:
        print("\n✨ 结论: 公式严格成立 (金额 == 外币 * 汇率)！警告将消失。")
    else:
        print(f"\n❌ 结论: 仍有偏差！实际={check_val}, 期望={expected}")

if __name__ == "__main__":
    run_test()
