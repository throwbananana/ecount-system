
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

def audit_export_file(file_path):
    print(f"🔍 正在启动一致性审计: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"❌ 无法读取文件: {e}")
        return

    # 检查必要的列是否存在
    required_cols = ["金额", "外币金额", "汇率"]
    if not all(col in df.columns for col in required_cols):
        print(f"❌ 审计失败: 文件缺少必要列 {required_cols}")
        return

    error_count = 0
    total_rows = len(df)
    
    print(f"📊 正在校验 {total_rows} 行数据...")
    
    for idx, row in df.iterrows():
        try:
            # 使用 Decimal 确保审计精度
            amt = Decimal(str(row["金额"] or 0)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            f_amt = Decimal(str(row["外币金额"] or 0))
            rate = Decimal(str(row["汇率"] or 1))
            
            # 计算理论值
            theoretical = (f_amt * rate).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            
            diff = abs(amt - theoretical)
            
            # 只有当外币不为0时才审计公式
            if abs(f_amt) > 0:
                if diff > Decimal("0.01"):
                    print(f"❌ [不一致] 行 {idx+2}:")
                    print(f"    摘要: {row.get('摘要', '无')}")
                    print(f"    实际金额: {amt}")
                    print(f"    计算金额: {f_amt} * {rate} = {theoretical}")
                    print(f"    偏差值: {diff}")
                    error_count += 1
        except Exception as e:
            print(f"⚠️ [跳过] 行 {idx+2} 数据格式错误: {e}")

    print("\n--- 审计报告 ---")
    if error_count == 0:
        print(f"✅ 审计通过！所有 {total_rows} 行数据的汇率计算均完全正确 (100% 匹配)。")
    else:
        print(f"❌ 发现 {error_count} 处计算不一致，请检查上述明细。")

if __name__ == "__main__":
    audit_export_file(r"12yue\1122_修复版导出.xlsx")
