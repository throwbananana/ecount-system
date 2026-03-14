# -*- coding: utf-8 -*-
"""
测试预处理预览功能
"""

import pandas as pd
from pathlib import Path

def create_test_excel():
    """创建测试用Excel文件"""

    # 创建测试数据
    test_data = {
        "日期": ["2025-01-15", "2025-01-16", "2025-01-17", "2025-01-18", "2025-01-19"],
        "摘要": [
            "销售鞋类商品给ABEL CAPOTE，金额500美元",
            "采购原材料入库，金额1000元",
            "收到ADAIN BENITEZ RODRIGUEZ的货款2000美元",
            "支付运费300元",
            "发放工资给巴拿马部门5000元"
        ],
        "备注": ["", "", "", "", ""]
    }

    df = pd.DataFrame(test_data)

    # 保存为Excel文件
    output_file = Path(__file__).parent / "测试数据_预览功能.xlsx"
    df.to_excel(output_file, index=False, engine='openpyxl')

    print("=" * 60)
    print("测试Excel文件创建成功！")
    print("=" * 60)
    print(f"\n文件位置: {output_file}")
    print(f"\n包含 {len(df)} 行测试数据：")
    print("-" * 60)

    for idx, row in df.iterrows():
        print(f"\n第 {idx + 1} 行:")
        print(f"  日期: {row['日期']}")
        print(f"  摘要: {row['摘要']}")

    print("\n" + "=" * 60)
    print("测试步骤：")
    print("=" * 60)
    print("\n1. 运行主程序：python 亿看智能识别系统.py")
    print("\n2. 在'Excel凭证转换'标签页中：")
    print("   - 点击'选择原始Excel' → 选择'测试数据_预览功能.xlsx'")
    print("   - 确认'启用摘要智能识别'已勾选")
    print("   - 点击'自动识别匹配'")
    print("   - 将'摘要'列映射到摘要字段")
    print("\n3. 点击'开始转换并导出'")
    print("\n4. 预处理预览窗口应该会弹出，显示：")
    print("   - 统计信息：共处理 5 行数据，识别到 5 行包含可识别信息")
    print("   - 识别详情页：显示每行的摘要和识别结果")
    print("   - 转换结果预览页：显示最终的转换数据")
    print("\n5. 在预览窗口中：")
    print("   - 查看识别详情，确认识别准确")
    print("   - 切换到转换结果预览，查看表格数据")
    print("   - 点击'确认并导出'或'取消'")
    print("\n" + "=" * 60)
    print("\n预期识别结果示例：")
    print("=" * 60)

    print("\n第 1 行摘要: 销售鞋类商品给ABEL CAPOTE，金额500美元")
    print("  应识别到：")
    print("    - 摘要编码: 01")
    print("    - 类型: 1")
    print("    - 科目编码: 1122")
    print("    - 往来单位名: ABEL CAPOTE")
    print("    - 金额: 500.0")

    print("\n第 2 行摘要: 采购原材料入库，金额1000元")
    print("  应识别到：")
    print("    - 摘要编码: 03")
    print("    - 类型: 2")
    print("    - 科目编码: 2202")
    print("    - 金额: 1000.0")

    print("\n第 3 行摘要: 收到ADAIN BENITEZ RODRIGUEZ的货款2000美元")
    print("  应识别到：")
    print("    - 摘要编码: 05")
    print("    - 类型: 1")
    print("    - 科目编码: 1002")
    print("    - 往来单位名: ADAIN BENITEZ RODRIGUEZ")
    print("    - 金额: 2000.0")

    print("\n" + "=" * 60)

if __name__ == "__main__":
    create_test_excel()
