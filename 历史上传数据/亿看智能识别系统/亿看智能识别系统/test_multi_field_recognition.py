# -*- coding: utf-8 -*-
"""
测试多字段智能识别和默认值功能
"""

from summary_intelligence import SummaryIntelligence
import pandas as pd

def test_multi_field_recognition():
    """测试从日期、摘要、金额、汇率多字段识别"""

    print("=" * 60)
    print("多字段智能识别测试")
    print("=" * 60)

    # 设置默认值
    default_values = {
        "部门": "10001",  # 默认巴拿马部门
        "外币金额": "0",
        "汇率": "1"
    }

    recognizer = SummaryIntelligence(default_values=default_values)

    # 测试用例1：从多个字段识别
    print("\n" + "=" * 60)
    print("测试1：从原始数据的多个字段识别")
    print("=" * 60)

    original_data = {
        "日期": "2025-01-15",
        "摘要": "销售鞋类商品给ABEL CAPOTE",
        "金额": "500.50",
        "汇率": "6.5"
    }

    print("\n原始数据:")
    for key, value in original_data.items():
        print(f"  {key}: {value}")

    result = recognizer.recognize(
        summary=original_data.get("摘要", ""),
        original_data=original_data
    )

    print("\n识别结果:")
    for key, value in result.items():
        print(f"  {key}: {value}")

    # 测试用例2：只有摘要，应用默认值
    print("\n" + "=" * 60)
    print("测试2：只有摘要，其他字段使用默认值")
    print("=" * 60)

    summary_only = "采购原材料入库"

    print(f"\n摘要: {summary_only}")

    result2 = recognizer.recognize(summary_only)

    print("\n识别结果（包含默认值）:")
    for key, value in result2.items():
        print(f"  {key}: {value}")

    # 测试用例3：完整的数据行
    print("\n" + "=" * 60)
    print("测试3：完整数据行（模拟实际Excel导入）")
    print("=" * 60)

    test_rows = [
        {
            "日期": "2025-01-15",
            "摘要": "收到ADAIN BENITEZ RODRIGUEZ的货款",
            "金额": 2000.00,
            "汇率": 6.8
        },
        {
            "日期": "2025-01-16",
            "摘要": "支付运费",
            "amt": 300  # 使用不同的字段名
        },
        {
            "日期": "2025-01-17",
            "摘要": "发放工资给巴拿马部门",
            "金额": "5000元"
        }
    ]

    for idx, row in enumerate(test_rows, 1):
        print(f"\n第 {idx} 行:")
        print("  原始数据:")
        for k, v in row.items():
            print(f"    {k}: {v}")

        result = recognizer.recognize(
            summary=row.get("摘要", ""),
            original_data=row
        )

        print("  识别结果:")
        for k, v in result.items():
            if v:  # 只显示非空值
                print(f"    {k}: {v}")

    # 测试用例4：优先级测试
    print("\n" + "=" * 60)
    print("测试4：优先级测试（默认值 < 原始数据 < 摘要识别）")
    print("=" * 60)

    # 场景：日期既在原始数据中，也在摘要中
    test_data = {
        "日期": "2025-01-20",  # 原始数据中的日期
        "摘要": "2025-01-25采购退回，金额800元",  # 摘要中也包含日期
        "金额": 1000  # 原始数据中的金额
    }

    print("\n原始数据:")
    for k, v in test_data.items():
        print(f"  {k}: {v}")

    result = recognizer.recognize(
        summary=test_data.get("摘要", ""),
        original_data=test_data
    )

    print("\n识别结果:")
    print(f"  凭证日期: {result.get('凭证日期')} （应该是20250120，来自原始数据）")
    print(f"  金额: {result.get('金额')} （应该是1000，来自原始数据）")
    print(f"  摘要编码: {result.get('摘要编码')} （应该是04，来自摘要识别）")
    print(f"  部门: {result.get('部门')} （应该是10001，来自默认值）")

    recognizer.close()

    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)

    print("\n优先级规则:")
    print("  1. 手动映射（最高优先级，在实际转换中）")
    print("  2. 原始数据字段（日期、金额、汇率等）")
    print("  3. 摘要智能识别")
    print("  4. 默认值（最低优先级）")

if __name__ == "__main__":
    test_multi_field_recognition()
