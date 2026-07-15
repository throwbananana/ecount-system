# -*- coding: utf-8 -*-
"""
测试实际转换流程 - 模拟GUI的转换逻辑
"""

import pandas as pd
from summary_intelligence import SummaryIntelligence

def test_actual_conversion():
    print("=" * 60)
    print("测试实际转换流程")
    print("=" * 60)

    # 初始化智能识别器
    recognizer = SummaryIntelligence()

    # 模拟源数据（来自 其他应收-巴拿马01.xlsx）
    source_data = [
        {"发生日期": "2025/6/6", "摘要": "朱德漪报销巴拿马床垫费用2852.3元", "收入明细金额": None, "支出明细金额": 4532.30},
        {"发生日期": "2025/6/6", "摘要": "张旖报销巴拿马安装电脑版vpn", "收入明细金额": None, "支出明细金额": 200.00},
        {"发生日期": "2025/6/6", "摘要": "付孔灵辉巴西巴拿马2024年下半年分红20万元", "收入明细金额": None, "支出明细金额": 100000.00},
        {"发生日期": "2025/6/6", "摘要": "付黄馨娴5月工资", "收入明细金额": None, "支出明细金额": 2250.00},
        {"发生日期": "2025/6/18", "摘要": "巴拿马货款汇入", "收入明细金额": 38721.00, "支出明细金额": None},
    ]

    print("\n源数据列名: 发生日期, 摘要, 收入明细金额, 支出明细金额")
    print("-" * 60)

    # 模拟映射（关键：摘要列需要正确映射）
    mapping = {
        "日期": "发生日期",
        "摘要": "摘要",  # 关键映射！
        "摘要名": "摘要",  # 也支持摘要名
        "金额": "支出明细金额",  # 或者根据实际情况选择
    }

    print(f"\n映射关系:")
    for k, v in mapping.items():
        print(f"  {k} <- {v}")

    print("\n" + "=" * 60)
    print("转换结果:")
    print("=" * 60)

    for i, row in enumerate(source_data, 1):
        print(f"\n[行 {i}]")
        print(f"  原始摘要: {row['摘要']}")

        # 获取摘要列（模拟 mapping.get("摘要") or mapping.get("摘要名")）
        summary_col = mapping.get("摘要") or mapping.get("摘要名")
        if summary_col:
            summary_value = row.get(summary_col)
            if summary_value:
                # 执行智能识别
                smart_data = recognizer.recognize(str(summary_value), row)

                print(f"  智能识别结果:")
                for key, value in smart_data.items():
                    if key != "摘要":  # 不重复打印摘要
                        print(f"    - {key}: {value}")

                # 检查是否识别到科目编码
                if "科目编码" in smart_data:
                    print(f"  [OK] 科目编码已识别: {smart_data['科目编码']}")
                else:
                    print(f"  [WARN] 未识别到科目编码")

                # 检查是否识别到部门
                if "部门" in smart_data:
                    print(f"  [OK] 部门已识别: {smart_data['部门']}")
            else:
                print(f"  [ERR] 摘要值为空")
        else:
            print(f"  [ERR] 未找到摘要列映射！智能识别无法执行！")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    test_actual_conversion()
