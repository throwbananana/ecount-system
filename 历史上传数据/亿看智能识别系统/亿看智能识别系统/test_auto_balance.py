# -*- coding: utf-8 -*-
"""
测试自动生成对方分录功能
"""

def test_auto_balance_logic():
    """模拟自动生成对方分录的逻辑"""
    print("=" * 60)
    print("测试自动生成对方分录功能")
    print("=" * 60)

    # 模拟默认值配置
    default_values = {
        "科目编码": "1002",      # 默认对方科目
        "默认账户往来": "",       # 默认对方往来
    }

    # 模拟转换后的行数据
    test_cases = [
        {
            "描述": "费用报销（借方）",
            "类型": "3",
            "金额": "4532.30",
            "科目编码": "510102",
            "往来单位编码": "",
            "部门": "10001",
        },
        {
            "描述": "收到货款（借方-银行）",
            "类型": "3",
            "金额": "38721.00",
            "科目编码": "1002",
            "往来单位编码": "H0405",
            "部门": "10001",
        },
        {
            "描述": "付款（贷方）",
            "类型": "4",
            "金额": "2000.00",
            "科目编码": "1002",
            "往来单位编码": "",
            "部门": "10001",
        },
        {
            "描述": "类型为空（不应生成对方分录）",
            "类型": "",
            "金额": "1000.00",
            "科目编码": "6602",
            "往来单位编码": "",
            "部门": "10001",
        },
    ]

    for i, case in enumerate(test_cases, 1):
        print(f"\n[测试 {i}] {case['描述']}")
        print(f"  原分录: 类型={case['类型']}, 科目={case['科目编码']}, 金额={case['金额']}, 往来={case['往来单位编码']}")

        current_type = case["类型"]
        current_subject = case["科目编码"]
        current_partner = case["往来单位编码"]

        # 只有类型为3或4时才生成对方分录
        if current_type not in ["3", "4"]:
            print(f"  [跳过] 类型不是3或4，不生成对方分录")
            continue

        # 反转类型
        new_type = "4" if current_type == "3" else "3"

        # 确定对方科目
        target_subject = default_values.get("科目编码", "")
        target_partner = default_values.get("默认账户往来", "")

        final_subject = target_subject
        final_partner = ""

        # 智能推断逻辑
        is_bank = current_subject.startswith("100")
        if is_bank and current_partner:
            # 如果原分录是银行类科目且有往来单位，对方科目设为 1122
            final_subject = "1122"
            final_partner = current_partner
            print(f"  [智能推断] 原分录是银行科目且有往来 -> 对方科目=1122")
        elif current_subject == target_subject:
            # 避免借贷同科目
            final_subject = ""
            if current_type == "3" and current_partner:
                final_subject = "1122"
                final_partner = current_partner
            print(f"  [智能推断] 原科目与默认科目相同 -> 尝试用1122")

        # 如果仍未确定对方科目，回退到默认值
        if not final_subject:
            final_subject = target_subject or "1002"

        print(f"  对方分录: 类型={new_type}, 科目={final_subject}, 往来={final_partner}")
        print(f"  [OK] 对方分录已生成")

    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

    print("\n" + "=" * 60)
    print("功能生效条件检查")
    print("=" * 60)
    print("""
要让"自动生成对方分录"功能生效，需要满足以下条件：

1. [必须] 勾选"自动生成对方分录（借贷平衡）"复选框
   - 位置：Excel凭证转换 标签页 -> 智能识别选项下方

2. [必须] 类型列必须有值且为 "3"(借) 或 "4"(贷)
   - 如果类型为空或其他值，不会生成对方分录
   - 检查方法：查看导出结果中的"类型"列

3. [建议] 设置默认值中的"科目编码"
   - 菜单：设置默认值... -> 科目编码
   - 这将作为对方分录的默认科目

4. [可选] 在源数据或摘要识别中提供往来单位编码
   - 系统已取消“默认往来单位”配置，不再按科目自动回填
""")

if __name__ == "__main__":
    test_auto_balance_logic()
