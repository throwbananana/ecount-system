# -*- coding: utf-8 -*-
"""
测试摘要智能识别功能
"""

from summary_intelligence import SummaryIntelligence

def test_smart_recognition():
    """测试智能识别各种摘要"""

    print("=" * 60)
    print("摘要智能识别测试")
    print("=" * 60)

    recognizer = SummaryIntelligence()

    # 测试用例
    test_cases = [
        "销售鞋类商品给ABEL CAPOTE，金额500美元",
        "采购原材料入库，金额1000元",
        "收到ADAIN BENITEZ RODRIGUEZ的货款2000美元",
        "支付运费300元",
        "发放工资给巴拿马部门",
        "办公费支出100元，购买文具",
        "库存商品出库，销售给ACTION SPORT CA",
        "2025-01-15采购退回，金额800元",
        "收到客户订金5000元",
        "支付货款给供应商3000元",
    ]

    print("\n测试用例:")
    for idx, summary in enumerate(test_cases, 1):
        print(f"\n{idx}. 摘要: {summary}")
        print("-" * 60)

        result = recognizer.recognize(summary)

        if result:
            print("   识别结果:")
            for key, value in result.items():
                if value:
                    print(f"   - {key}: {value}")
        else:
            print("   未识别到任何信息")

    recognizer.close()

    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)

if __name__ == "__main__":
    test_smart_recognition()
