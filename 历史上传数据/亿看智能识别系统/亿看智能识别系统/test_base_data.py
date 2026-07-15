# -*- coding: utf-8 -*-
"""
基础数据管理器测试脚本
"""

from base_data_manager import BaseDataManager
import os


def test_cache_match_items():
    """验证智能识别缓存支持多匹配项"""
    print("\n" + "=" * 60)
    print("智能识别缓存匹配项测试")
    print("=" * 60)

    mgr = BaseDataManager(db_path=":memory:")
    cols = mgr.get_table_columns("smart_recognition_cache")
    print(f"缓存表列: {cols}")

    # 保存包含多匹配项的记录
    mgr.save_cached_recognition("销售货款", "1122", ["收款", "收货款"])

    exact_hit = mgr.get_cached_recognition("销售货款")
    alias_hit = mgr.get_cached_recognition("收款")
    fuzzy_hit = mgr.get_cached_recognition_fuzzy("收货", min_ratio=0.5)

    print(f"直接命中: {exact_hit}")
    print(f"匹配项命中: {alias_hit}")
    print(f"模糊命中: {fuzzy_hit}")

    assert exact_hit == "1122"
    assert alias_hit == "1122"
    assert fuzzy_hit == "1122"

    mgr.close()


def test_cache_fuzzy_match_prefers_more_similar_summary():
    """避免短摘要因包含关系误命中到更早的缓存记录"""
    mgr = BaseDataManager(db_path=":memory:")
    try:
        mgr.save_cached_recognition("王力报销油费", "660232")
        mgr.save_cached_recognition("王力报销油费浙J7H1E1", "660230")

        hit = mgr.get_cached_recognition_fuzzy("王力报销油费浙J7H1E1路桥停车")
        assert hit == "660230", f"期望命中更相似的车辆费用缓存，实际为 {hit}"
    finally:
        mgr.close()


def test_save_cached_recognition_refreshes_alias_map():
    """重复保存同一摘要后，旧别名不应继续残留在内存索引中"""
    mgr = BaseDataManager(db_path=":memory:")
    try:
        mgr.save_cached_recognition("销售货款", "1122", ["收货款"])
        assert mgr.get_cached_recognition("收货款") == "1122"

        mgr.save_cached_recognition("销售货款", "2203", [])

        assert mgr.get_cached_recognition("销售货款") == "2203"
        assert mgr.get_cached_recognition("收货款") is None
    finally:
        mgr.close()


def test_base_data_manager():
    """测试基础数据管理器"""

    print("=" * 60)
    print("基础数据管理器测试")
    print("=" * 60)

    # 创建管理器
    print("\n1. 创建基础数据管理器...")
    mgr = BaseDataManager(db_path=":memory:")
    print(f"   数据库文件: {mgr.db_path}")
    print(f"   文件存在: {os.path.exists(mgr.db_path)}")

    # 获取初始统计
    print("\n2. 获取初始统计...")
    stats = mgr.get_statistics()
    for table, count in stats.items():
        print(f"   {table}: {count} 条记录")

    total = sum(stats.values())
    print(f"   总计: {total} 条记录")

    # 导入基础数据
    if total == 0:
        print("\n3. 导入基础数据...")
        result = mgr.import_all_data()
        print(f"   导入结果: {result['message']}")
        print("\n   详细信息:")
        for file, info in result["details"].items():
            status = "[OK]" if info["success"] else "[FAIL]"
            print(f"   {status} {file}: {info['message']}")

        # 再次获取统计
        print("\n4. 导入后统计...")
        stats = mgr.get_statistics()
        for table, count in stats.items():
            print(f"   {table}: {count} 条记录")

        total = sum(stats.values())
        print(f"   总计: {total} 条记录")
    else:
        print("\n3. 数据库已有数据，跳过导入")

    # 测试查询功能
    print("\n5. 测试查询功能...")

    # 查询科目编码（前5条）
    print("\n   查询科目编码（前5条）:")
    subjects = mgr.query("account_subject")[:5]
    for subj in subjects:
        print(f"   - {subj.get('code_name', 'N/A')}")

    # 查询往来单位（前5条）
    print("\n   查询往来单位（前5条）:")
    partners = mgr.query("business_partner")[:5]
    for partner in partners:
        print(f"   - [{partner.get('code', 'N/A')}] {partner.get('name', 'N/A')}")

    # 测试搜索功能
    print("\n6. 测试搜索功能...")
    print("\n   搜索往来单位名称包含 'ABEL':")
    results = mgr.search_by_name("business_partner", "ABEL")
    for r in results:
        print(f"   - [{r.get('code', 'N/A')}] {r.get('name', 'N/A')}")

    # 关闭连接
    print("\n7. 关闭数据库连接...")
    mgr.close()
    print("   完成")

    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)


def test_table_name_validation():
    """验证表名白名单防注入"""
    mgr = BaseDataManager(db_path=":memory:")
    try:
        mgr.query("account_subject;DROP TABLE smart_recognition_cache;--")
    except ValueError:
        print("表名校验正常生效")
    else:
        raise AssertionError("危险表名未被拒绝")
    finally:
        mgr.close()

if __name__ == "__main__":
    test_base_data_manager()
    test_cache_match_items()
    test_cache_fuzzy_match_prefers_more_similar_summary()
    test_save_cached_recognition_refreshes_alias_map()
    test_table_name_validation()
