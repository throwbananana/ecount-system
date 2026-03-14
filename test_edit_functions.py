# -*- coding: utf-8 -*-
"""
测试基础数据编辑功能
"""

from base_data_manager import BaseDataManager

def test_edit_functions():
    """测试增删改功能"""
    print("=" * 60)
    print("测试基础数据编辑功能")
    print("=" * 60)

    mgr = BaseDataManager()

    # 1. 测试添加记录
    print("\n1. 测试添加部门记录...")
    new_dept = {
        "code": "TEST001",
        "name": "测试部门",
        "is_active": "是"
    }
    result = mgr.add_record("department", new_dept)
    print(f"   结果: {result['message']}")

    if result["success"]:
        new_id = result["id"]
        print(f"   新记录ID: {new_id}")

        # 2. 测试查询记录
        print("\n2. 测试查询新增的记录...")
        record = mgr.get_record_by_id("department", new_id)
        if record:
            print(f"   查询到: {record}")

        # 3. 测试更新记录
        print("\n3. 测试更新记录...")
        update_data = {
            "code": "TEST001",
            "name": "测试部门（已修改）",
            "is_active": "否"
        }
        result = mgr.update_record("department", new_id, update_data)
        print(f"   结果: {result['message']}")

        # 4. 验证更新
        print("\n4. 验证更新后的记录...")
        record = mgr.get_record_by_id("department", new_id)
        if record:
            print(f"   更新后: {record}")

        # 5. 测试删除记录
        print("\n5. 测试删除记录...")
        result = mgr.delete_record("department", new_id)
        print(f"   结果: {result['message']}")

        # 6. 验证删除
        print("\n6. 验证删除后...")
        record = mgr.get_record_by_id("department", new_id)
        if record is None:
            print("   记录已成功删除")
        else:
            print(f"   删除失败，记录仍存在: {record}")

    # 7. 测试重复编码
    print("\n7. 测试添加重复编码...")
    dup_dept = {
        "code": "10001",  # 这个编码已存在
        "name": "重复编码测试",
        "is_active": "是"
    }
    result = mgr.add_record("department", dup_dept)
    print(f"   结果: {result['message']}")
    if not result["success"]:
        print("   正确拒绝了重复编码 [OK]")

    # 8. 测试获取表列名
    print("\n8. 测试获取表列名...")
    columns = mgr.get_table_columns("department")
    print(f"   部门表列名: {columns}")

    mgr.close()

    print("\n" + "=" * 60)
    print("测试完成！")
    print("=" * 60)

if __name__ == "__main__":
    test_edit_functions()
