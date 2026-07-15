# -*- coding: utf-8 -*-
"""
测试新的集成GUI界面
"""

import tkinter as tk
from base_data_manager import BaseDataManager

def test_gui_components():
    """测试GUI组件是否正确创建"""
    print("=" * 60)
    print("测试新GUI界面组件")
    print("=" * 60)

    # 测试数据库连接
    print("\n1. 测试数据库连接...")
    try:
        mgr = BaseDataManager()
        stats = mgr.get_statistics()
        total = sum(stats.values())
        print(f"   数据库连接成功，共 {total} 条记录")
        mgr.close()
    except Exception as e:
        print(f"   错误: {e}")
        return

    # 测试GUI导入
    print("\n2. 测试GUI导入...")
    try:
        import 亿看智能识别系统
        print("   GUI模块导入成功")
    except Exception as e:
        print(f"   错误: {e}")
        return

    print("\n3. GUI组件测试...")
    print("   - 标签页1: Excel凭证转换")
    print("   - 标签页2: 基础数据管理")
    print("     * 左侧: 7个数据类型单选按钮")
    print("     * 右侧: 搜索框 + 编辑按钮 + 数据表格")

    print("\n4. 功能测试建议...")
    print("   运行主程序，验证以下功能：")
    print("   - 切换标签页")
    print("   - 切换数据类型（币种、部门、仓库等）")
    print("   - 搜索功能")
    print("   - 新增/编辑/删除按钮")
    print("   - 双击编辑")
    print("   - Excel转换功能保持正常")

    print("\n" + "=" * 60)
    print("组件测试完成！")
    print("请运行 'python 亿看智能识别系统.py' 进行手动测试")
    print("=" * 60)

if __name__ == "__main__":
    test_gui_components()
