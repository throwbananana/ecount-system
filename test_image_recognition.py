# -*- coding: utf-8 -*-
"""
测试图片智能识别功能
"""

import os
import sys

def test_dependencies():
    """测试依赖安装"""
    print("=" * 60)
    print("测试依赖安装")
    print("=" * 60)

    try:
        from image_intelligence import check_and_install_dependencies
        status = check_and_install_dependencies(auto_install=True)
        print("\n依赖状态:")
        for name, available in status.items():
            print(f"  {name}: {'已安装' if available else '未安装'}")
        return True
    except Exception as e:
        print(f"错误: {e}")
        return False


def test_image_intelligence():
    """测试图片识别模块"""
    print("\n" + "=" * 60)
    print("测试图片识别模块")
    print("=" * 60)

    try:
        from image_intelligence import ImageIntelligence

        # 使用环境变量中的智谱AI Key，避免硬编码泄露
        api_key = os.environ.get("ZHIPU_API_KEY", "")
        if not api_key:
            raise RuntimeError("未设置 ZHIPU_API_KEY 环境变量，无法进行远程图片识别测试")

        recognizer = ImageIntelligence(
            ai_provider="zhipu",
            api_key=api_key,
            auto_install=True
        )
        print("  图片识别器初始化成功")
        return recognizer
    except Exception as e:
        print(f"  错误: {e}")
        return None


def test_image_recognition(recognizer, image_path):
    """测试单张图片识别"""
    print(f"\n正在识别: {os.path.basename(image_path)}")

    if not os.path.exists(image_path):
        print(f"  文件不存在: {image_path}")
        return None

    result = recognizer.recognize_image(image_path, use_ai=True)

    print(f"  状态: {result.get('status')}")

    if result.get("status") == "success":
        headers = result.get("headers", [])
        rows = result.get("rows", [])
        print(f"  表头: {headers}")
        print(f"  数据行数: {len(rows)}")

        if rows:
            print(f"  第一行数据: {rows[0]}")

    elif result.get("status") == "partial":
        print(f"  消息: {result.get('message')}")
        raw_text = result.get("raw_text", "")
        print(f"  原始文本 (前200字): {raw_text[:200]}...")

    else:
        print(f"  错误: {result.get('message')}")

    return result


def test_batch_recognition(recognizer, folder_path):
    """测试批量识别"""
    print("\n" + "=" * 60)
    print("测试批量识别")
    print("=" * 60)

    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return []

    # 收集图片文件
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
    image_files = []

    for filename in sorted(os.listdir(folder_path)):
        ext = os.path.splitext(filename)[1].lower()
        if ext in image_extensions:
            image_files.append(os.path.join(folder_path, filename))

    print(f"找到 {len(image_files)} 张图片")

    if not image_files:
        return []

    # 限制测试数量
    test_files = image_files[:3]
    print(f"测试前 {len(test_files)} 张...")

    results = recognizer.batch_recognize(test_files, use_ai=True)

    # 统计结果
    success_count = sum(1 for r in results if r.get("status") == "success")
    print(f"\n识别完成: {success_count}/{len(test_files)} 成功")

    return results


def test_one_click_batch(recognizer, folder_path, output_path):
    """测试一键批量识别并合并导出"""
    print("\n" + "=" * 60)
    print("测试一键批量识别并合并导出")
    print("=" * 60)

    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return None

    # 收集所有图片文件
    image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
    image_files = []

    for filename in sorted(os.listdir(folder_path)):
        ext = os.path.splitext(filename)[1].lower()
        if ext in image_extensions:
            image_files.append(os.path.join(folder_path, filename))

    print(f"找到 {len(image_files)} 张图片")

    if not image_files:
        return None

    # 一键批量识别并合并导出
    print(f"开始一键批量识别...")
    result = recognizer.batch_recognize_and_merge(
        image_files,
        use_ai=True,
        output_path=output_path
    )

    print(f"\n识别统计:")
    print(f"  - 总图片数: {result['total_images']}")
    print(f"  - 成功识别: {result['success_count']}")
    print(f"  - 部分识别: {result['partial_count']}")
    print(f"  - 识别失败: {result['error_count']}")
    print(f"  - 合并数据行数: {result['total_rows']}")

    if result.get('exported'):
        print(f"  - 已导出到: {result['export_path']}")
    elif result['total_rows'] == 0:
        print(f"  - 警告: 没有数据可导出")

    return result


def test_export(recognizer, results, output_path):
    """测试导出功能"""
    print("\n" + "=" * 60)
    print("测试导出功能")
    print("=" * 60)

    # 合并结果
    headers, rows = recognizer.merge_results_to_table(results)

    print(f"合并后: {len(rows)} 行数据")

    if not rows:
        print("没有数据可导出")
        return False

    # 导出
    success = recognizer.export_to_excel(headers, rows, output_path)

    if success:
        print(f"导出成功: {output_path}")
    else:
        print("导出失败")

    return success


def test_gui():
    """测试GUI窗口"""
    print("\n" + "=" * 60)
    print("测试GUI窗口")
    print("=" * 60)

    try:
        import tkinter as tk
        from image_recognition_gui import ImageRecognitionWindow

        root = tk.Tk()
        root.withdraw()

        api_key = os.environ.get("ZHIPU_API_KEY", "")
        if not api_key:
            raise RuntimeError("未设置 ZHIPU_API_KEY 环境变量，无法启动带云识别的 GUI 测试")

        window = ImageRecognitionWindow(
            parent=root,
            api_key=api_key
        )

        print("GUI窗口创建成功")
        print("请在窗口中测试以下功能:")
        print("  1. 导入图片")
        print("  2. 识别当前/批量识别")
        print("  3. 导出Excel")
        print("")
        print("关闭窗口后程序将退出")

        window.window.protocol("WM_DELETE_WINDOW", lambda: (root.quit(), root.destroy()))
        root.mainloop()

        return True
    except Exception as e:
        print(f"GUI测试失败: {e}")
        return False


def main():
    print("=" * 60)
    print("亿看智能识别系统 - 图片识别功能测试")
    print("=" * 60)

    # 1. 测试依赖
    if not test_dependencies():
        print("\n依赖测试失败，请检查环境")
        return

    # 2. 测试模块初始化
    recognizer = test_image_intelligence()
    if not recognizer:
        print("\n模块初始化失败")
        return

    # 3. 测试单张图片识别
    test_folder = os.path.join(os.path.dirname(__file__), "图片智能识别")
    if os.path.exists(test_folder):
        # 找第一张图片测试
        for f in os.listdir(test_folder):
            if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                test_image = os.path.join(test_folder, f)
                test_image_recognition(recognizer, test_image)
                break

        # 4. 测试一键批量识别并合并导出
        output_file = os.path.join(os.path.dirname(__file__), "图片识别合并结果.xlsx")
        test_one_click_batch(recognizer, test_folder, output_file)

    else:
        print(f"\n测试文件夹不存在: {test_folder}")
        print("跳过图片识别测试")

    # 6. 询问是否测试GUI
    print("\n" + "=" * 60)
    response = input("是否打开GUI窗口进行测试? (y/n): ").strip().lower()
    if response == 'y':
        test_gui()

    print("\n" + "=" * 60)
    print("测试完成!")
    print("=" * 60)


if __name__ == "__main__":
    main()
