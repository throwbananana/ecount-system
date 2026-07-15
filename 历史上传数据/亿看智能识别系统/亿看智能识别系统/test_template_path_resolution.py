# -*- coding: utf-8 -*-
"""
验证模板路径解析在非项目目录启动时仍可正常工作。
"""

import importlib.util
import os
import shutil
import tempfile
from pathlib import Path


MODULE_PATH = Path(__file__).with_name("亿看智能识别系统.py")
EXPECTED_TEMPLATE = (MODULE_PATH.parent / "Template.xlsx").resolve()
EXPECTED_BASE_DATA_DIR = (MODULE_PATH.parent / "基础数据" / "基础数据").resolve()
EXPECTED_REPORT_BASE_DIR = (MODULE_PATH.parent / "基础资料").resolve()
EXPECTED_REPORT_TEMPLATE = (
    MODULE_PATH.parent / "11月汇总结果_整理优化美化_含仪表盘_目标预算达成异常 (1).xlsx"
).resolve()


def load_app_module():
    spec = importlib.util.spec_from_file_location("yikan_app", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_template_resolution_from_outside_project_dir():
    module = load_app_module()
    original_cwd = os.getcwd()
    tmp_dir = tempfile.mkdtemp()

    try:
        os.chdir(tmp_dir)
        resolved = Path(module.resolve_template_path("Template.xlsx")).resolve()
        assert resolved == EXPECTED_TEMPLATE, (
            f"模板路径解析错误: expected={EXPECTED_TEMPLATE}, actual={resolved}"
        )

        resolved_base_data_dir = Path(
            module.resolve_resource_dir("基础数据/基础数据", module.BASE_DATA_DIR_NAME)
        ).resolve()
        assert resolved_base_data_dir == EXPECTED_BASE_DATA_DIR, (
            f"基础数据目录解析错误: expected={EXPECTED_BASE_DATA_DIR}, actual={resolved_base_data_dir}"
        )

        resolved_report_base_dir = Path(
            module.resolve_resource_dir(module.REPORT_BASE_DIR_NAME, module.REPORT_BASE_DIR_NAME)
        ).resolve()
        assert resolved_report_base_dir == EXPECTED_REPORT_BASE_DIR, (
            f"经营报告基础资料目录解析错误: expected={EXPECTED_REPORT_BASE_DIR}, actual={resolved_report_base_dir}"
        )

        resolved_report_template = Path(
            module.resolve_resource_file(
                module.DEFAULT_REPORT_TEMPLATE_FILE,
                module.DEFAULT_REPORT_TEMPLATE_FILE,
            )
        ).resolve()
        assert resolved_report_template == EXPECTED_REPORT_TEMPLATE, (
            f"经营报告模板解析错误: expected={EXPECTED_REPORT_TEMPLATE}, actual={resolved_report_template}"
        )

        headers, workbook, _ = module.load_template_headers("Template.xlsx")
        assert headers, "模板表头不能为空"
        workbook.close()
    finally:
        os.chdir(original_cwd)
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    test_template_resolution_from_outside_project_dir()
    print("模板路径解析测试通过")
