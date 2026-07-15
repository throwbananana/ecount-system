#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校验经营分析报告关键值与仪表盘公式范围。"""

import argparse
import os
import sys

from report_generator import ReportGenerator


def parse_args():
    parser = argparse.ArgumentParser(description="校验经营分析报告关键值")
    parser.add_argument("--report", required=True, help="报表文件路径 (.xlsx)")
    parser.add_argument("--year", required=True, help="目标年份，例如 2025")
    parser.add_argument("--month", required=True, help="目标月份，例如 12")
    parser.add_argument(
        "--base-dir",
        default=os.path.join(os.path.dirname(__file__), "基础资料"),
        help="基础资料目录，默认: ./基础资料",
    )
    parser.add_argument(
        "--year-scope",
        default="current",
        choices=["current", "all"],
        help="年份范围，默认 current",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if not os.path.exists(args.report):
        print(f"[ERROR] 报表不存在: {args.report}")
        return 2
    if not os.path.exists(args.base_dir):
        print(f"[ERROR] 基础资料目录不存在: {args.base_dir}")
        return 2

    gen = ReportGenerator(args.base_dir)
    gen.load_all_data()
    issues = gen.validate_report_file(args.report, args.year, args.month, args.year_scope)

    print("\n=== 校验结果 ===")
    if not issues:
        print("PASS: 未发现关键问题")
        return 0

    error_count = 0
    warn_count = 0
    for item in issues:
        sev = (item.get("severity") or "WARN").upper()
        sheet = item.get("sheet") or "未知"
        msg = item.get("message") or ""
        print(f"[{sev}] {sheet}: {msg}")
        if sev == "ERROR":
            error_count += 1
        else:
            warn_count += 1

    print(f"汇总: ERROR={error_count}, WARN={warn_count}, TOTAL={len(issues)}")
    return 1 if error_count else 0


if __name__ == "__main__":
    sys.exit(main())
