# -*- coding: utf-8 -*-
"""
经营报告库存快照同步测试
"""

import pandas as pd

from base_data_manager import BaseDataManager
from report_generator import ReportGenerator


def test_report_generator_latest_inventory_snapshot_uses_latest_month():
    generator = ReportGenerator(".")
    generator.data["sales"]["2026-02"] = pd.DataFrame(
        {
            "品目编码": ["A001"],
            "品目名": ["产品A正式名"],
            "品目组合1名": ["电器类"],
            "日期": ["2026-02-15"],
        }
    )
    generator.data["cost"]["2026-01"] = pd.DataFrame(
        {
            "品目编码": ["A001"],
            "品目名": ["产品A"],
            "期初": [8],
            "期初.2": [40],
            "期末": [10],
            "期末.2": [50],
        }
    )
    generator.data["cost"]["2026-02"] = pd.DataFrame(
        {
            "品目编码": ["A001"],
            "品目名": ["产品A"],
            "期初": [10],
            "期初.2": [50],
            "期末": [20],
            "期末.2": [120],
        }
    )

    snapshot = generator.get_latest_product_inventory_snapshot()

    assert snapshot["month_key"] == "2026-02"
    assert len(snapshot["records"]) == 1
    assert snapshot["records"][0]["code"] == "A001"
    assert snapshot["records"][0]["latest_inventory_qty"] == 20
    assert snapshot["records"][0]["latest_inventory_cost"] == 6.0
    assert snapshot["records"][0]["name"] == "产品A正式名"
    assert snapshot["records"][0]["product_type"] == "电器类"
    assert snapshot["records"][0]["spec_info"] == "产品A"


def test_base_data_manager_sync_product_inventory_snapshot_keeps_newer_month():
    mgr = BaseDataManager(db_path=":memory:")
    try:
        add_result = mgr.add_record(
            "product",
            {
                "code": "A001",
                "name": "产品A",
                "product_type": "测试",
            },
        )
        assert add_result["success"], add_result

        first_sync = mgr.sync_product_inventory_snapshot(
            [
                {
                    "code": "A001",
                    "latest_inventory_qty": 20,
                    "latest_inventory_cost": 6.0,
                    "latest_inventory_date": "2026-02",
                }
            ],
            "2026-02",
        )
        assert first_sync["updated"] == 1, first_sync

        older_sync = mgr.sync_product_inventory_snapshot(
            [
                {
                    "code": "A001",
                    "latest_inventory_qty": 5,
                    "latest_inventory_cost": 2.0,
                    "latest_inventory_date": "2026-01",
                }
            ],
            "2026-01",
        )
        assert older_sync["updated"] == 0, older_sync
        assert older_sync["skipped_older"] == 1, older_sync

        newer_sync = mgr.sync_product_inventory_snapshot(
            [
                {
                    "code": "A001",
                    "latest_inventory_qty": 30,
                    "latest_inventory_cost": 7.5,
                    "latest_inventory_date": "2026-03",
                }
            ],
            "2026-03",
        )
        assert newer_sync["updated"] == 1, newer_sync

        row = mgr.query("product", "A001")[0]
        assert row["latest_inventory_qty"] == 30
        assert row["latest_inventory_cost"] == 7.5
        assert row["latest_inventory_date"] == "2026-03"
    finally:
        mgr.close()


def test_base_data_manager_sync_product_inventory_snapshot_adds_new_codes_and_only_fills_blank_metadata():
    mgr = BaseDataManager(db_path=":memory:")
    try:
        add_result = mgr.add_record(
            "product",
            {
                "code": "A001",
                "name": "人工维护名称",
                "product_type": None,
                "spec_info": None,
                "specification": None,
            },
        )
        assert add_result["success"], add_result

        sync_result = mgr.sync_product_inventory_snapshot(
            [
                {
                    "code": "A001",
                    "name": "报告名称",
                    "product_type": "电器类",
                    "spec_info": "规格A",
                    "specification": "规格A",
                    "latest_inventory_qty": 10,
                    "latest_inventory_cost": 3.2,
                    "latest_inventory_date": "2026-02",
                },
                {
                    "code": "B002",
                    "name": "新品B",
                    "product_type": "鞋类",
                    "spec_info": "规格B",
                    "specification": "规格B",
                    "latest_inventory_qty": 8,
                    "latest_inventory_cost": 5.5,
                    "latest_inventory_date": "2026-02",
                },
            ],
            "2026-02",
        )

        assert sync_result["updated"] == 1, sync_result
        assert sync_result["inserted"] == 1, sync_result
        assert sync_result["metadata_filled"] >= 2, sync_result

        existing = mgr.query("product", "A001")[0]
        assert existing["name"] == "人工维护名称"
        assert existing["product_type"] == "电器类"
        assert existing["spec_info"] == "规格A"
        assert existing["latest_inventory_qty"] == 10

        inserted = mgr.query("product", "B002")[0]
        assert inserted["name"] == "新品B"
        assert inserted["product_type"] == "鞋类"
        assert inserted["spec_info"] == "规格B"
        assert inserted["latest_inventory_cost"] == 5.5
    finally:
        mgr.close()


def test_report_generator_latest_business_partner_snapshot_uses_latest_coded_rows():
    generator = ReportGenerator(".")
    generator.data["sales"]["2026-02"] = pd.DataFrame(
        {
            "客户编码": ["C001", "C002"],
            "客户名称": ["客户甲", "客户乙"],
            "科目编码": ["1122", "2202"],
            "科目名": ["应收账款", "应付账款"],
            "日期": ["2026-02-01", "2026-02-20"],
        }
    )
    generator.ar_detail_df = pd.DataFrame(
        {
            "客户编码": ["C001", "C003"],
            "客户名称": ["客户甲应收", "客户丙"],
            "科目编码": ["1122", "1123"],
            "科目名": ["应收账款", "其他应收款"],
            "日期": ["2026-01-15", "2026-03-05"],
        }
    )

    snapshot = generator.get_latest_business_partner_snapshot()

    assert snapshot["month_key"] == "2026-03"
    records = {row["code"]: row for row in snapshot["records"]}
    assert records["C001"]["name"] == "客户甲"
    assert records["C002"]["name"] == "客户乙"
    assert records["C003"]["name"] == "客户丙"
    assert records["C003"]["local_code"] == "C003"
    assert records["C002"]["account_subject"] == "[2202] 应付账款"
    assert records["C003"]["account_subject"] == "[1123] 其他应收款"


def test_base_data_manager_sync_business_partner_snapshot_adds_new_and_matches_local_code():
    mgr = BaseDataManager(db_path=":memory:")
    try:
        add_result = mgr.add_record(
            "business_partner",
            {
                "code": "YK001",
                "name": "人工维护客户",
                "category": None,
                "local_code": "C001",
                "account_subject": "[1122] 应收账款",
            },
        )
        assert add_result["success"], add_result

        sync_result = mgr.sync_business_partner_snapshot(
            [
                {
                    "code": "C001",
                    "name": "报表客户甲",
                    "local_code": "C001",
                    "category": "零售",
                    "account_subject": "[2202] 应付账款",
                },
                {
                    "code": "C002",
                    "name": "新客户乙",
                    "local_code": "C002",
                    "category": "批发",
                    "account_subject": "[1122] 应收账款",
                },
            ],
            "2026-03",
        )

        assert sync_result["updated"] == 1, sync_result
        assert sync_result["inserted"] == 1, sync_result
        assert sync_result["account_subject_updated"] == 2, sync_result

        existing = mgr.query("business_partner", "YK001")[0]
        assert existing["name"] == "人工维护客户"
        assert existing["category"] == "零售"
        assert existing["local_code"] == "C001"
        assert existing["account_subject"] == "[2202] 应付账款"

        inserted = mgr.query("business_partner", "C002")[0]
        assert inserted["name"] == "新客户乙"
        assert inserted["category"] == "批发"
        assert inserted["local_code"] == "C002"
        assert inserted["account_subject"] == "[1122] 应收账款"
    finally:
        mgr.close()


if __name__ == "__main__":
    test_report_generator_latest_inventory_snapshot_uses_latest_month()
    test_base_data_manager_sync_product_inventory_snapshot_keeps_newer_month()
    test_base_data_manager_sync_product_inventory_snapshot_adds_new_codes_and_only_fills_blank_metadata()
    test_report_generator_latest_business_partner_snapshot_uses_latest_coded_rows()
    test_base_data_manager_sync_business_partner_snapshot_adds_new_and_matches_local_code()
    print("test_report_inventory_sync passed")
