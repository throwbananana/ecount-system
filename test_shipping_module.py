# -*- coding: utf-8 -*-
"""
报关清单模块回归测试
"""

import math
import tempfile
from pathlib import Path

import pandas as pd

from shipping_module import ShippingDB


def _build_shipping_excel(path: Path):
    df = pd.DataFrame([
        {
            "集装箱号": "TEST001",
            "厂家": "金奥",
            "名称": "产品A",
            "型号": "A-1",
            "数量": 10,
            "单价": 10,
            "总金额": 100,
            "总体积": 1.0,
        },
        {
            "集装箱号": "TEST001",
            "厂家": "金奥",
            "名称": "产品B",
            "型号": "B-1",
            "数量": 20,
            "单价": 10,
            "总金额": 200,
            "总体积": 3.0,
        },
        {
            "集装箱号": "TEST001",
            "海运费": 100.0,
            "包干费": 200.0,
        },
        {
            "集装箱号": "TEST001",
            "说明": "代理费",
            "值": 50.0,
        },
        {
            "集装箱号": "TEST001",
            "说明": "保费",
            "值": 5.0,
        },
        {
            "集装箱号": "TEST001",
            "说明": "汇率",
            "值": 7.0,
        },
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="报关清单9%", index=False)


def _sum_allocated_cost(db: ShippingDB, container_id: int) -> float:
    row = db.conn.execute(
        "SELECT COALESCE(SUM(allocated_cost), 0) AS total FROM products WHERE container_id=?",
        (container_id,),
    ).fetchone()
    return float(row["total"] or 0.0)


def test_repeat_import_replaces_existing_products():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        db = ShippingDB(str(tmp_path / "shipping.bd"))
        try:
            excel_path = tmp_path / "sample_ZKP2026999.xlsx"
            _build_shipping_excel(excel_path)

            db.import_excel(str(excel_path))
            product_count_1 = db.conn.execute("SELECT COUNT(*) AS c FROM products").fetchone()["c"]
            container_count_1 = db.conn.execute("SELECT COUNT(*) AS c FROM containers").fetchone()["c"]

            db.import_excel(str(excel_path))
            product_count_2 = db.conn.execute("SELECT COUNT(*) AS c FROM products").fetchone()["c"]
            container_count_2 = db.conn.execute("SELECT COUNT(*) AS c FROM containers").fetchone()["c"]

            assert product_count_1 == 2, f"首次导入产品数异常: {product_count_1}"
            assert product_count_2 == 2, f"重复导入后产品数不应累积: {product_count_2}"
            assert container_count_1 == 1 and container_count_2 == 1, "同一柜重复导入不应新增货柜记录"
        finally:
            db.conn.close()


def test_update_container_fees_keeps_allocations_in_sync():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        db = ShippingDB(str(tmp_path / "shipping.bd"))
        try:
            excel_path = tmp_path / "sample_ZKP2026999.xlsx"
            _build_shipping_excel(excel_path)

            db.import_excel(str(excel_path))
            container = db.query_containers()[0]
            container_id = container["id"]

            db.allocate_misc_fees(container_id)
            initial_total = float(db.query_containers()[0]["misc_total_rmb"])
            initial_alloc_sum = _sum_allocated_cost(db, container_id)
            assert math.isclose(initial_alloc_sum, initial_total, rel_tol=0, abs_tol=1e-6)

            fees = db._get_container_fee_parts(container_id)
            fees.update({
                "all_in_rmb": 300.0,
                "agency_fee_rmb": 80.0,
                "sea_freight_usd": 120.0,
                "insurance_usd": 8.0,
                "exchange_rate": 7.2,
            })
            new_total = db.update_container_fees(container_id, fees)

            alloc_sum = _sum_allocated_cost(db, container_id)
            assert math.isclose(alloc_sum, new_total, rel_tol=0, abs_tol=1e-6), (
                f"更新货柜费用后，产品分摊未同步: 分摊合计={alloc_sum}, 货柜总额={new_total}"
            )
        finally:
            db.conn.close()


def test_update_container_field_recalculates_allocations():
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        db = ShippingDB(str(tmp_path / "shipping.bd"))
        try:
            excel_path = tmp_path / "sample_ZKP2026999.xlsx"
            _build_shipping_excel(excel_path)

            db.import_excel(str(excel_path))
            container = db.query_containers()[0]
            container_id = container["id"]

            db.allocate_misc_fees(container_id)
            new_total = db.update_container_field(container_id, "agency_fee_rmb", 120.0)
            alloc_sum = _sum_allocated_cost(db, container_id)

            assert math.isclose(alloc_sum, float(new_total), rel_tol=0, abs_tol=1e-6), (
                f"单字段修改后，产品分摊未同步: 分摊合计={alloc_sum}, 货柜总额={new_total}"
            )
        finally:
            db.conn.close()


if __name__ == "__main__":
    test_repeat_import_replaces_existing_products()
    test_update_container_fees_keeps_allocations_in_sync()
    test_update_container_field_recalculates_allocations()
    print("test_shipping_module.py: all tests passed")
