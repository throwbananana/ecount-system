import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import sqlite3
import re
import math
import os
import json
from collections import defaultdict

import pandas as pd
import numpy as np
from export_format_manager import (
    apply_export_format,
    get_active_export_format_name,
    get_export_format_names,
    open_export_format_editor,
    set_active_export_format,
)
from treeview_tools import add_smart_restore_menu, attach_treeview_tools
from shipping_report_utils import add_charts_to_product_report, add_charts_to_container_report


# ======================== 数据库层 ========================

class ShippingDB:
    """
    使用 SQLite 做本地持久化，文件名默认 shipping.bd
    """
    def __init__(self, db_path="shipping.bd"):
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.init_schema()

    def init_schema(self):
        c = self.conn.cursor()
        # 货柜表：一票一柜一行
        c.execute("""
        CREATE TABLE IF NOT EXISTS containers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shipment_code TEXT,          -- 如 ZL202505
            container_no TEXT,           -- 集装箱号
            file_name TEXT,              -- 来源 Excel 文件名

            tax_refund REAL,             -- 退税额(人民币)
            sea_freight_usd REAL,        -- 海运费(USD)
            all_in_rmb REAL,             -- 包干费(RMB)
            insurance_usd REAL,          -- 保费(USD) - 修改为美元
            exchange_rate REAL,          -- 汇率(USD→RMB)
            agency_fee_rmb REAL,         -- 代理费(RMB)
            misc_rmb REAL,               -- 其他杂费(RMB)
            misc_total_rmb REAL,         -- 杂费汇总(RMB)

            UNIQUE (shipment_code, container_no)
        )
        """ )

        # 尝试添加 insurance_usd 列（兼容旧库）
        try:
            c.execute("ALTER TABLE containers ADD COLUMN insurance_usd REAL")
        except sqlite3.OperationalError:
            pass

        # 产品明细表
        c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            container_id INTEGER,        -- 关联 containers.id

            factory TEXT,                -- 厂家
            name TEXT,                   -- 品名
            model TEXT,                  -- 型号
            color TEXT,                  -- 颜色
            remark TEXT,                 -- 备注

            carton_count REAL,           -- 件数
            pack_per_carton REAL,        -- 装数
            quantity REAL,               -- 数量
            unit_price REAL,             -- 单价
            amount REAL,                 -- 总金额

            gross_weight REAL,           -- 总毛重
            net_weight REAL,             -- 总净重
            volume REAL,                 -- 总体积
            
            allocated_cost REAL,         -- 分摊费用
            tax_rate REAL,               -- 税率

            FOREIGN KEY (container_id) REFERENCES containers(id)
        )
        """ )
        
        # 尝试添加 allocated_cost 列（兼容旧库）
        try:
            c.execute("ALTER TABLE products ADD COLUMN allocated_cost REAL")
        except sqlite3.OperationalError:
            pass

        # 尝试添加 tax_rate 列（兼容旧库）
        try:
            c.execute("ALTER TABLE products ADD COLUMN tax_rate REAL")
        except sqlite3.OperationalError:
            pass

        self.conn.commit()

    # ---------- 工具函数 ----------

    def parse_shipment_code(self, file_path: str):
        """
        从文件名中提取指令号，例如：... ZL202505.xlsx
        """
        basename = os.path.basename(file_path)
        m = re.search(r"(ZKP|ZL|ZK)\d+", basename)
        if m:
            return m.group(0)
        return None

    # ---------- 导入 Excel ----------

    def import_excel(self, file_path: str, special_linkage: bool = False):
        """
        导入一个 Excel 文件
        """
        xls = pd.ExcelFile(file_path)
        
        # 识别需要读取的Sheet
        target_sheets = []
        for sheet_name in xls.sheet_names:
            # 只要包含"报关"或"清单"字样，或者整个文件只有一个Sheet
            if "报关" in sheet_name or "清单" in sheet_name:
                target_sheets.append(sheet_name)
        
        if not target_sheets:
            # 如果没有找到匹配的名称，默认读取第一个
            target_sheets = [xls.sheet_names[0]]

        shipment = self.parse_shipment_code(file_path)
        file_name = os.path.basename(file_path)

        # 1. 预读取所有Sheet，统计每个货柜涉及的税率集合
        # 结构: { "ContainerA": {0.09, 0.13}, "ContainerB": {0.13} }
        container_tax_map = defaultdict(set)
        
        # 暂存预处理后的数据，避免重复读取Excel: [(df, tax_rate), ...]
        prepared_data = []

        for sheet_name in target_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df = df.dropna(how="all")
            
            # --- 识别税率 ---
            # 1. 优先从Sheet名称获取: "报关清单9%", "报关清单 13%"
            tax_rate = 0.13 # 默认13%
            m_sheet = re.search(r"(\d+)%", sheet_name)
            if m_sheet:
                tax_rate = float(m_sheet.group(1)) / 100.0
            else:
                # 2. 从列名获取: "开票金额9%", "开票金额13%"
                found_in_col = False
                for col in df.columns:
                    col_str = str(col)
                    m_col = re.search(r"开票金额.*?(\d+)%", col_str)
                    if m_col:
                        tax_rate = float(m_col.group(1)) / 100.0
                        found_in_col = True
                        break
            
            df = self._normalize_product_columns(df)

            if '集装箱号' not in df.columns:
                df['集装箱号'] = None
            df['集装箱号'] = df['集装箱号'].ffill()
            
            # 记录该Sheet中的货柜号及其对应的税率
            # 注意: 可能存在None的货柜号（未填写），这里先简单处理
            unique_cns = df['集装箱号'].dropna().unique()
            if len(unique_cns) == 0:
                # 如果没有货柜号，视为 None 货柜
                container_tax_map[None].add(tax_rate)
            else:
                for cn in unique_cns:
                    container_tax_map[cn].add(tax_rate)

            prepared_data.append((df, tax_rate))

        # 2. 正式导入
        for df, tax_rate in prepared_data:
            container_nos = df['集装箱号'].dropna().unique()
            if len(container_nos) == 0:
                container_nos = [None]

            for cn in container_nos:
                if cn is not None:
                    block_df = df[df['集装箱号'] == cn]
                else:
                    block_df = df
                
                # 逻辑判断：如果该货柜在所有Sheet中只涉及一种税率，则不加后缀；
                # 如果涉及多种税率，则必须拆分，加后缀区分。
                import_cn = cn
                rates_set = container_tax_map[cn]
                
                if len(rates_set) > 1 and cn is not None:
                    # 有多种税率，拆分
                    rate_percent = int(round(tax_rate * 100))
                    import_cn = f"{str(cn).strip()}({rate_percent}%)"
                else:
                    # 只有一种税率，或者 cn is None，保持原样
                    # 这样可以满足"同一税率不用分开，只有两个税率才分柜号"的需求
                    pass

                self._import_container_block(block_df, shipment, import_cn, file_name, special_linkage=special_linkage, tax_rate=tax_rate)

        self.conn.commit()

    def _normalize_product_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """识别常见变体列名，规范为系统使用的列名。"""
        if df is None or df.empty:
            return df

        col_map = {}
        has_carton = "件数" in df.columns
        has_pack = "装数" in df.columns
        has_name = "名称" in df.columns
        has_qty = "数量" in df.columns
        has_price = "单价" in df.columns
        has_amount = "总金额" in df.columns

        for col in df.columns:
            col_str = str(col).strip()
            norm = re.sub(r"\s+", "", col_str)
            if not has_carton:
                if re.search(r"(件数|箱数|CTN|carton)", norm, flags=re.IGNORECASE):
                    if "装" not in norm:
                        col_map[col] = "件数"
                        has_carton = True
                        continue
            if not has_pack:
                if re.search(r"(装数|装箱数|装箱数量|每箱|每箱数|PACK|装箱)", norm, flags=re.IGNORECASE):
                    col_map[col] = "装数"
                    has_pack = True
            
            if not has_name:
                if norm in ["品名", "品目名", "商品名称", "货物名称", "DESC", "DESCRIPTION"]:
                    col_map[col] = "名称"
                    has_name = True
            
            if not has_qty:
                if norm in ["数量", "实发数量", "QTY", "QUANTITY"]:
                    col_map[col] = "数量"
                    has_qty = True
            
            if not has_price:
                if norm in ["单价", "报关单价", "PRICE", "UNITPRICE"]:
                    col_map[col] = "单价"
                    has_price = True

            if not has_amount:
                if norm in ["总金额", "金额", "AMOUNT", "TOTALAMOUNT"]:
                    col_map[col] = "总金额"
                    has_amount = True

        if col_map:
            df = df.rename(columns=col_map)
        return df

    def _extract_fees_from_block(self, block_df: pd.DataFrame) -> dict:
        """
        解析费用
        """
        fees = {
            "tax_refund": None,
            "sea_freight_usd": None,
            "all_in_rmb": None,
            "insurance_usd": None, # 改为 USD
            "exchange_rate": None,
            "agency_fee_rmb": None,
            "misc_rmb": None,
            "misc_total_rmb": None,
        }

        cols = list(block_df.columns)

        def find_by_label(label: str):
            for _, row in block_df.iterrows():
                for col in cols:
                    val = row[col]
                    if isinstance(val, str) and label in val:
                        start = cols.index(col) + 1
                        for c2 in cols[start:]:
                            v2 = row[c2]
                            if isinstance(v2, (int, float, np.number)) and not pd.isna(v2):
                                return float(v2)
                        for c2 in cols:
                            v2 = row[c2]
                            if isinstance(v2, (int, float, np.number)) and not pd.isna(v2):
                                return float(v2)
            return None

        fees["tax_refund"] = find_by_label("退税额")

        if "海运费" in block_df.columns:
            sea_vals = block_df["海运费"].dropna()
            if not sea_vals.empty:
                fees["sea_freight_usd"] = float(sea_vals.iloc[0])
        if fees["sea_freight_usd"] is None:
            v = find_by_label("海运费")
            if v is not None:
                fees["sea_freight_usd"] = v

        if "包干费" in block_df.columns:
            vals = block_df["包干费"].dropna()
            if not vals.empty:
                fees["all_in_rmb"] = float(vals.iloc[0])
        
        # 保费 (视为USD)
        fees["insurance_usd"] = find_by_label("保费")
        if fees["insurance_usd"] is None:
            fees["insurance_usd"] = find_by_label("保险")

        fees["exchange_rate"] = find_by_label("汇率")

        ag = find_by_label("代理费")
        if ag is None:
            ag = find_by_label("佣金")
        fees["agency_fee_rmb"] = ag

        fees["misc_rmb"] = find_by_label("杂费")

        # 计算杂费汇总
        # 公式：包干费 + 代理费 + 杂费 + (海运费 * 汇率) + (保费 * 汇率)
        # 不包含退税额
        fees["misc_total_rmb"] = self.calculate_misc_total(fees)

        return fees

    def calculate_misc_total(self, fees: dict) -> float:
        parts = []
        if fees.get("all_in_rmb") is not None:
            parts.append(fees["all_in_rmb"])
        if fees.get("agency_fee_rmb") is not None:
            parts.append(fees["agency_fee_rmb"])
        if fees.get("misc_rmb") is not None:
            parts.append(fees["misc_rmb"])
            
        rate = fees.get("exchange_rate")
        if rate is not None:
            if fees.get("sea_freight_usd") is not None:
                parts.append(fees["sea_freight_usd"] * rate)
            if fees.get("insurance_usd") is not None:
                parts.append(fees["insurance_usd"] * rate)
        
        return float(sum(parts)) if parts else 0.0

    def _insert_or_get_container(self, shipment_code, container_no, file_name, fees):
        c = self.conn.cursor()
        # 注意：这里如果 insurance_usd 是新增列，旧代码的 insert 可能会有问题，需要写完整
        c.execute("""
        INSERT INTO containers (
            shipment_code, container_no, file_name,
            tax_refund, sea_freight_usd, all_in_rmb,
            insurance_usd, exchange_rate, agency_fee_rmb,
            misc_rmb, misc_total_rmb
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(shipment_code, container_no) DO UPDATE SET
            tax_refund      = excluded.tax_refund,
            sea_freight_usd = excluded.sea_freight_usd,
            all_in_rmb      = excluded.all_in_rmb,
            insurance_usd   = excluded.insurance_usd,
            exchange_rate   = excluded.exchange_rate,
            agency_fee_rmb  = excluded.agency_fee_rmb,
            misc_rmb        = excluded.misc_rmb,
            misc_total_rmb  = excluded.misc_total_rmb
        """, (
            shipment_code, container_no, file_name,
            fees["tax_refund"], fees["sea_freight_usd"], fees["all_in_rmb"],
            fees["insurance_usd"], fees["exchange_rate"], fees["agency_fee_rmb"],
            fees["misc_rmb"], fees["misc_total_rmb"]
        ))
        self.conn.commit()

        c.execute(
            "SELECT id FROM containers WHERE shipment_code IS ? AND container_no IS ? ORDER BY id DESC LIMIT 1",
            (shipment_code, container_no)
        )
        row = c.fetchone()
        if row is None:
            return c.lastrowid
        return row["id"]

    def _is_product_row(self, row: pd.Series) -> bool:
        text_vals = []
        for key in ["厂家", "名称", "型号", "件数", "装数", "数量", "单价", "总金额"]:
            if key in row.index:
                v = row[key]
                if isinstance(v, str):
                    text_vals.append(v)
        text_join = "".join(text_vals)

        keywords = [
            "总计", "报关总美金", "CFR价", "开票金额",
            "出口国", "以下开票",
            "退税额", "海运费", "保费", "汇率", "代理费", "杂费"
        ]
        if any(k in text_join for k in keywords):
            return False

        name = row.get("名称", None)
        model = row.get("型号", None)
        if (pd.isna(name) or str(name).strip() == "") and \
           (pd.isna(model) or str(model).strip() == ""):
            return False

        def to_float(x):
            try:
                f = float(x)
                if math.isnan(f):
                    return None
                return f
            except Exception:
                return None

        qty = to_float(row.get("数量"))
        price = to_float(row.get("单价"))
        amount = to_float(row.get("总金额"))
        if qty is None and price is None and amount is None:
            return False

        return True

    def _import_container_block(self, block_df, shipment_code, container_no, file_name, special_linkage: bool = False, tax_rate: float = 0.13):
        fees = self._extract_fees_from_block(block_df)
        container_id = self._insert_or_get_container(
            shipment_code, container_no, file_name, fees
        )

        c = self.conn.cursor()

        for _, row in block_df.iterrows():
            if not self._is_product_row(row):
                continue

            def to_float(x):
                try:
                    f = float(x)
                    if math.isnan(f):
                        return None
                    return f
                except Exception:
                    return None

            p_name = None if pd.isna(row.get("名称")) else str(row.get("名称"))
            qty = to_float(row.get("数量"))
            price = to_float(row.get("单价"))
            amount = to_float(row.get("总金额"))
            
            # 保证金额不变：如果Excel有总金额则用总金额，否则先根据原始单价数量算出总金额
            if amount is None and qty is not None and price is not None:
                amount = qty * price

            # 特殊联动逻辑：如果是鞋，数量/12，单价*12
            if special_linkage and p_name and "鞋" in p_name:
                if qty is not None:
                    qty = qty / 12.0
                if price is not None:
                    price = price * 12.0

            c.execute("""
            INSERT INTO products (
                container_id, factory, name, model, color, remark,
                carton_count, pack_per_carton, quantity, unit_price, amount,
                gross_weight, net_weight, volume, allocated_cost, tax_rate
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,0,?)
            """, (
                container_id,
                None if pd.isna(row.get("厂家")) else str(row.get("厂家")),
                p_name,
                None if pd.isna(row.get("型号")) else str(row.get("型号")),
                None if pd.isna(row.get("颜色")) else str(row.get("颜色")),
                None if pd.isna(row.get("备注")) else str(row.get("备注")),
                to_float(row.get("件数")),
                to_float(row.get("装数")),
                qty,
                price,
                amount,
                to_float(row.get("总毛重")),
                to_float(row.get("总净重")),
                to_float(row.get("总体积")),
                tax_rate
            ))

        self.conn.commit()

    # ---------- 业务逻辑：费用分摊 ----------
    
    def allocate_misc_fees(self, container_id: int):
        """
        根据 货柜的杂费汇总 和 产品的体积占比，分摊费用
        """
        c = self.conn.cursor()
        # 1. 获取货柜信息
        c.execute("SELECT misc_total_rmb FROM containers WHERE id=?", (container_id,))
        row = c.fetchone()
        if not row:
            return
        total_fees = row["misc_total_rmb"] or 0.0

        # 2. 获取该货柜所有产品
        c.execute("SELECT id, volume FROM products WHERE container_id=?", (container_id,))
        products = c.fetchall()
        if not products:
            return

        # 3. 计算总体积
        total_volume = sum((p["volume"] or 0.0) for p in products)

        # 4. 分摊
        for p in products:
            p_vol = p["volume"] or 0.0
            if total_volume > 0:
                ratio = p_vol / total_volume
                alloc = total_fees * ratio
            else:
                alloc = 0.0
            
            c.execute("UPDATE products SET allocated_cost=? WHERE id=?", (alloc, p["id"]))
        
        self.conn.commit()

    def update_container_fees(self, container_id: int, fees: dict):
        """
        更新货柜费用信息，并重新计算 misc_total_rmb
        """
        # 重新计算汇总
        new_total = self.calculate_misc_total(fees)
        fees["misc_total_rmb"] = new_total
        
        c = self.conn.cursor()
        c.execute("""
            UPDATE containers SET
                tax_refund=?, sea_freight_usd=?, all_in_rmb=?,
                insurance_usd=?, exchange_rate=?, agency_fee_rmb=?,
                misc_rmb=?, misc_total_rmb=?
            WHERE id=?
        """, (
            fees["tax_refund"], fees["sea_freight_usd"], fees["all_in_rmb"],
            fees["insurance_usd"], fees["exchange_rate"], fees["agency_fee_rmb"],
            fees["misc_rmb"], fees["misc_total_rmb"],
            container_id
        ))
        self.conn.commit()
        return new_total

    # ---------- 数据管理 ----------

    def clear_data(self):
        c = self.conn.cursor()
        c.execute("DELETE FROM products")
        c.execute("DELETE FROM containers")
        self.conn.commit()

    def delete_container(self, container_id):
        c = self.conn.cursor()
        c.execute("DELETE FROM products WHERE container_id=?", (container_id,))
        c.execute("DELETE FROM containers WHERE id=?", (container_id,))
        self.conn.commit()

    def delete_product(self, product_id):
        c = self.conn.cursor()
        c.execute("DELETE FROM products WHERE id=?", (product_id,))
        self.conn.commit()

    def update_product_field(self, product_id: int, field_name: str, value):
        """
        更新单个产品字段
        """
        # 确保只更新允许的字段
        allowed_fields = [
            "factory", "name", "model", "color", "remark",
            "carton_count", "pack_per_carton", "quantity", "unit_price", "amount",
            "gross_weight", "net_weight", "volume", "allocated_cost", "tax_rate"
        ]
        if field_name not in allowed_fields:
            raise ValueError(f"不允许修改字段: {field_name}")

        c = self.conn.cursor()
        sql = f"UPDATE products SET {field_name}=? WHERE id=?"
        c.execute(sql, (value, product_id))
        self.conn.commit()


    # ---------- 查询接口 ----------

    def query_products(self, keyword: str = None, container_no: str = None):
        c = self.conn.cursor()
        # 增加查询字段：id, gross_weight, net_weight, volume, allocated_cost
        sql = """
        SELECT products.id,
               containers.shipment_code, containers.container_no,
               products.factory, products.name, products.model,
               products.color, products.carton_count, products.pack_per_carton,
               products.quantity, products.unit_price,
               products.amount,
               products.gross_weight, products.net_weight, products.volume,
               products.allocated_cost, products.tax_rate
        FROM products
        JOIN containers ON products.container_id = containers.id
        WHERE 1=1
        """
        params = []
        if keyword:
            sql += " AND (products.name LIKE ? OR products.model LIKE ?)"
            like = f"%{keyword}%"
            params.extend([like, like])
        if container_no:
            sql += " AND containers.container_no LIKE ?"
            params.append(f"%{container_no}%")
        sql += " ORDER BY containers.container_no, products.name"
        c.execute(sql, params)
        return c.fetchall()

    def query_containers(self, container_no: str = None):
        c = self.conn.cursor()
        sql = """
        SELECT c.id, c.shipment_code, c.container_no, c.file_name,
               c.tax_refund, c.sea_freight_usd, c.all_in_rmb,
               c.insurance_usd, c.exchange_rate, c.agency_fee_rmb,
               c.misc_rmb, c.misc_total_rmb,
               (SELECT p.tax_rate FROM products p WHERE p.container_id = c.id LIMIT 1) as tax_rate
        FROM containers c WHERE 1=1
        """
        params = []
        if container_no:
            sql += " AND c.container_no LIKE ?"
            params.append(f"%{container_no}%")
        sql += " ORDER BY c.shipment_code, c.container_no"
        c.execute(sql, params)
        return c.fetchall()


# ======================== GUI 层 ======================== 

class ShippingModule(ttk.Frame):
    def __init__(self, master, db_path="shipping.bd", open_in_converter=None):
        super().__init__(master)
        self.db = ShippingDB(db_path)
        self.clipboard_data = None # 用于剪贴板
        self.open_in_converter = open_in_converter
        self.product_filters = {}
        self.product_sort_col = None
        self.product_sort_desc = False
        self.container_filters = {}
        self.container_sort_col = None
        self.container_sort_desc = False
        self.product_header_selected = set()
        self.product_header_last_index = None
        self.product_drag_start_index = None
        self.product_drag_active = False
        self.product_drag_col_name = None
        self.product_drag_allowed = True
        self.container_header_selected = set()
        self.container_header_last_index = None
        self.container_drag_start_index = None
        self.container_drag_active = False
        self.container_drag_col_name = None
        self.container_drag_allowed = True
        self.product_last_cell = None
        self.container_last_cell = None
        self.product_export_format_var = tk.StringVar(value=get_active_export_format_name("shipping_product"))
        self.container_export_format_var = tk.StringVar(value=get_active_export_format_name("shipping_container"))
        self.special_linkage_var = tk.BooleanVar(value=True) # 特殊品名联动
        self._load_filter_state()
        self._create_widgets()
        self._refresh_export_format_options()

    def _create_widgets(self):
        # 顶部工具条
        top = ttk.Frame(self)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Button(top, text="导入Excel", command=self.on_import).pack(side="left", padx=2)
        ttk.Button(top, text="清空所有数据", command=self.on_clear_all).pack(side="left", padx=2)
        
        ttk.Label(top, text=" | ").pack(side="left", padx=5)

        ttk.Checkbutton(
            top, 
            text="品名联动(鞋:数量/12,单价*12,金额不变)", 
            variable=self.special_linkage_var,
            command=self._save_filter_state
        ).pack(side="left", padx=5)

        ttk.Label(top, text=" | ").pack(side="left", padx=5)
        
        ttk.Label(top, text="数据库:").pack(side="left")
        ttk.Label(top, text=self.db.db_path, foreground="gray").pack(side="left", padx=2)

        ttk.Button(top, text="保存", command=self.on_save_and_exit).pack(side="right", padx=5)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self._build_product_tab()
        self._build_container_tab()
        self._build_log_tab()

        # Auto-load data on startup
        self.refresh_container_table()
        self.refresh_product_table()

    # ---------- 产品查询页 ----------

    def _build_product_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="产品查询")

        # 操作栏
        action_frame = ttk.Frame(frame)
        action_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(action_frame, text="关键字:").pack(side="left")
        self.prod_keyword_var = tk.StringVar()
        ttk.Entry(action_frame, textvariable=self.prod_keyword_var, width=15).pack(side="left", padx=2)

        ttk.Label(action_frame, text="货柜号:").pack(side="left")
        self.prod_container_var = tk.StringVar()
        ttk.Entry(action_frame, textvariable=self.prod_container_var, width=15).pack(side="left", padx=2)

        ttk.Button(action_frame, text="查询", command=self.refresh_product_table).pack(side="left", padx=5)
        
        ttk.Label(action_frame, text="|").pack(side="left", padx=5)
        ttk.Button(action_frame, text="删除选中", command=self.on_delete_product).pack(side="left", padx=2)
        ttk.Button(action_frame, text="导出Excel", command=self.on_export_product).pack(side="left", padx=2)
        ttk.Button(action_frame, text="导出并转入凭证转换", command=lambda: self.on_export_product(True)).pack(side="left", padx=2)
        ttk.Button(action_frame, text="清除筛选", command=self.on_clear_product_filters).pack(side="left", padx=2)
        ttk.Label(action_frame, text="|").pack(side="left", padx=5)
        ttk.Label(action_frame, text="导出格式:").pack(side="left")
        self.product_export_format_combo = ttk.Combobox(
            action_frame,
            textvariable=self.product_export_format_var,
            values=get_export_format_names("shipping_product"),
            state="readonly",
            width=12
        )
        self.product_export_format_combo.pack(side="left", padx=2)
        self.product_export_format_combo.bind("<<ComboboxSelected>>", self._on_product_export_format_changed)
        ttk.Button(action_frame, text="设置", command=self._open_product_export_format_editor).pack(side="left", padx=2)

        # 表格
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # 增加字段
        columns = (
            "id", "shipment_code", "container_no", "factory",
            "name", "model", "color",
            "carton_count", "pack_per_carton", "quantity",
            "unit_price", "unit_price_no_tax", "amount", "amount_no_tax",
            "gross_weight", "net_weight", "volume", "allocated_cost",
            "unit_allocated_cost", "unit_inventory_cost", "inventory_cost", "tax_rate"
        )
        self.product_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.product_tree._skip_smart_restore_header_menu = True

        # 将数据库字段名映射到 Treeview 列名
        self.product_column_map = {
            "id": "id",
            "指令号": "shipment_code",
            "货柜号": "container_no",
            "厂家": "factory",
            "品名": "name",
            "型号": "model",
            "颜色": "color",
            "件数": "carton_count",
            "装数": "pack_per_carton",
            "数量": "quantity",
            "单价": "unit_price",
            "不含税单价": "unit_price_no_tax",
            "总金额": "amount",
            "不含税总金额": "amount_no_tax",
            "总毛重": "gross_weight",
            "总净重": "net_weight",
            "总体积": "volume",
            "分摊杂费": "allocated_cost",
            "单个分摊杂费": "unit_allocated_cost",
            "单个库存成本": "unit_inventory_cost",
            "库存成本": "inventory_cost",
            "税率": "tax_rate"
        }
        # 反向映射用于更新
        self.product_column_reverse_map = {v: k for k, v in self.product_column_map.items()}

        for col in columns:
            text_heading = self.product_column_reverse_map.get(col, col) # 获取中文标题
            self.product_tree.heading(col, text=text_heading, command=lambda c=col: self.on_product_heading_click(c))
            w = 80
            if col in ["name", "model"]: w = 120
            if col == "id": w = 40
            if col in ["unit_price_no_tax", "amount_no_tax", "unit_allocated_cost", "unit_inventory_cost", "inventory_cost"]:
                w = 110
            self.product_tree.column(col, width=w, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.product_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.product_tree.xview)
        self.product_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.product_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # 绑定双击事件进行编辑
        self.product_tree.bind("<Double-1>", self.on_product_cell_double_click)
        
        # 使用增强版 Treeview 工具 (支持单元格级复制粘贴)
        attach_treeview_tools(self.product_tree)
        self.product_tree.bind("<<TreeviewPaste>>", self.on_product_paste_batch)
        self.product_tree.bind("<<TreeviewRefresh>>", lambda e: self.refresh_product_table())
        
        self.product_tree.bind("<Button-1>", self.on_product_cell_click, add="+")
        self.product_tree.bind("<ButtonPress-1>", self.on_product_header_press, add="+")
        self.product_tree.bind("<B1-Motion>", self.on_product_header_drag, add="+")
        self.product_tree.bind("<ButtonRelease-1>", self.on_product_header_release, add="+")
        self.product_tree.bind("<Button-3>", self.on_product_heading_right_click, add="+")

    # ---------- 货柜查询页 ----------

    def _build_container_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="货柜查询")

        # 顶部操作
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="货柜号:").pack(side="left")
        self.cont_keyword_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.cont_keyword_var, width=15).pack(side="left", padx=2)
        ttk.Button(top_frame, text="查询", command=self.refresh_container_table).pack(side="left", padx=5)

        ttk.Label(top_frame, text="|").pack(side="left", padx=5)
        ttk.Button(top_frame, text="删除选中货柜", command=self.on_delete_container).pack(side="left", padx=2)
        ttk.Button(top_frame, text="导出Excel", command=self.on_export_container).pack(side="left", padx=2)
        ttk.Button(top_frame, text="导出并转入凭证转换", command=lambda: self.on_export_container(True)).pack(side="left", padx=2)
        ttk.Label(top_frame, text="|").pack(side="left", padx=5)
        ttk.Label(top_frame, text="导出格式:").pack(side="left")
        self.container_export_format_combo = ttk.Combobox(
            top_frame,
            textvariable=self.container_export_format_var,
            values=get_export_format_names("shipping_container"),
            state="readonly",
            width=12
        )
        self.container_export_format_combo.pack(side="left", padx=2)
        self.container_export_format_combo.bind("<<ComboboxSelected>>", self._on_container_export_format_changed)
        ttk.Button(top_frame, text="设置", command=self._open_container_export_format_editor).pack(side="left", padx=2)
        ttk.Button(top_frame, text="清除筛选", command=self.on_clear_container_filters).pack(side="left", padx=2)

        # 费用编辑区域
        info_frame = ttk.LabelFrame(frame, text="货柜费用管理 (修改后请按'保存并汇总')")
        info_frame.pack(fill="x", padx=5, pady=5)

        self.container_info_vars = {}
        # 映射: 字段名 -> 中文显示
        labels = [
            ("tax_refund",      "退税额"),
            ("sea_freight_usd", "海运费($)"),
            ("all_in_rmb",      "包干费(￥)"),
            ("insurance_usd",   "保费($)"),
            ("exchange_rate",   "汇率"),
            ("agency_fee_rmb",  "代理费(￥)"),
            ("misc_rmb",        "其他杂费(￥)"),
            ("misc_total_rmb",  "杂费汇总(￥)"), # 只读
        ]

        grid_frame = ttk.Frame(info_frame)
        grid_frame.pack(fill="x", padx=5, pady=5)

        for i, (key, text) in enumerate(labels):
            row = i // 4
            col = (i % 4) * 2
            ttk.Label(grid_frame, text=text + ":").grid(row=row, column=col, sticky="e", padx=4, pady=5)
            
            var = tk.DoubleVar()
            self.container_info_vars[key] = var
            
            if key == "misc_total_rmb":
                # 汇总只读
                entry = ttk.Entry(grid_frame, textvariable=var, width=12, state="readonly")
            else:
                entry = ttk.Entry(grid_frame, textvariable=var, width=12)
            
            entry.grid(row=row, column=col + 1, sticky="w", padx=4, pady=5)

        # 功能按钮区
        btn_frame = ttk.Frame(info_frame)
        btn_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Button(btn_frame, text="保存修改并汇总", command=self.on_save_container_fees).pack(side="left", padx=20)
        ttk.Button(btn_frame, text=">> 执行分摊 (到产品) <<", command=self.on_allocate_click).pack(side="left", padx=20)
        ttk.Label(btn_frame, text="提示：分摊将根据产品体积占比，将【杂费汇总】分配到各产品的【分摊费用】字段").pack(side="left", padx=10)

        # 货柜表格
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = (
            "id", "shipment_code", "container_no", "file_name",
            "tax_refund", "sea_freight_usd", "all_in_rmb",
            "insurance_usd", "exchange_rate", "agency_fee_rmb",
            "misc_rmb", "misc_total_rmb", "tax_rate"
        )
        self.container_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.container_tree._skip_smart_restore_header_menu = True

        headings = {
            "id": "ID",
            "shipment_code":   "指令号",
            "container_no":    "货柜号",
            "file_name":       "来源文件",
            "tax_refund":      "退税额",
            "sea_freight_usd": "海运费($)",
            "all_in_rmb":      "包干费(￥)",
            "insurance_usd":   "保费($)",
            "exchange_rate":   "汇率",
            "agency_fee_rmb":  "代理费(￥)",
            "misc_rmb":        "其他杂费(￥)",
            "misc_total_rmb":  "杂费汇总(￥)",
            "tax_rate":        "税率",
        }
        for col in columns:
            self.container_tree.heading(col, text=headings[col], command=lambda c=col: self.on_container_heading_click(c))
            w = 90
            if col == "id": w = 40
            self.container_tree.column(col, width=w, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.container_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.container_tree.xview)
        self.container_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.container_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.container_tree.bind("<<TreeviewSelect>>", self.on_container_select)
        self.container_tree.bind("<Button-3>", self.on_container_heading_right_click, add="+")
        
        # 使用增强版 Treeview 工具 (支持单元格级复制粘贴)
        attach_treeview_tools(self.container_tree)
        self.container_tree.bind("<<TreeviewPaste>>", self.on_container_paste_batch)
        self.container_tree.bind("<<TreeviewRefresh>>", lambda e: self.refresh_container_table())

        self.container_tree.bind("<Button-1>", self.on_container_cell_click, add="+")

    def _build_log_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="日志")
        self.log_text = tk.Text(frame, height=10)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    # ======================== 事件处理 ======================== 

    def on_import(self):
        paths = filedialog.askopenfilenames(title="选择Excel", filetypes=[("Excel", "*.xlsx *.xls")])
        if not paths: return
        linkage_enabled = self.special_linkage_var.get()
        for p in paths:
            try:
                self.db.import_excel(p, special_linkage=linkage_enabled)
                self.log(f"导入成功: {os.path.basename(p)}")
            except Exception as e:
                self.log(f"导入失败 {p}: {e}")
        messagebox.showinfo("完成", "导入完成")
        self.refresh_container_table()
        self.refresh_product_table()

    def on_clear_all(self):
        if messagebox.askyesno("警告", "确定要清空所有数据吗？此操作不可恢复！"):
            self.db.clear_data()
            self.refresh_container_table()
            self.refresh_product_table()
            self.log("已清空所有数据")

    # --- 产品页操作 ---

    def refresh_product_table(self):
        for i in self.product_tree.get_children():
            self.product_tree.delete(i)
        
        kw = self.prod_keyword_var.get().strip()
        ct = self.prod_container_var.get().strip()
        row_entries = self._get_product_row_entries(kw, ct)
        
        # 定义哪些列需要格式化为两位小数
        decimal_cols = self._get_product_decimal_cols()
        
        # 统计汇总
        sums = {k: 0.0 for k in decimal_cols}
        for key in ["unit_price", "unit_price_no_tax", "unit_allocated_cost", "unit_inventory_cost"]:
            sums.pop(key)

        # 过滤
        row_entries = self._apply_filters(row_entries, self.product_filters, "product")

        # 排序
        if self.product_sort_col:
            row_entries.sort(
                key=lambda e: self._sort_key(e.get(self.product_sort_col)),
                reverse=self.product_sort_desc
            )

        formatted_rows = []
        for entry in row_entries:
            values = []
            for col_name in self.product_tree["columns"]:
                val = entry.get(col_name)
                
                # 累加
                if col_name in sums and isinstance(val, (int, float)):
                    sums[col_name] += val

                if col_name in decimal_cols and isinstance(val, (int, float)):
                    values.append(f"{val:.2f}")
                elif val is None:
                    values.append("") # 显示空字符串
                else:
                    values.append(str(val))
            formatted_rows.append(values)

        # 插入汇总行
        if formatted_rows:
            cols = self.product_tree["columns"]
            summary_vals = [""] * len(cols)
            summary_vals[0] = "汇总"
            
            for k, v in sums.items():
                if k in cols:
                    idx = cols.index(k)
                    summary_vals[idx] = f"{v:.2f}"
            
            self.product_tree.insert("", "end", values=summary_vals, tags=("summary",))
            self.product_tree.tag_configure("summary", background="#e1e1e1", font=("System", 10, "bold"))

        for vals in formatted_rows:
            self.product_tree.insert("", "end", values=vals)

        self._refresh_product_headings()

    def on_delete_product(self):
        sels = self.product_tree.selection()
        if not sels: return
        if not messagebox.askyesno("确认", f"删除选中的 {len(sels)} 条记录？"):
            return
        for item in sels:
            vals = self.product_tree.item(item, "values")
            pid = vals[0] # id
            self.db.delete_product(pid)
        self.refresh_product_table()

    def on_export_product(self, open_in_converter=False):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not f: return
        
        # 获取当前视觉显示的数据（尊重列排序和隐藏）
        if hasattr(self.product_tree, "_treeview_tools"):
            columns, data = self.product_tree._treeview_tools.get_visual_data()
        else:
            # 回退方案
            data = []
            columns = [self.product_column_reverse_map[col] for col in self.product_tree["columns"]]
            for item in self.product_tree.get_children():
                data.append(self.product_tree.item(item, "values"))

        columns, data, mapped = apply_export_format("shipping_product", columns, data)
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(f, index=False)
        
        # 增加图表
        add_charts_to_product_report(f)
        
        messagebox.showinfo("成功", "导出成功")
        if open_in_converter and self.open_in_converter:
            self.open_in_converter(f)

    def _on_product_export_format_changed(self, event=None):
        name = self.product_export_format_var.get().strip()
        set_active_export_format("shipping_product", name)

    def _open_product_export_format_editor(self):
        columns = [self.product_column_reverse_map[col] for col in self.product_tree["columns"]]
        open_export_format_editor(
            self,
            "shipping_product",
            columns,
            title="导出格式设置 - 报关清单产品"
        )
        self._refresh_export_format_options()

    def on_product_cell_double_click(self, event):
        # 获取点击的单元格
        region = self.product_tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.product_tree.identify_column(event.x)
        item_id = self.product_tree.identify_row(event.y)

        if not item_id: # 没点击到有效行
            return

        # 获取列的索引 (e.g., #0, #1...)
        col_index = int(column[1:]) - 1 # #0 是 Treeview row ID, #1 是第一个数据列
        
        # 获取当前值和产品ID
        current_values = self.product_tree.item(item_id, "values")
        product_id = current_values[0] # 产品ID是第一列
        
        if str(product_id) == "汇总":
            return

        current_text = current_values[col_index]

        # ID列和指令号/货柜号列不可编辑
        editable_cols = ["factory", "name", "model", "color", "remark",
                         "carton_count", "pack_per_carton", "quantity", "unit_price", "amount",
                         "gross_weight", "net_weight", "volume", "allocated_cost"]
        
        # 将Treeview的列ID转换为数据库字段名
        db_col_name = self.product_tree["columns"][col_index]

        if db_col_name not in editable_cols:
            return

        # 创建一个Entry Widget进行编辑
        bbox = self.product_tree.bbox(item_id, column)
        if bbox == '': # 如果单元格不可见
            return

        entry_edit = ttk.Entry(self.product_tree, width=bbox[2])
        entry_edit.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry_edit.insert(0, current_text)
        entry_edit.focus_set()

        def on_entry_edit_confirm(e):
            new_value_str = entry_edit.get()
            # 尝试转换成数字，如果字段是数字类型
            if db_col_name in ["carton_count", "pack_per_carton", "quantity", "unit_price", "amount",
                               "gross_weight", "net_weight", "volume", "allocated_cost"]:
                try:
                    new_value = float(new_value_str)
                except ValueError:
                    messagebox.showerror("输入错误", f"{self.product_column_reverse_map[db_col_name]} 必须是数字！")
                    entry_edit.destroy()
                    return
            else:
                new_value = new_value_str if new_value_str.strip() != "" else None # 允许清空文本字段

            try:
                self.db.update_product_field(product_id, db_col_name, new_value)
                self.log(f"产品ID {product_id} 的 {self.product_column_reverse_map[db_col_name]} 已更新为: {new_value_str}")
                self.refresh_product_table() # 刷新整个表格以显示更新和重新排序
            except Exception as ex:
                messagebox.showerror("数据库错误", f"更新失败: {ex}")
            finally:
                entry_edit.destroy()

        entry_edit.bind("<Return>", on_entry_edit_confirm)
        entry_edit.bind("<FocusOut>", on_entry_edit_confirm) # 失去焦点也保存

    def on_product_heading_click(self, col_name: str):
        if self.product_sort_col == col_name:
            self.product_sort_desc = not self.product_sort_desc
        else:
            self.product_sort_col = col_name
            self.product_sort_desc = False
        self.refresh_product_table()

    def on_product_heading_right_click(self, event):
        region = self.product_tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        col_id = self.product_tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        if col_index < 0:
            return
        columns = self._get_display_columns(self.product_tree)
        if col_index >= len(columns):
            return
        col_name = columns[col_index]
        menu = tk.Menu(self.product_tree, tearoff=0)
        menu.add_command(label="筛选...", command=lambda: self._open_filter_dialog("product", col_name))
        menu.add_command(label="清除此列筛选", command=lambda: self._clear_product_filter(col_name))
        menu.add_command(label="清除所有筛选", command=self.on_clear_product_filters)
        menu.add_separator()
        if self.product_header_selected:
            menu.add_command(
                label=f"隐藏选中列 ({len(self.product_header_selected)})",
                command=lambda: self._hide_tree_columns(self.product_tree, self.product_header_selected)
            )
        menu.add_command(label="隐藏此列", command=lambda: self._hide_tree_column(self.product_tree, col_name))
        menu.add_command(label="显示全部列", command=lambda: self._show_all_tree_columns(self.product_tree))
        menu.add_command(label="列管理...", command=lambda: self._open_column_manager_dialog(
            self.product_tree, self.product_column_reverse_map
        ))
        menu.add_separator()
        add_smart_restore_menu(menu, self.product_tree)
        menu.tk_popup(event.x_root, event.y_root)

    def on_clear_product_filters(self):
        self.product_filters = {}
        self._save_filter_state()
        self.refresh_product_table()

    def on_product_paste_batch(self, event=None):
        """处理产品明细表的批量粘贴保存"""
        if not messagebox.askyesno("保存确认", "检测到批量粘贴操作，是否将更改保存到数据库？\n\n注意：这将根据第一列的 ID 更新对应的记录。"):
            self.refresh_product_table()
            return
            
        success_count = 0
        error_count = 0
        
        columns = self.product_tree["columns"]
        editable_cols = ["factory", "name", "model", "color", "remark",
                         "carton_count", "pack_per_carton", "quantity", "unit_price", "amount",
                         "gross_weight", "net_weight", "volume", "allocated_cost"]
        num_cols = ["carton_count", "pack_per_carton", "quantity", "unit_price", "amount",
                    "gross_weight", "net_weight", "volume", "allocated_cost"]

        items = self.product_tree.get_children("")
        for iid in items:
            values = self.product_tree.item(iid, "values")
            if not values or str(values[0]) == "汇总":
                continue
                
            try:
                product_id = values[0]
                for i in range(1, len(columns)):
                    col_name = columns[i]
                    if col_name in editable_cols and i < len(values):
                        new_val_str = str(values[i])
                        new_val = new_val_str
                        if col_name in num_cols:
                            try:
                                new_val = float(new_val_str) if new_val_str.strip() else 0.0
                            except ValueError:
                                continue
                        self.db.update_product_field(product_id, col_name, new_val)
                success_count += 1
            except Exception:
                error_count += 1
                
        self.log(f"批量粘贴更新完成：成功 {success_count} 行")
        self.refresh_product_table()

    def on_container_paste_batch(self, event=None):
        """处理货柜列表的批量粘贴保存"""
        if not messagebox.askyesno("保存确认", "检测到批量粘贴操作，是否将更改保存到数据库？"):
            self.refresh_container_table()
            return
            
        success_count = 0
        columns = self.container_tree["columns"]
        # 可编辑列 (大致基于数据库结构)
        editable_cols = ["tax_refund", "sea_freight_usd", "all_in_rmb", "insurance_usd", "exchange_rate", "agency_fee_rmb", "misc_rmb"]
        
        items = self.container_tree.get_children("")
        for iid in items:
            values = self.container_tree.item(iid, "values")
            if not values: continue
            
            try:
                container_id = values[0]
                for i in range(1, len(columns)):
                    col_name = columns[i]
                    if col_name in editable_cols and i < len(values):
                        val_str = str(values[i])
                        try:
                            val = float(val_str) if val_str.strip() else 0.0
                            self.db.update_container_field(container_id, col_name, val)
                        except ValueError:
                            continue
                success_count += 1
            except Exception:
                continue
        
        self.log(f"货柜批量更新完成：{success_count} 行")
        self.refresh_container_table()

    def on_product_cell_click(self, event):
        pass # 使用 TreeviewTools 处理点击
        region = self.product_tree.identify_region(event.x, event.y)
        if region != "cell":
            self.product_last_cell = None
            return
        column_id = self.product_tree.identify_column(event.x)
        item_id = self.product_tree.identify_row(event.y)
        if not item_id:
            self.product_last_cell = None
            return
        col_index = int(column_id[1:]) - 1
        self.product_last_cell = (item_id, col_index)

    def on_container_cell_click(self, event):
        pass # 使用 TreeviewTools 处理点击
        region = self.container_tree.identify_region(event.x, event.y)
        if region != "cell":
            self.container_last_cell = None
            return
        column_id = self.container_tree.identify_column(event.x)
        item_id = self.container_tree.identify_row(event.y)
        if not item_id:
            self.container_last_cell = None
            return
        col_index = int(column_id[1:]) - 1
        self.container_last_cell = (item_id, col_index)


    # --- 货柜页操作 ---

    def refresh_container_table(self):
        for i in self.container_tree.get_children():
            self.container_tree.delete(i)
        
        ct = self.cont_keyword_var.get().strip()
        row_entries = self._get_container_row_entries(ct)
        
        # 定义哪些列需要格式化为两位小数 (所有货币和汇率)
        decimal_cols = self._get_container_decimal_cols()
        
        # 针对汇率，不要格式化为两位小数
        if "exchange_rate" in decimal_cols:
            decimal_cols.remove("exchange_rate")

        # 统计汇总
        sums = {k: 0.0 for k in decimal_cols}
        # sums.pop("exchange_rate") # 汇率已经不在decimal_cols里了，不需要pop

        # 过滤
        row_entries = self._apply_filters(row_entries, self.container_filters, "container")

        # 排序
        if self.container_sort_col:
            row_entries.sort(
                key=lambda e: self._sort_key(e.get(self.container_sort_col)),
                reverse=self.container_sort_desc
            )

        formatted_rows = []
        for entry in row_entries:
            values = []
            for col_name in self.container_tree["columns"]:
                val = entry.get(col_name)
                
                # 累加
                if col_name in sums and isinstance(val, (int, float)):
                    sums[col_name] += val

                if col_name in decimal_cols and isinstance(val, (int, float)):
                    values.append(f"{val:.2f}")
                elif val is None:
                    values.append("")
                else:
                    values.append(str(val))
            formatted_rows.append(values)
        
        # 插入汇总行
        if formatted_rows:
            cols = self.container_tree["columns"]
            summary_vals = [""] * len(cols)
            summary_vals[0] = "汇总"
            
            for k, v in sums.items():
                if k in cols:
                    idx = cols.index(k)
                    summary_vals[idx] = f"{v:.2f}"
            
            self.container_tree.insert("", "end", values=summary_vals, tags=("summary",))
            self.container_tree.tag_configure("summary", background="#e1e1e1", font=("System", 10, "bold"))

        for vals in formatted_rows:
            self.container_tree.insert("", "end", values=vals)

        self._refresh_container_headings()
        
        if formatted_rows:
            # 尝试选中第一行有效数据（跳过汇总行）
            children = self.container_tree.get_children()
            if len(children) > 1:
                self.container_tree.selection_set(children[1])
                self.on_container_select()
            elif len(children) == 1 and children[0] != "汇总": # 应该不会发生，因为有数据就有汇总
                 self.container_tree.selection_set(children[0])
                 self.on_container_select()
        else:
            # 清空输入框
            for v in self.container_info_vars.values():
                v.set(0.0)

    def on_container_select(self, event=None):
        sel = self.container_tree.selection()
        if not sel: return
        vals = self.container_tree.item(sel[0], "values")
        
        if vals[0] == "汇总":
             # 如果选中的是汇总行，清空输入框或不做处理
             for v in self.container_info_vars.values():
                 v.set(0.0)
             return

        # Treeview columns 顺序
        cols = (
            "id", "shipment_code", "container_no", "file_name",
            "tax_refund", "sea_freight_usd", "all_in_rmb",
            "insurance_usd", "exchange_rate", "agency_fee_rmb",
            "misc_rmb", "misc_total_rmb"
        )
        data = dict(zip(cols, vals))
        
        # 填充到输入框
        for k, var in self.container_info_vars.items():
            val = data.get(k)
            try:
                # 确保以浮点数形式填充，便于编辑，显示时再格式化
                if val and val != "None":
                    var.set(float(val))
                else:
                    var.set(0.0)
            except Exception as e:
                self.log(f"Error setting container info var {k}: {e}")
                var.set(0.0)

    def on_container_heading_click(self, col_name: str):
        if self.container_sort_col == col_name:
            self.container_sort_desc = not self.container_sort_desc
        else:
            self.container_sort_col = col_name
            self.container_sort_desc = False
        self.refresh_container_table()

    def on_container_heading_right_click(self, event):
        region = self.container_tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        col_id = self.container_tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        if col_index < 0:
            return
        columns = self._get_display_columns(self.container_tree)
        if col_index >= len(columns):
            return
        col_name = columns[col_index]
        menu = tk.Menu(self.container_tree, tearoff=0)
        menu.add_command(label="筛选...", command=lambda: self._open_filter_dialog("container", col_name))
        menu.add_command(label="清除此列筛选", command=lambda: self._clear_container_filter(col_name))
        menu.add_command(label="清除所有筛选", command=self.on_clear_container_filters)
        menu.add_separator()
        if self.container_header_selected:
            menu.add_command(
                label=f"隐藏选中列 ({len(self.container_header_selected)})",
                command=lambda: self._hide_tree_columns(self.container_tree, self.container_header_selected)
            )
        menu.add_command(label="隐藏此列", command=lambda: self._hide_tree_column(self.container_tree, col_name))
        menu.add_command(label="显示全部列", command=lambda: self._show_all_tree_columns(self.container_tree))
        menu.add_command(label="列管理...", command=lambda: self._open_column_manager_dialog(
            self.container_tree, None
        ))
        menu.add_separator()
        add_smart_restore_menu(menu, self.container_tree)
        menu.tk_popup(event.x_root, event.y_root)

    def on_clear_container_filters(self):
        self.container_filters = {}
        self._save_filter_state()
        self.refresh_container_table()

    def get_selected_container_id(self):
        sel = self.container_tree.selection()
        if not sel: return None
        vals = self.container_tree.item(sel[0], "values")
        return vals[0] # id is first col

    def on_save_container_fees(self):
        cid = self.get_selected_container_id()
        if not cid:
            messagebox.showwarning("提示", "请先选择一个货柜")
            return
        
        fees = {}
        for k, var in self.container_info_vars.items():
            if k == "misc_total_rmb": continue # misc_total_rmb是计算所得，不从输入框取
            try:
                fees[k] = float(var.get())
            except ValueError:
                messagebox.showerror("输入错误", f"'{k}' 必须是数字！")
                return
        
        try:
            new_total = self.db.update_container_fees(cid, fees)
            self.container_info_vars["misc_total_rmb"].set(f"{new_total:.2f}") # 格式化为两位小数
            
            self.refresh_container_table()
            
            # 恢复选中状态
            for item in self.container_tree.get_children():
                if self.container_tree.item(item, "values")[0] == cid:
                    self.container_tree.selection_set(item)
                    break
            
            messagebox.showinfo("成功", "保存并重新计算汇总成功！")
        except Exception as ex:
            messagebox.showerror("数据库错误", f"更新失败: {ex}")

    def on_allocate_click(self):
        cid = self.get_selected_container_id()
        if not cid:
            messagebox.showwarning("提示", "请先选择一个货柜")
            return
        
        try:
            self.db.allocate_misc_fees(cid)
            messagebox.showinfo("成功", "费用分摊完成！请到[产品查询]页查看结果。")
            self.refresh_product_table()
        except Exception as ex:
            messagebox.showerror("错误", f"分摊失败: {ex}")

    def on_delete_container(self):
        sel = self.container_tree.selection()
        if not sel: return
        if not messagebox.askyesno("警告", "删除货柜会同时删除其下所有产品数据！确定吗？"):
            return
        for item in sel:
            cid = self.container_tree.item(item, "values")[0]
            self.db.delete_container(cid)
        self.refresh_container_table()
        self.refresh_product_table()

    def on_export_container(self, open_in_converter=False):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not f: return
        
        # 获取当前视觉显示的数据（尊重列排序和隐藏）
        if hasattr(self.container_tree, "_treeview_tools"):
            columns, data = self.container_tree._treeview_tools.get_visual_data()
        else:
            # 获取所有行
            data = []
            columns = [self.container_column_reverse_map[col] for col in self.container_tree["columns"]]
            for iid in self.container_tree.get_children():
                data.append(self.container_tree.item(iid, "values"))

        columns, data, mapped = apply_export_format("shipping_container", columns, data)
        df = pd.DataFrame(data, columns=columns)
        df.to_excel(f, index=False)
        
        # 增加图表
        add_charts_to_container_report(f)
        
        messagebox.showinfo("成功", "导出成功")
        if open_in_converter and self.open_in_converter:
            self.open_in_converter(f)

    def _on_container_export_format_changed(self, event=None):
        name = self.container_export_format_var.get().strip()
        set_active_export_format("shipping_container", name)

    def _open_container_export_format_editor(self):
        columns = [
            "ID", "指令号", "货柜号", "来源文件",
            "退税额", "海运费USD", "包干费", "保费USD", "汇率",
            "代理费", "杂费", "杂费汇总"
        ]
        open_export_format_editor(
            self,
            "shipping_container",
            columns,
            title="导出格式设置 - 报关清单货柜"
        )
        self._refresh_export_format_options()

    def _refresh_export_format_options(self):
        if hasattr(self, "product_export_format_combo"):
            names = get_export_format_names("shipping_product")
            self.product_export_format_combo["values"] = names
            active = get_active_export_format_name("shipping_product")
            if active:
                self.product_export_format_var.set(active)
        if hasattr(self, "container_export_format_combo"):
            names = get_export_format_names("shipping_container")
            self.container_export_format_combo["values"] = names
            active = get_active_export_format_name("shipping_container")
            if active:
                self.container_export_format_var.set(active)

    def on_save_and_exit(self):
        try:
            self.db.conn.commit()
            messagebox.showinfo("成功", "保存成功")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def _apply_filters(self, entries, filters, table_type):
        if not filters:
            return entries
        if table_type == "product":
            decimal_cols = set(self._get_product_decimal_cols())
        else:
            decimal_cols = set(self._get_container_decimal_cols())
            if "exchange_rate" in decimal_cols:
                decimal_cols.remove("exchange_rate")

        filtered = []
        for entry in entries:
            ok = True
            for col_name, needle in filters.items():
                if not needle:
                    continue
                value = entry.get(col_name)
                if isinstance(needle, dict):
                    mode = needle.get("mode")
                    if mode == "values":
                        selected = set(needle.get("values", []))
                        display = self._display_value_for_filter(table_type, col_name, value, decimal_cols)
                        if display not in selected:
                            ok = False
                            break
                    elif mode == "text":
                        op = needle.get("op", "contains")
                        target = needle.get("value", "")
                        if not self._match_text_filter(value, op, target):
                            ok = False
                            break
                    elif mode == "number":
                        expr = needle.get("expr", "")
                        if not self._match_number_filter(value, expr):
                            ok = False
                            break
                else:
                    if not self._match_text_filter(value, "contains", str(needle)):
                        ok = False
                        break
            if ok:
                filtered.append(entry)
        return filtered

    def _sort_key(self, value):
        if value is None:
            return (1, "")
        if isinstance(value, (int, float)):
            return (0, value)
        return (0, str(value).lower())

    def _refresh_product_headings(self):
        for col in self.product_tree["columns"]:
            base = self.product_column_reverse_map.get(col, col)
            if col in self.product_filters:
                base += " [F]"
            if self.product_sort_col == col:
                base += " v" if self.product_sort_desc else " ^"
            self.product_tree.heading(col, text=base, command=lambda c=col: self.on_product_heading_click(c))

    def _refresh_container_headings(self):
        headings = {
            "id": "ID",
            "shipment_code":   "指令号",
            "container_no":    "货柜号",
            "file_name":       "来源文件",
            "tax_refund":      "退税额",
            "sea_freight_usd": "海运费($)",
            "all_in_rmb":      "包干费(￥)",
            "insurance_usd":   "保费($)",
            "exchange_rate":   "汇率",
            "agency_fee_rmb":  "代理费(￥)",
            "misc_rmb":        "其他杂费(￥)",
            "misc_total_rmb":  "杂费汇总(￥)",
        }
        for col in self.container_tree["columns"]:
            base = headings.get(col, col)
            if col in self.container_filters:
                base += " [F]"
            if self.container_sort_col == col:
                base += " v" if self.container_sort_desc else " ^"
            self.container_tree.heading(col, text=base, command=lambda c=col: self.on_container_heading_click(c))

    def _clear_product_filter(self, col_name):
        if col_name in self.product_filters:
            self.product_filters.pop(col_name, None)
            self._save_filter_state()
            self.refresh_product_table()

    def _clear_container_filter(self, col_name):
        if col_name in self.container_filters:
            self.container_filters.pop(col_name, None)
            self._save_filter_state()
            self.refresh_container_table()

    def _get_display_columns(self, tree):
        display = tree["displaycolumns"]
        if display == "#all":
            return list(tree["columns"])
        if isinstance(display, (list, tuple)):
            if "#all" in display:
                return list(tree["columns"])
            return list(display)
        return list(tree["columns"])

    def _set_display_columns(self, tree, cols):
        if list(cols) == list(tree["columns"]):
            tree["displaycolumns"] = "#all"
        else:
            tree["displaycolumns"] = cols

    def _hide_tree_column(self, tree, col_name):
        display = self._get_display_columns(tree)
        if col_name not in display:
            return
        if len(display) <= 1:
            messagebox.showinfo("提示", "至少保留一列显示。")
            return
        display = [c for c in display if c != col_name]
        self._set_display_columns(tree, display)

    def _hide_tree_columns(self, tree, col_names):
        display = self._get_display_columns(tree)
        remain = [c for c in display if c not in col_names]
        if not remain:
            messagebox.showinfo("提示", "至少保留一列显示。")
            return
        self._set_display_columns(tree, remain)

    def _show_all_tree_columns(self, tree):
        self._set_display_columns(tree, list(tree["columns"]))

    def on_product_header_press(self, event):
        region = self.product_tree.identify_region(event.x, event.y)
        if region != "heading":
            self.product_drag_start_index = None
            self.product_drag_active = False
            self.product_drag_col_name = None
            return
        col_id = self.product_tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        columns = self._get_display_columns(self.product_tree)
        if col_index < 0 or col_index >= len(columns):
            return
        col_name = columns[col_index]
        ctrl = bool(event.state & 0x0004)
        shift = bool(event.state & 0x0001)
        self.product_drag_allowed = not (ctrl or shift)
        self.product_drag_start_index = col_index
        self.product_drag_active = False
        self.product_drag_col_name = col_name
        if shift and self.product_header_last_index is not None:
            start = min(self.product_header_last_index, col_index)
            end = max(self.product_header_last_index, col_index)
            self.product_header_selected.update(columns[start:end + 1])
            return "break"
        if ctrl:
            if col_name in self.product_header_selected:
                self.product_header_selected.remove(col_name)
            else:
                self.product_header_selected.add(col_name)
            self.product_header_last_index = col_index
            return "break"
        self.product_header_selected = {col_name}
        self.product_header_last_index = col_index
        return None

    def on_product_header_drag(self, event):
        if self.product_drag_start_index is None or not self.product_drag_allowed:
            return
        region = self.product_tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        col_id = self.product_tree.identify_column(event.x)
        if not col_id:
            return
        target_index = int(col_id[1:]) - 1
        columns = self._get_display_columns(self.product_tree)
        if target_index < 0 or target_index >= len(columns):
            return
        if target_index == self.product_drag_start_index:
            return
        col_name = self.product_drag_col_name
        if not col_name or col_name not in columns:
            return
        columns.remove(col_name)
        columns.insert(target_index, col_name)
        self._set_display_columns(self.product_tree, columns)
        self.product_drag_start_index = target_index
        self.product_drag_active = True

    def on_product_header_release(self, event):
        if self.product_drag_active:
            self.product_drag_active = False
            self.product_drag_start_index = None
            self.product_drag_col_name = None
            return "break"
        self.product_drag_start_index = None
        self.product_drag_col_name = None
        return None

    def on_container_header_press(self, event):
        region = self.container_tree.identify_region(event.x, event.y)
        if region != "heading":
            self.container_drag_start_index = None
            self.container_drag_active = False
            self.container_drag_col_name = None
            return
        col_id = self.container_tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        columns = self._get_display_columns(self.container_tree)
        if col_index < 0 or col_index >= len(columns):
            return
        col_name = columns[col_index]
        ctrl = bool(event.state & 0x0004)
        shift = bool(event.state & 0x0001)
        self.container_drag_allowed = not (ctrl or shift)
        self.container_drag_start_index = col_index
        self.container_drag_active = False
        self.container_drag_col_name = col_name
        if shift and self.container_header_last_index is not None:
            start = min(self.container_header_last_index, col_index)
            end = max(self.container_header_last_index, col_index)
            self.container_header_selected.update(columns[start:end + 1])
            return "break"
        if ctrl:
            if col_name in self.container_header_selected:
                self.container_header_selected.remove(col_name)
            else:
                self.container_header_selected.add(col_name)
            self.container_header_last_index = col_index
            return "break"
        self.container_header_selected = {col_name}
        self.container_header_last_index = col_index
        return None

    def on_container_header_drag(self, event):
        if self.container_drag_start_index is None or not self.container_drag_allowed:
            return
        region = self.container_tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        col_id = self.container_tree.identify_column(event.x)
        if not col_id:
            return
        target_index = int(col_id[1:]) - 1
        columns = self._get_display_columns(self.container_tree)
        if target_index < 0 or target_index >= len(columns):
            return
        if target_index == self.container_drag_start_index:
            return
        col_name = self.container_drag_col_name
        if not col_name or col_name not in columns:
            return
        columns.remove(col_name)
        columns.insert(target_index, col_name)
        self._set_display_columns(self.container_tree, columns)
        self.container_drag_start_index = target_index
        self.container_drag_active = True

    def on_container_header_release(self, event):
        if self.container_drag_active:
            self.container_drag_active = False
            self.container_drag_start_index = None
            self.container_drag_col_name = None
            return "break"
        self.container_drag_start_index = None
        self.container_drag_col_name = None
        return None

    def _open_column_manager_dialog(self, tree, headings):
        columns = list(tree["columns"])
        visible = self._get_display_columns(tree)
        hidden = [c for c in columns if c not in visible]
        order = list(visible) + hidden
        hidden_set = set(hidden)

        dialog = tk.Toplevel(self)
        dialog.title("列管理")
        dialog.geometry("360x420")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        list_frame = ttk.Frame(dialog, padding=10)
        list_frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(list_frame, selectmode="extended", exportselection=False)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=vsb.set)
        listbox.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def format_label(col):
            name = col
            if headings:
                name = headings.get(col, col)
            prefix = "[H] " if col in hidden_set else "    "
            return f"{prefix}{name}"

        def refresh_listbox():
            listbox.delete(0, "end")
            for col in order:
                listbox.insert("end", format_label(col))

        refresh_listbox()

        btn_frame = ttk.Frame(dialog, padding=10)
        btn_frame.pack(fill="x")

        def get_selected_indices():
            return list(listbox.curselection())

        def hide_selected():
            for idx in get_selected_indices():
                hidden_set.add(order[idx])
            refresh_listbox()

        def show_selected():
            for idx in get_selected_indices():
                hidden_set.discard(order[idx])
            refresh_listbox()

        def show_all():
            hidden_set.clear()
            refresh_listbox()

        def move_up():
            indices = get_selected_indices()
            if not indices:
                return
            current_indices = set(indices)
            # Process in ascending order
            for idx in sorted(indices):
                if idx == 0:
                    continue
                if (idx - 1) in current_indices:
                    continue
                order[idx - 1], order[idx] = order[idx], order[idx - 1]
                current_indices.remove(idx)
                current_indices.add(idx - 1)
            refresh_listbox()
            listbox.selection_clear(0, "end")
            for idx in current_indices:
                listbox.select_set(idx)

        def move_down():
            indices = get_selected_indices()
            if not indices:
                return
            current_indices = set(indices)
            # Process in descending order
            for idx in sorted(indices, reverse=True):
                if idx >= len(order) - 1:
                    continue
                if (idx + 1) in current_indices:
                    continue
                order[idx + 1], order[idx] = order[idx], order[idx + 1]
                current_indices.remove(idx)
                current_indices.add(idx + 1)
            refresh_listbox()
            listbox.selection_clear(0, "end")
            for idx in current_indices:
                listbox.select_set(idx)

        ttk.Button(btn_frame, text="隐藏选中", command=hide_selected).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="显示选中", command=show_selected).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="显示全部", command=show_all).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="上移", command=move_up).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="下移", command=move_down).pack(side="left", padx=4)

        action_frame = ttk.Frame(dialog, padding=10)
        action_frame.pack(fill="x")

        def apply_changes():
            visible_cols = [c for c in order if c not in hidden_set]
            if not visible_cols:
                messagebox.showinfo("提示", "至少保留一列显示。")
                return
            self._set_display_columns(tree, visible_cols)
            dialog.destroy()

        ttk.Button(action_frame, text="应用", command=apply_changes).pack(side="right", padx=4)
        ttk.Button(action_frame, text="取消", command=dialog.destroy).pack(side="right")

    def _get_product_decimal_cols(self):
        return [
            "carton_count", "pack_per_carton",
            "quantity", "unit_price", "unit_price_no_tax", "amount", "amount_no_tax",
            "gross_weight", "net_weight", "volume", "allocated_cost",
            "unit_allocated_cost", "unit_inventory_cost", "inventory_cost", "tax_rate"
        ]

    def _get_container_decimal_cols(self):
        return [
            "tax_refund", "sea_freight_usd", "all_in_rmb",
            "insurance_usd", "exchange_rate", "agency_fee_rmb",
            "misc_rmb", "misc_total_rmb", "tax_rate"
        ]

    def _get_product_row_entries(self, keyword, container_no):
        rows = self.db.query_products(keyword, container_no)

        def get_product_cell_value(row, col_name):
            # 动态获取税率，默认为 0.13
            # sqlite3.Row 不支持 .get()，直接通过键名访问
            try:
                tax_rate = row["tax_rate"]
            except (IndexError, KeyError):
                tax_rate = 0.13
            
            if tax_rate is None:
                tax_rate = 0.13
            divisor = 1 + tax_rate

            if col_name == "unit_price_no_tax":
                unit_price = row["unit_price"]
                if unit_price is None:
                    return None
                return unit_price / divisor
            if col_name == "amount_no_tax":
                amount = row["amount"]
                qty = row["quantity"]
                if amount is not None:
                    return amount / divisor
                if qty is None:
                    return None
                unit_price_no_tax = get_product_cell_value(row, "unit_price_no_tax")
                if unit_price_no_tax is None:
                    return None
                return unit_price_no_tax * qty
            if col_name == "unit_allocated_cost":
                alloc = row["allocated_cost"]
                qty = row["quantity"]
                if alloc is None or qty in (None, 0):
                    return None
                return alloc / qty
            if col_name == "unit_inventory_cost":
                unit_price_no_tax = get_product_cell_value(row, "unit_price_no_tax")
                unit_alloc = get_product_cell_value(row, "unit_allocated_cost")
                if unit_price_no_tax is None and unit_alloc is None:
                    return None
                return (unit_price_no_tax or 0.0) + (unit_alloc or 0.0)
            if col_name == "inventory_cost":
                amount_no_tax = get_product_cell_value(row, "amount_no_tax")
                alloc = row["allocated_cost"]
                if amount_no_tax is None and alloc is None:
                    return None
                return (amount_no_tax or 0.0) + (alloc or 0.0)
            return row[col_name]

        row_entries = []
        for r in rows:
            entry = {}
            for col_name in self.product_tree["columns"]:
                entry[col_name] = get_product_cell_value(r, col_name)
            row_entries.append(entry)
        return row_entries

    def _get_container_row_entries(self, container_no):
        rows = self.db.query_containers(container_no)
        row_entries = []
        for r in rows:
            entry = {}
            for col_name in self.container_tree["columns"]:
                entry[col_name] = r[col_name]
            row_entries.append(entry)
        return row_entries

    def _open_filter_dialog(self, table_type, col_name):
        if table_type == "product":
            title = self.product_column_reverse_map.get(col_name, col_name)
            decimal_cols = set(self._get_product_decimal_cols())
            numeric_cols = decimal_cols | {"id"}
            kw = self.prod_keyword_var.get().strip()
            ct = self.prod_container_var.get().strip()
            entries = self._get_product_row_entries(kw, ct)
            filters = self.product_filters
        else:
            headings = {
                "id": "ID",
                "shipment_code":   "指令号",
                "container_no":    "货柜号",
                "file_name":       "来源文件",
                "tax_refund":      "退税额",
                "sea_freight_usd": "海运费($)",
                "all_in_rmb":      "包干费(￥)",
                "insurance_usd":   "保费($)",
                "exchange_rate":   "汇率",
                "agency_fee_rmb":  "代理费(￥)",
                "misc_rmb":        "其他杂费(￥)",
                "misc_total_rmb":  "杂费汇总(￥)",
            }
            title = headings.get(col_name, col_name)
            decimal_cols = set(self._get_container_decimal_cols())
            if "exchange_rate" in decimal_cols:
                decimal_cols.remove("exchange_rate")
            numeric_cols = decimal_cols | {"id", "exchange_rate"}
            ct = self.cont_keyword_var.get().strip()
            entries = self._get_container_row_entries(ct)
            filters = self.container_filters

        other_filters = {k: v for k, v in filters.items() if k != col_name}
        entries = self._apply_filters(entries, other_filters, table_type)

        value_items = []
        seen = set()
        for entry in entries:
            display = self._display_value_for_filter(table_type, col_name, entry.get(col_name), decimal_cols)
            if display not in seen:
                seen.add(display)
                value_items.append(display)
        value_items.sort(key=lambda v: (v == "(空白)", v))

        dialog = tk.Toplevel(self)
        dialog.title(f"筛选 - {title}")
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        note = ttk.Notebook(dialog)
        note.pack(fill="both", expand=True, padx=10, pady=10)

        values_frame = ttk.Frame(note)
        text_frame = ttk.Frame(note)
        number_frame = ttk.Frame(note)
        note.add(values_frame, text="值筛选")
        note.add(text_frame, text="文本筛选")
        note.add(number_frame, text="数值筛选")

        list_frame = ttk.Frame(values_frame)
        list_frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(list_frame, selectmode="multiple")
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=vsb.set)
        listbox.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for item in value_items:
            listbox.insert("end", item)

        btn_row = ttk.Frame(values_frame)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="全选", command=lambda: listbox.select_set(0, "end")).pack(side="left", padx=4)
        ttk.Button(btn_row, text="清空", command=lambda: listbox.selection_clear(0, "end")).pack(side="left", padx=4)

        ttk.Label(text_frame, text="条件:").grid(row=0, column=0, sticky="e", padx=4, pady=6)
        text_op = tk.StringVar(value="包含")
        text_op_menu = ttk.Combobox(
            text_frame,
            textvariable=text_op,
            values=["包含", "不包含", "开头", "结尾", "正则"],
            width=10,
            state="readonly"
        )
        text_op_menu.grid(row=0, column=1, sticky="w", padx=4, pady=6)
        ttk.Label(text_frame, text="内容:").grid(row=1, column=0, sticky="e", padx=4, pady=6)
        text_value = tk.StringVar()
        ttk.Entry(text_frame, textvariable=text_value, width=25).grid(row=1, column=1, sticky="w", padx=4, pady=6)

        ttk.Label(number_frame, text="条件:").grid(row=0, column=0, sticky="e", padx=4, pady=6)
        num_value = tk.StringVar()
        ttk.Entry(number_frame, textvariable=num_value, width=25).grid(row=0, column=1, sticky="w", padx=4, pady=6)
        ttk.Label(number_frame, text="示例: >10, <=5, 10~20").grid(row=1, column=1, sticky="w", padx=4, pady=2)

        current = filters.get(col_name)
        if isinstance(current, dict):
            mode = current.get("mode")
            if mode == "values":
                selected = set(current.get("values", []))
                for i, item in enumerate(value_items):
                    if item in selected:
                        listbox.select_set(i)
                note.select(values_frame)
            elif mode == "text":
                text_op.set(current.get("op", "包含"))
                text_value.set(current.get("value", ""))
                note.select(text_frame)
            elif mode == "number":
                num_value.set(current.get("expr", ""))
                note.select(number_frame)
        if not (isinstance(current, dict) and current.get("mode") == "values"):
            if value_items:
                listbox.select_set(0, "end")

        if col_name not in numeric_cols:
            note.tab(number_frame, state="disabled")

        def apply_filter():
            current_tab = note.tab(note.select(), "text")
            new_filter = None

            if current_tab == "值筛选":
                selected = [value_items[i] for i in listbox.curselection()]
                if selected and len(selected) < len(value_items):
                    new_filter = {"mode": "values", "values": selected}
            elif current_tab == "文本筛选":
                val = text_value.get().strip()
                if val:
                    op_map = {
                        "包含": "contains",
                        "不包含": "not_contains",
                        "开头": "startswith",
                        "结尾": "endswith",
                        "正则": "regex"
                    }
                    op = op_map.get(text_op.get(), "contains")
                    if op == "regex":
                        try:
                            re.compile(val)
                        except re.error as ex:
                            messagebox.showerror("正则错误", f"正则表达式无效: {ex}")
                            return
                    new_filter = {"mode": "text", "op": op, "value": val}
            else:
                expr = num_value.get().strip()
                if expr:
                    if not self._parse_number_expr(expr):
                        messagebox.showerror("数值条件错误", "数值条件格式无效，请参考示例。")
                        return
                    new_filter = {"mode": "number", "expr": expr}

            if new_filter:
                filters[col_name] = new_filter
            else:
                filters.pop(col_name, None)

            self._save_filter_state()
            if table_type == "product":
                self.refresh_product_table()
            else:
                self.refresh_container_table()
            dialog.destroy()

        def clear_filter():
            filters.pop(col_name, None)
            self._save_filter_state()
            if table_type == "product":
                self.refresh_product_table()
            else:
                self.refresh_container_table()
            dialog.destroy()

        btns = ttk.Frame(dialog)
        btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(btns, text="应用", command=apply_filter).pack(side="right", padx=4)
        ttk.Button(btns, text="清除", command=clear_filter).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=dialog.destroy).pack(side="right", padx=4)

        dialog.geometry("360x380")

    def _display_value_for_filter(self, table_type, col_name, value, decimal_cols):
        if value is None or value == "":
            return "(空白)"
        if isinstance(value, (int, float)) and col_name in decimal_cols:
            return f"{value:.2f}"
        return str(value)

    def _match_text_filter(self, value, op, target):
        text = "" if value is None else str(value)
        text_lower = text.lower()
        target_lower = target.lower()
        if op == "contains":
            return target_lower in text_lower
        if op == "not_contains":
            return target_lower not in text_lower
        if op == "startswith":
            return text_lower.startswith(target_lower)
        if op == "endswith":
            return text_lower.endswith(target_lower)
        if op == "regex":
            try:
                return re.search(target, text) is not None
            except re.error:
                return False
        return False

    def _parse_number_expr(self, expr):
        expr = expr.strip()
        if not expr:
            return None
        if "~" in expr:
            parts = expr.split("~", 1)
            try:
                low = float(parts[0].strip())
                high = float(parts[1].strip())
                if low > high:
                    low, high = high, low
                return ("range", low, high)
            except ValueError:
                return None
        for op in [">=", "<=", ">", "<", "="]:
            if expr.startswith(op):
                try:
                    num = float(expr[len(op):].strip())
                    return (op, num)
                except ValueError:
                    return None
        try:
            num = float(expr)
            return ("=", num)
        except ValueError:
            return None

    def _match_number_filter(self, value, expr):
        parsed = self._parse_number_expr(expr)
        if not parsed:
            return False
        if value is None:
            return False
        try:
            num = float(value)
        except (TypeError, ValueError):
            return False
        if parsed[0] == "range":
            return parsed[1] <= num <= parsed[2]
        op, target = parsed
        if op == ">=":
            return num >= target
        if op == "<=":
            return num <= target
        if op == ">":
            return num > target
        if op == "<":
            return num < target
        return num == target

    def _filter_state_path(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, "shipping_filters.json")

    def _load_filter_state(self):
        path = self._filter_state_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.product_filters = data.get("product_filters", {}) or {}
            self.container_filters = data.get("container_filters", {}) or {}
            self.special_linkage_var.set(data.get("special_linkage_enabled", True))
        except Exception:
            self.product_filters = {}
            self.container_filters = {}

    def _save_filter_state(self):
        path = self._filter_state_path()
        data = {
            "product_filters": self.product_filters,
            "container_filters": self.container_filters,
            "special_linkage_enabled": self.special_linkage_var.get()
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
