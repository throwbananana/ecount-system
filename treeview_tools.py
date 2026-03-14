import tkinter as tk
from tkinter import ttk, messagebox
import re
from typing import Optional

from base_data_manager import BaseDataManager


_SMART_RESTORE_DEFAULT_MGR = None

_SMART_RESTORE_SPECS = {
    "account": {
        "label": "科目编码",
        "table": "account_subject",
        "code_cols": ["科目编码", "科目代码", "会计科目", "科目"],
        "name_cols": ["科目名称", "科目名", "科目"],
    },
    "partner": {
        "label": "往来单位编码",
        "table": "business_partner",
        "code_cols": ["往来单位编码", "客户编码", "供应商编码", "往来编码"],
        "name_cols": ["往来单位名", "往来单位名称", "客户名称", "供应商名称", "单位名称"],
    },
    "product": {
        "label": "品目编码",
        "table": "product",
        "code_cols": ["品目编码", "品目代码", "商品编码", "货品编码"],
        "name_cols": ["品目名", "品目名称", "商品名称", "货品名称", "品名"],
    },
}

_DEFAULT_TABLE_LABELS = {
    "currency": "币种",
    "department": "部门",
    "warehouse": "仓库",
    "account_subject": "科目编码",
    "product": "品目信息",
    "business_partner": "往来单位",
    "bank_account": "账户",
}


def _normalize_label(value):
    return re.sub(r"\s+", "", str(value or "")).lower()


def _get_tree_columns(tree: ttk.Treeview):
    return list(tree["columns"]) if tree["columns"] else []


def _get_heading_text(tree: ttk.Treeview, col_name: str):
    try:
        return tree.heading(col_name).get("text") or col_name
    except Exception:
        return col_name


def _find_column_index(tree: ttk.Treeview, candidates):
    if not candidates:
        return None
    candidates_norm = {_normalize_label(c) for c in candidates}
    cols = _get_tree_columns(tree)
    for idx, col in enumerate(cols):
        if _normalize_label(col) in candidates_norm:
            return idx
        if _normalize_label(_get_heading_text(tree, col)) in candidates_norm:
            return idx
    return None


def _safe_str(val):
    return "" if val is None else str(val).strip()


def _get_base_data_mgr_for_tree(tree: Optional[ttk.Treeview], override=None):
    if override is not None:
        return override, False
    if tree is not None:
        mgr = getattr(tree, "_base_data_mgr", None)
    else:
        mgr = None
    if mgr is not None:
        return mgr, False
    if tree is not None:
        try:
            top = tree.winfo_toplevel()
        except Exception:
            top = None
        if top is not None:
            mgr = getattr(top, "base_data_mgr", None)
            if mgr is not None:
                return mgr, False
    if _SMART_RESTORE_DEFAULT_MGR is not None:
        return _SMART_RESTORE_DEFAULT_MGR, False
    try:
        return BaseDataManager(), True
    except Exception:
        return None, False


def _resolve_col_index(tree: ttk.Treeview, col_key):
    if not col_key:
        return None
    cols = _get_tree_columns(tree)
    if col_key in cols:
        return cols.index(col_key)
    for idx, col in enumerate(cols):
        if _get_heading_text(tree, col) == col_key:
            return idx
    return None


def _restore_codes_in_tree(tree: ttk.Treeview, category_key: str, base_data_mgr=None,
                           code_col=None, name_col=None, label_override=None):
    delegate = getattr(tree, "_smart_restore_delegate", None)
    if callable(delegate):
        try:
            delegate(category_key, code_col, name_col)
        except TypeError:
            delegate(category_key)
        return
    spec = _SMART_RESTORE_SPECS.get(category_key) or {}
    table_name = spec.get("table", category_key)
    label = label_override or spec.get("label") or str(category_key)
    code_idx = _resolve_col_index(tree, code_col) if code_col else None
    name_idx = _resolve_col_index(tree, name_col) if name_col else None
    if code_idx is None and not code_col:
        code_idx = _find_column_index(tree, spec.get("code_cols", []))
    if name_idx is None and not name_col:
        name_idx = _find_column_index(tree, spec.get("name_cols", []))
    if code_idx is None:
        messagebox.showwarning("提示", f"未找到 {label} 列。")
        return
    mgr, created = _get_base_data_mgr_for_tree(tree, base_data_mgr)
    if mgr is None:
        messagebox.showwarning("提示", "基础数据管理器未初始化。")
        return
    updated = 0
    missing_actions = {}
    try:
        for iid in tree.get_children(""):
            values = list(tree.item(iid, "values"))
            if code_idx >= len(values):
                continue
            current_val = values[code_idx]
            match = mgr.find_best_match(table_name, _safe_str(current_val))
            if not match and name_idx is not None and name_idx < len(values):
                match = mgr.find_best_match(table_name, _safe_str(values[name_idx]))
            if match and str(match) != str(current_val):
                values[code_idx] = match
                tree.item(iid, values=values)
                updated += 1
                continue
            if match or not _safe_str(current_val):
                continue

            code_str = _safe_str(current_val)
            name_str = _safe_str(values[name_idx]) if name_idx is not None and name_idx < len(values) else ""
            if code_str in missing_actions:
                action = missing_actions[code_str]
            else:
                action = _prompt_unmatched_action(tree.winfo_toplevel(), label, code_str, name_str)
                missing_actions[code_str] = action

            if action == "fuzzy":
                match = mgr.find_best_match(table_name, code_str, min_score=0.6)
                if not match and name_str:
                    match = mgr.find_best_match(table_name, name_str, min_score=0.6)
                if match and str(match) != str(current_val):
                    values[code_idx] = match
                    tree.item(iid, values=values)
                    updated += 1
                else:
                    follow = _prompt_unmatched_action(tree.winfo_toplevel(), label, code_str, name_str)
                    if follow == "add":
                        payload = {"code": code_str, "name": name_str or code_str}
                        try:
                            mgr.add_record(table_name, payload)
                            if hasattr(mgr, "clear_lookup_cache"):
                                mgr.clear_lookup_cache()
                        except Exception:
                            messagebox.showerror("错误", f"加入基础数据失败: {code_str}")
            elif action == "add":
                payload = {"code": code_str, "name": name_str or code_str}
                try:
                    mgr.add_record(table_name, payload)
                    if hasattr(mgr, "clear_lookup_cache"):
                        mgr.clear_lookup_cache()
                except Exception:
                    messagebox.showerror("错误", f"加入基础数据失败: {code_str}")
        if hasattr(mgr, "clear_lookup_cache"):
            mgr.clear_lookup_cache()
    finally:
        if created:
            try:
                mgr.close()
            except Exception:
                pass
    if hasattr(tree, "_treeview_tools"):
        try:
            tree._treeview_tools.reset_items()
        except Exception:
            pass
    messagebox.showinfo("完成", f"{label} 智能还原完成，更新 {updated} 行。")


def _prompt_restore_columns(tree: ttk.Treeview, category_key: str, base_data_mgr=None, label_override=None):
    cols = _get_tree_columns(tree)
    if not cols:
        messagebox.showwarning("提示", "当前表格没有可用列。")
        return
    spec = _SMART_RESTORE_SPECS.get(category_key, {})
    title = label_override or spec.get("label", str(category_key))

    display_map = []
    for col in cols:
        display = _get_heading_text(tree, col)
        display_map.append((display, col))

    dialog = tk.Toplevel(tree.winfo_toplevel())
    dialog.title(f"{title} - 选择列")
    dialog.transient(tree.winfo_toplevel())
    dialog.grab_set()

    ttk.Label(dialog, text="编码列:").grid(row=0, column=0, padx=10, pady=8, sticky="e")
    code_var = tk.StringVar()
    code_combo = ttk.Combobox(dialog, textvariable=code_var, state="readonly", width=28)
    code_combo["values"] = [d for d, _ in display_map]
    code_combo.grid(row=0, column=1, padx=6, pady=8, sticky="w")
    code_combo.current(0)

    ttk.Label(dialog, text="名称列(可选):").grid(row=1, column=0, padx=10, pady=8, sticky="e")
    name_var = tk.StringVar(value="(不使用)")
    name_combo = ttk.Combobox(dialog, textvariable=name_var, state="readonly", width=28)
    name_combo["values"] = ["(不使用)"] + [d for d, _ in display_map]
    name_combo.grid(row=1, column=1, padx=6, pady=8, sticky="w")

    def _confirm():
        code_display = code_var.get().strip()
        name_display = name_var.get().strip()
        code_col = None
        name_col = None
        for d, key in display_map:
            if d == code_display:
                code_col = key
                break
        if name_display and name_display != "(不使用)":
            for d, key in display_map:
                if d == name_display:
                    name_col = key
                    break
        if not code_col:
            messagebox.showwarning("提示", "请选择编码列。")
            return
        dialog.destroy()
        _restore_codes_in_tree(
            tree,
            category_key,
            base_data_mgr,
            code_col=code_col,
            name_col=name_col,
            label_override=label_override
        )

    btns = ttk.Frame(dialog)
    btns.grid(row=2, column=0, columnspan=2, pady=10)
    ttk.Button(btns, text="开始匹配", command=_confirm).pack(side="right", padx=6)
    ttk.Button(btns, text="取消", command=dialog.destroy).pack(side="right")


def _prompt_unmatched_action(parent, label, code_value, name_value):
    result = {"choice": "skip"}

    dialog = tk.Toplevel(parent)
    dialog.title("未匹配提示")
    dialog.transient(parent)
    dialog.grab_set()

    msg = f"{label} 未在基础数据中找到：\n\n编码: {code_value}"
    if name_value:
        msg += f"\n名称: {name_value}"
    ttk.Label(dialog, text=msg, justify="left").pack(padx=12, pady=(12, 6))
    ttk.Label(dialog, text="请选择处理方式：", foreground="gray").pack(padx=12, pady=(0, 6), anchor="w")

    btns = ttk.Frame(dialog)
    btns.pack(padx=12, pady=12, fill="x")

    def _choose(val):
        result["choice"] = val
        dialog.destroy()

    ttk.Button(btns, text="不加入", command=lambda: _choose("skip"), width=10).pack(side="right", padx=6)
    ttk.Button(btns, text="继续模糊匹配", command=lambda: _choose("fuzzy"), width=14).pack(side="right", padx=6)
    ttk.Button(btns, text="加入基础数据", command=lambda: _choose("add"), width=12).pack(side="right")

    parent.wait_window(dialog)
    return result["choice"]


def _build_restore_categories(base_data_mgr=None):
    mgr, created = _get_base_data_mgr_for_tree(None, base_data_mgr)
    try:
        if mgr is None:
            return [("科目编码", "account_subject"), ("往来单位", "business_partner"), ("品目信息", "product")]
        items = []
        for table in mgr.DATA_FILES.values():
            label = _DEFAULT_TABLE_LABELS.get(table, table)
            items.append((label, table))
        try:
            for cat in mgr.list_custom_categories():
                name_key = cat.get("name")
                display = cat.get("display_name") or name_key
                if name_key:
                    items.append((display, f"custom:{name_key}"))
        except Exception:
            pass
        return items
    finally:
        if created:
            try:
                mgr.close()
            except Exception:
                pass


def _add_smart_restore_menu(menu: tk.Menu, tree: ttk.Treeview, base_data_mgr=None):
    restore_menu = tk.Menu(menu, tearoff=0)
    for label, table_name in _build_restore_categories(base_data_mgr):
        restore_menu.add_command(
            label=label,
            command=lambda t=table_name, l=label: _prompt_restore_columns(tree, t, base_data_mgr, label_override=l)
        )
    menu.add_cascade(label="智能匹配", menu=restore_menu)


def add_smart_restore_menu(menu: tk.Menu, tree: ttk.Treeview, base_data_mgr=None):
    _add_smart_restore_menu(menu, tree, base_data_mgr)


def install_smart_restore_header(root: tk.Misc, base_data_mgr=None):
    global _SMART_RESTORE_DEFAULT_MGR
    _SMART_RESTORE_DEFAULT_MGR = base_data_mgr

    def _on_treeview_header_right_click(event):
        tree = event.widget
        if hasattr(tree, "_treeview_tools") or getattr(tree, "_skip_smart_restore_header_menu", False):
            return
        region = tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        menu = tk.Menu(tree, tearoff=0)
        _add_smart_restore_menu(menu, tree, base_data_mgr)
        menu.tk_popup(event.x_root, event.y_root)

    root.bind_class("Treeview", "<Button-3>", _on_treeview_header_right_click, add="+")


class TreeviewTools:
    def __init__(self, tree: ttk.Treeview, headings=None, allow_reorder=True, base_data_mgr=None):
        self.tree = tree
        self.headings = headings or {}
        self.allow_reorder = allow_reorder
        self.base_data_mgr = base_data_mgr
        self.filters = {}
        self.hidden_columns = set()
        self.all_items = []
        self.selected_header_cols = set()
        self.last_header_index = None
        self.drag_start_index = None
        self.drag_active = False
        self.drag_col_name = None
        self.drag_allowed = True
        self.last_clicked_cell = (None, None) # (item_id, col_id)
        self.drag_selecting = False
        self.drag_start_item = None
        self.drag_start_col = None
        self.cell_select_range = None  # (row_start, row_end, col_start, col_end) in display indices
        
        # 使用四条边线模拟 Excel 的选区边框，避免遮挡内容
        self.focus_lines = {
            "top": tk.Frame(self.tree, bg="#0078d7", height=2, bd=0),
            "bottom": tk.Frame(self.tree, bg="#0078d7", height=2, bd=0),
            "left": tk.Frame(self.tree, bg="#0078d7", width=2, bd=0),
            "right": tk.Frame(self.tree, bg="#0078d7", width=2, bd=0),
        }
        for line in self.focus_lines.values():
            line.lower()
            line.bind("<Button-1>", self._on_box_click)
            line.bind("<Double-1>", self._on_box_double_click)
            line.bind("<Button-3>", self._on_box_right_click)
        
        self.tree.bind("<Button-3>", self._on_right_click, add="+")
        self.tree.bind("<ButtonPress-1>", self._on_left_press, add="+")
        self.tree.bind("<B1-Motion>", self._on_left_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_left_release, add="+")
        self.tree.bind("<Control-c>", self._on_copy)
        self.tree.bind("<Control-v>", self._on_paste)
        self.tree.bind("<Control-x>", self._on_cut)
        self.tree.bind("<Delete>", self._on_clear)

    def _on_box_click(self, event):
        # 转发单击
        x = event.widget.winfo_x() + event.x
        y = event.widget.winfo_y() + event.y
        self.tree.event_generate("<Button-1>", x=x, y=y, state=event.state)
        return "break"

    def _on_box_double_click(self, event):
        # 转发双击
        x = event.widget.winfo_x() + event.x
        y = event.widget.winfo_y() + event.y
        self.tree.event_generate("<Double-1>", x=x, y=y, state=event.state)
        return "break"

    def _on_box_right_click(self, event):
        # 转发右键
        x = event.widget.winfo_x() + event.x
        y = event.widget.winfo_y() + event.y
        self.tree.event_generate("<Button-3>", x=x, y=y, state=event.state)
        return "break"

    def reset_items(self):
        self._capture_items()

    def _capture_items(self):
        self.all_items = []
        children = self.tree.get_children("")
        for idx, iid in enumerate(children):
            self.all_items.append({
                "iid": iid,
                "index": idx,
                "values": self.tree.item(iid, "values")
            })

    def _ensure_items(self):
        if not self.all_items:
            self._capture_items()
            return
        missing = [item for item in self.all_items if not self.tree.exists(item["iid"])]
        if missing:
            if len(missing) >= max(1, len(self.all_items) // 2):
                self._capture_items()
            else:
                self.all_items = [item for item in self.all_items if self.tree.exists(item["iid"])]

    def _on_right_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            # ... (保持原有的列管理菜单代码)
            col_id = self.tree.identify_column(event.x)
            if not col_id: return
            col_index = int(col_id[1:]) - 1
            columns = self._get_displaycolumns()
            if col_index < 0 or col_index >= len(columns): return
            col_name = columns[col_index]

            menu = tk.Menu(self.tree, tearoff=0)
            menu.add_command(label="筛选...", command=lambda: self._open_filter_dialog(col_name))
            menu.add_command(label="清除此列筛选", command=lambda: self._clear_filter(col_name))
            menu.add_command(label="清除所有筛选", command=self._clear_all_filters)
            menu.add_separator()
            if self.selected_header_cols:
                menu.add_command(label=f"隐藏选中列 ({len(self.selected_header_cols)})",
                                 command=lambda: self._hide_columns(self.selected_header_cols))
            menu.add_command(label="隐藏此列", command=lambda: self._hide_column(col_name))
            menu.add_command(label="显示全部列", command=self._show_all_columns)
            menu.add_command(label="列管理...", command=self._open_column_manager)
            menu.add_separator()
            _add_smart_restore_menu(menu, self.tree, self.base_data_mgr)
            menu.tk_popup(event.x_root, event.y_root)
        elif region in ("cell", "tree", "item"):
            # 数据行右键菜单
            item_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            if item_id:
                if item_id not in self.tree.selection():
                    self.tree.selection_set(item_id)
                self.last_clicked_cell = (item_id, col_id)
                
                # 同时也移动焦点框
                bbox = self.tree.bbox(item_id, col_id)
                if bbox:
                    self._place_focus_lines(bbox)

                menu = tk.Menu(self.tree, tearoff=0)
                menu.add_command(label="复制选中行", command=self._on_copy, accelerator="Ctrl+C")
                menu.add_command(label="粘贴到此处", command=self._on_paste, accelerator="Ctrl+V")
                menu.add_separator()
                menu.add_command(label="刷新表格", command=lambda: self.tree.event_generate("<<TreeviewRefresh>>"))
                menu.tk_popup(event.x_root, event.y_root)

    def _get_displaycolumns(self):
        display = self.tree["displaycolumns"]
        if display == "#all":
            return list(self.tree["columns"])
        if isinstance(display, (list, tuple)):
            if "#all" in display:
                return list(self.tree["columns"])
            return list(display)
        return list(self.tree["columns"])

    def _set_displaycolumns(self, cols):
        if list(cols) == list(self.tree["columns"]):
            self.tree["displaycolumns"] = "#all"
        else:
            self.tree["displaycolumns"] = cols

    def get_visual_data(self):
        """返回当前视觉显示顺序的表头和行数据"""
        all_cols = list(self.tree["columns"])
        display_cols = self._get_displaycolumns()
        if not display_cols:
            display_cols = all_cols
        
        # 获取表头文本
        headers = []
        for col in display_cols:
            header_text = self.tree.heading(col).get("text") or col
            headers.append(header_text)
            
        # 获取行数据
        rows = []
        col_indices = [all_cols.index(c) for c in display_cols]
        for iid in self.tree.get_children(""):
            all_values = self.tree.item(iid, "values")
            # 确保 values 长度足够
            vals_list = list(all_values)
            while len(vals_list) < len(all_cols):
                vals_list.append("")
            
            # 按视觉顺序提取
            row_vals = [vals_list[idx] for idx in col_indices]
            rows.append(row_vals)
            
        return headers, rows

    def _hide_column(self, col_name):
        display = self._get_displaycolumns()
        if col_name not in display:
            return
        if len(display) <= 1:
            messagebox.showinfo("提示", "至少保留一列显示。")
            return
        display = [c for c in display if c != col_name]
        self.hidden_columns.add(col_name)
        self._set_displaycolumns(display)

    def _hide_columns(self, col_names):
        display = self._get_displaycolumns()
        remain = [c for c in display if c not in col_names]
        if not remain:
            messagebox.showinfo("提示", "至少保留一列显示。")
            return
        self.hidden_columns.update(col_names)
        self._set_displaycolumns(remain)

    def _show_all_columns(self):
        self.hidden_columns.clear()
        self._set_displaycolumns(list(self.tree["columns"]))

    def _open_filter_dialog(self, col_name):
        self._ensure_items()
        title = self.headings.get(col_name) or self.tree.heading(col_name).get("text") or col_name

        other_filters = {k: v for k, v in self.filters.items() if k != col_name}
        value_items = []
        seen = set()
        for item in self.all_items:
            if not self.tree.exists(item["iid"]):
                continue
            values = self.tree.item(item["iid"], "values")
            item["values"] = values
            if not self._match_filters(values, other_filters):
                continue
            display = self._display_value(self._get_value_by_col(values, col_name))
            if display not in seen:
                seen.add(display)
                value_items.append(display)
        value_items.sort(key=lambda v: (v == "(空白)", v))

        dialog = tk.Toplevel(self.tree.winfo_toplevel())
        dialog.title(f"筛选 - {title}")
        dialog.transient(self.tree.winfo_toplevel())
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

        current = self.filters.get(col_name)
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
                self.filters[col_name] = new_filter
            else:
                self.filters.pop(col_name, None)
            self._apply_filters_to_tree()
            dialog.destroy()

        def clear_filter():
            self.filters.pop(col_name, None)
            self._apply_filters_to_tree()
            dialog.destroy()

        btns = ttk.Frame(dialog)
        btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(btns, text="应用", command=apply_filter).pack(side="right", padx=4)
        ttk.Button(btns, text="清除", command=clear_filter).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=dialog.destroy).pack(side="right", padx=4)

        dialog.geometry("360x380")

    def _clear_filter(self, col_name):
        if col_name in self.filters:
            self.filters.pop(col_name, None)
            self._apply_filters_to_tree()

    def _clear_all_filters(self):
        if self.filters:
            self.filters = {}
            self._apply_filters_to_tree()

    def _apply_filters_to_tree(self):
        self._ensure_items()
        for item in self.all_items:
            if self.tree.exists(item["iid"]):
                self.tree.detach(item["iid"])
        for item in self.all_items:
            if not self.tree.exists(item["iid"]):
                continue
            values = self.tree.item(item["iid"], "values")
            item["values"] = values
            if self._match_filters(values, self.filters):
                self.tree.reattach(item["iid"], "", "end")

    def _match_filters(self, values, filters):
        columns = list(self.tree["columns"])
        for col_name, flt in filters.items():
            val = self._get_value_by_col(values, col_name)
            if isinstance(flt, dict):
                mode = flt.get("mode")
                if mode == "values":
                    selected = set(flt.get("values", []))
                    display = self._display_value(val)
                    if display not in selected:
                        return False
                elif mode == "text":
                    if not self._match_text(val, flt.get("op", "contains"), flt.get("value", "")):
                        return False
                elif mode == "number":
                    if not self._match_number(val, flt.get("expr", "")):
                        return False
            else:
                if not self._match_text(val, "contains", str(flt)):
                    return False
        return True

    def _get_value_by_col(self, values, col_name):
        columns = list(self.tree["columns"])
        if col_name not in columns:
            return None
        idx = columns.index(col_name)
        if idx >= len(values):
            return None
        return values[idx]

    def _display_value(self, value):
        if value is None or value == "":
            return "(空白)"
        return str(value)

    def _match_text(self, value, op, target):
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

    def _match_number(self, value, expr):
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

    def _on_left_press(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region in ("cell", "tree", "item"):
            item_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            if item_id and col_id:
                self.last_clicked_cell = (item_id, col_id)
                self.drag_selecting = True
                self.drag_start_item = item_id
                self.drag_start_col = col_id
                
                # 如果没按住 Ctrl/Shift，点击新单元格时先清空之前的选择 (除非已经是选中状态)
                ctrl = bool(event.state & 0x0004)
                shift = bool(event.state & 0x0001)
                if not (ctrl or shift):
                    # 只有在点击未选中的行时才清除，方便后续可能的拖动
                    # 但为了手感，通常点击即选中
                    pass

                # 更新视觉焦点框
                try:
                    self._update_cell_selection(item_id, col_id, item_id, col_id, keep_start=shift)
                except Exception:
                    self._hide_focus_lines()
        else:
            self.drag_selecting = False
            self.drag_start_item = None
            self.drag_start_col = None
            self.cell_select_range = None
            self._hide_focus_lines()

        if not self.allow_reorder:
            return None
        if region != "heading":
            self.drag_start_index = None
            self.drag_active = False
            self.drag_col_name = None
            return

        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        columns = self._get_displaycolumns()
        if col_index < 0 or col_index >= len(columns):
            return

        col_name = columns[col_index]
        ctrl = bool(event.state & 0x0004)
        shift = bool(event.state & 0x0001)
        self.drag_allowed = not (ctrl or shift)
        self.drag_start_index = col_index
        self.drag_active = False
        self.drag_col_name = col_name

        if shift and self.last_header_index is not None:
            start = min(self.last_header_index, col_index)
            end = max(self.last_header_index, col_index)
            self.selected_header_cols.update(columns[start:end + 1])
            return "break"
        if ctrl:
            if col_name in self.selected_header_cols:
                self.selected_header_cols.remove(col_name)
            else:
                self.selected_header_cols.add(col_name)
            self.last_header_index = col_index
            return "break"

        self.selected_header_cols = {col_name}
        self.last_header_index = col_index
        return None

    def _on_copy(self, event=None):
        lines = []
        all_cols = list(self.tree["columns"])
        display_cols = self._get_displaycolumns()
        children = list(self.tree.get_children(""))

        if self.cell_select_range:
            row_start, row_end, col_start, col_end = self.cell_select_range
            row_start = max(0, row_start)
            col_start = max(0, col_start)
            row_end = min(len(children) - 1, row_end)
            col_end = min(len(display_cols) - 1, col_end)
            for r_idx in range(row_start, row_end + 1):
                iid = children[r_idx]
                values = self.tree.item(iid, "values")
                row_data = []
                for c_idx in range(col_start, col_end + 1):
                    col_name = display_cols[c_idx]
                    try:
                        real_idx = all_cols.index(col_name)
                        if real_idx < len(values):
                            row_data.append(str(values[real_idx]))
                        else:
                            row_data.append("")
                    except ValueError:
                        row_data.append("")
                lines.append("\t".join(row_data))
        else:
            selection = self.tree.selection()
            if not selection:
                return
            for iid in selection:
                values = self.tree.item(iid, "values")
                if not values:
                    continue
                row_data = []
                for col in display_cols:
                    try:
                        idx = all_cols.index(col)
                        if idx < len(values):
                            row_data.append(str(values[idx]))
                        else:
                            row_data.append("")
                    except ValueError:
                        row_data.append("")
                lines.append("\t".join(row_data))

        if not lines:
            return

        text = "\n".join(lines)
        self.tree.clipboard_clear()
        self.tree.clipboard_append(text)
        return "break"

    def _on_paste(self, event=None):
        try:
            text = self.tree.clipboard_get()
        except tk.TclError:
            return
        
        if not text:
            return
            
        # Parse TSV
        rows_data = [line.split("\t") for line in text.splitlines()]
        if not rows_data:
            return
            
        selection = self.tree.selection()
        all_items = list(self.tree.get_children(""))
        if not all_items:
            return

        display_cols = self._get_displaycolumns()
        all_cols = list(self.tree["columns"])
        start_row_idx = 0
        start_col_display_idx = 0

        if self.cell_select_range:
            start_row_idx = self.cell_select_range[0]
            start_col_display_idx = self.cell_select_range[2]
        else:
            start_item, start_col_id = self.last_clicked_cell
            if start_item in all_items:
                start_row_idx = all_items.index(start_item)
            elif selection:
                start_row_idx = all_items.index(selection[0]) if selection[0] in all_items else 0

            if start_col_id:
                try:
                    if start_col_id.startswith("#"):
                        start_col_display_idx = int(start_col_id[1:]) - 1
                    elif start_col_id in display_cols:
                        start_col_display_idx = display_cols.index(start_col_id)
                except (ValueError, IndexError):
                    start_col_display_idx = 0

        updated_count = 0
        for r_idx, row_vals in enumerate(rows_data):
            target_row_idx = start_row_idx + r_idx
            if target_row_idx >= len(all_items):
                break
            iid = all_items[target_row_idx]
            current_values = list(self.tree.item(iid, "values"))
            # Ensure we have enough value slots
            while len(current_values) < len(all_cols):
                current_values.append("")
                
            row_changed = False
            for c_idx, val in enumerate(row_vals):
                target_display_idx = start_col_display_idx + c_idx
                if target_display_idx >= len(display_cols):
                    break
                
                col_name = display_cols[target_display_idx]
                try:
                    real_idx = all_cols.index(col_name)
                    if real_idx < len(current_values):
                        if str(current_values[real_idx]) != str(val):
                            current_values[real_idx] = val
                            row_changed = True
                except ValueError:
                    continue
            
            if row_changed:
                self.tree.item(iid, values=current_values)
                updated_count += 1
        
        if updated_count > 0:
            self.tree.event_generate("<<TreeviewPaste>>", when="tail")
        return "break"

    def _on_cut(self, event=None):
        if self._on_copy() == "break":
            if self._clear_selected_cells():
                self.tree.event_generate("<<TreeviewPaste>>", when="tail")
        return "break"

    def _on_clear(self, event=None):
        if self._clear_selected_cells():
            self.tree.event_generate("<<TreeviewPaste>>", when="tail")
        return "break"

    def _on_left_drag(self, event):
        region = self.tree.identify_region(event.x, event.y)
        
        # 处理数据区域的“框选”（行范围选择）
        if self.drag_selecting and self.drag_start_item:
            cur_item = self.tree.identify_row(event.y)
            cur_col = self.tree.identify_column(event.x)
            if cur_item and cur_col:
                self._update_cell_selection(self.drag_start_item, self.drag_start_col, cur_item, cur_col)
            return

        if not self.allow_reorder:
            return
        if self.drag_start_index is None or not self.drag_allowed:
            return
        if region != "heading":
            return
        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return
        target_index = int(col_id[1:]) - 1
        columns = self._get_displaycolumns()
        if target_index < 0 or target_index >= len(columns):
            return
        if target_index == self.drag_start_index:
            return
        col_name = self.drag_col_name
        if not col_name or col_name not in columns:
            return
        columns.remove(col_name)
        columns.insert(target_index, col_name)
        self._set_displaycolumns(columns)
        self.drag_start_index = target_index
        self.drag_active = True

    def _on_left_release(self, event):
        self.drag_selecting = False
        self.drag_start_item = None
        self.drag_start_col = None
        
        if not self.allow_reorder:
            return None
        if self.drag_active:
            self.drag_active = False
            self.drag_start_index = None
            self.drag_col_name = None
            return "break"
        self.drag_start_index = None
        self.drag_col_name = None
        return None

    def _update_cell_selection(self, start_item, start_col_id, end_item, end_col_id, keep_start=False):
        if not start_item or not start_col_id or not end_item or not end_col_id:
            return
        display_cols = self._get_displaycolumns()
        children = list(self.tree.get_children(""))
        if not children or not display_cols:
            return

        try:
            if keep_start and self.cell_select_range:
                row_start = self.cell_select_range[0]
                col_start = self.cell_select_range[2]
            else:
                row_start = children.index(start_item)
                if start_col_id.startswith("#"):
                    col_start = int(start_col_id[1:]) - 1
                else:
                    col_start = display_cols.index(start_col_id)
            row_end = children.index(end_item)
            if end_col_id.startswith("#"):
                col_end = int(end_col_id[1:]) - 1
            else:
                col_end = display_cols.index(end_col_id)
        except (ValueError, IndexError):
            return
        if col_start < 0 or col_end < 0:
            return

        row_min = min(row_start, row_end)
        row_max = max(row_start, row_end)
        col_min = min(col_start, col_end)
        col_max = max(col_start, col_end)
        self.cell_select_range = (row_min, row_max, col_min, col_max)

        try:
            if row_min == row_max and col_min == col_max:
                self.tree.selection_set(children[row_min])
            else:
                self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass

        try:
            top_left_iid = children[row_min]
            bottom_right_iid = children[row_max]
            tl_col_id = f"#{col_min + 1}"
            br_col_id = f"#{col_max + 1}"
            bbox1 = self.tree.bbox(top_left_iid, tl_col_id)
            bbox2 = self.tree.bbox(bottom_right_iid, br_col_id)
            if bbox1 and bbox2:
                x1, y1, w1, h1 = bbox1
                x2, y2, w2, h2 = bbox2
                x = x1
                y = y1
                width = (x2 + w2) - x1
                height = (y2 + h2) - y1
                self._place_focus_lines((x, y, width, height))
            else:
                self._hide_focus_lines()
        except Exception:
            self._hide_focus_lines()

    def _hide_focus_lines(self):
        for line in self.focus_lines.values():
            line.place_forget()

    def _place_focus_lines(self, bbox):
        x, y, width, height = bbox
        if width <= 0 or height <= 0:
            self._hide_focus_lines()
            return
        thickness = 2
        self.focus_lines["top"].place(x=x, y=y, width=width, height=thickness)
        self.focus_lines["bottom"].place(x=x, y=y + height - thickness, width=width, height=thickness)
        self.focus_lines["left"].place(x=x, y=y, width=thickness, height=height)
        self.focus_lines["right"].place(x=x + width - thickness, y=y, width=thickness, height=height)
        for line in self.focus_lines.values():
            line.lift()

    def _clear_selected_cells(self):
        all_cols = list(self.tree["columns"])
        display_cols = self._get_displaycolumns()
        children = list(self.tree.get_children(""))
        if not children or not display_cols:
            return False

        updated_count = 0
        if self.cell_select_range:
            row_start, row_end, col_start, col_end = self.cell_select_range
            row_start = max(0, row_start)
            col_start = max(0, col_start)
            row_end = min(len(children) - 1, row_end)
            col_end = min(len(display_cols) - 1, col_end)
            for r_idx in range(row_start, row_end + 1):
                iid = children[r_idx]
                values = list(self.tree.item(iid, "values"))
                while len(values) < len(all_cols):
                    values.append("")
                row_changed = False
                for c_idx in range(col_start, col_end + 1):
                    col_name = display_cols[c_idx]
                    try:
                        real_idx = all_cols.index(col_name)
                    except ValueError:
                        continue
                    if real_idx < len(values) and values[real_idx] != "":
                        values[real_idx] = ""
                        row_changed = True
                if row_changed:
                    self.tree.item(iid, values=values)
                    updated_count += 1
        else:
            selection = self.tree.selection()
            if not selection:
                return False
            for iid in selection:
                values = list(self.tree.item(iid, "values"))
                while len(values) < len(all_cols):
                    values.append("")
                row_changed = False
                for col in display_cols:
                    try:
                        real_idx = all_cols.index(col)
                    except ValueError:
                        continue
                    if real_idx < len(values) and values[real_idx] != "":
                        values[real_idx] = ""
                        row_changed = True
                if row_changed:
                    self.tree.item(iid, values=values)
                    updated_count += 1
        return updated_count > 0

    def _open_column_manager(self):
        columns = list(self.tree["columns"])
        visible = self._get_displaycolumns()
        hidden = [c for c in columns if c not in visible]
        order = list(visible) + hidden
        hidden_set = set(hidden)

        dialog = tk.Toplevel(self.tree.winfo_toplevel())
        dialog.title("列管理")
        dialog.geometry("360x420")
        dialog.transient(self.tree.winfo_toplevel())
        dialog.grab_set()

        list_frame = ttk.Frame(dialog, padding=10)
        list_frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(list_frame, selectmode="extended")
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=vsb.set)
        listbox.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        def format_label(col):
            name = self.headings.get(col) or self.tree.heading(col).get("text") or col
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
            for idx in indices:
                if idx == 0:
                    continue
                order[idx - 1], order[idx] = order[idx], order[idx - 1]
            refresh_listbox()
            for idx in [i - 1 if i > 0 else i for i in indices]:
                listbox.select_set(idx)

        def move_down():
            indices = get_selected_indices()
            if not indices:
                return
            for idx in reversed(indices):
                if idx >= len(order) - 1:
                    continue
                order[idx + 1], order[idx] = order[idx], order[idx + 1]
            refresh_listbox()
            for idx in [i + 1 if i < len(order) - 1 else i for i in indices]:
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
            self.hidden_columns = set(hidden_set)
            self._set_displaycolumns(visible_cols)
            dialog.destroy()

        ttk.Button(action_frame, text="应用", command=apply_changes).pack(side="right", padx=4)
        ttk.Button(action_frame, text="取消", command=dialog.destroy).pack(side="right")


def attach_treeview_tools(tree: ttk.Treeview, headings=None, allow_reorder=True, base_data_mgr=None):
    # 强制开启多选模式，支持框选
    tree.configure(selectmode="extended")
    tools = TreeviewTools(tree, headings=headings, allow_reorder=allow_reorder, base_data_mgr=base_data_mgr)
    tree._treeview_tools = tools
    return tools
