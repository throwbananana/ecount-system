# -*- coding: utf-8 -*-
"""
Excel 批量合并工具模块
用于整合多个 Excel 文件或 Sheet。
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import numpy as np
from treeview_tools import attach_treeview_tools

class ExcelMergerGUI:
    def __init__(self, parent_frame, open_in_converter=None):
        self.parent = parent_frame
        self.file_list = []
        self.open_in_converter = open_in_converter
        
        self._build_ui()

    def _build_ui(self):
        # 布局分为：顶部操作区（添加/删除文件），中部列表区，底部选项和执行区
        
        # --- 顶部操作区 ---
        top_frame = ttk.Frame(self.parent, padding=10)
        top_frame.pack(fill="x")
        
        ttk.Button(top_frame, text="添加文件...", command=self._add_files).pack(side="left", padx=5)
        ttk.Button(top_frame, text="清空列表", command=self._clear_files).pack(side="left", padx=5)
        ttk.Button(top_frame, text="移除选中", command=self._remove_selected).pack(side="left", padx=5)
        
        # --- 中部列表区 ---
        list_frame = ttk.Frame(self.parent, padding=10)
        list_frame.pack(fill="both", expand=True)
        
        columns = ("path", "size", "mod_time")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("path", text="文件路径")
        self.tree.heading("size", text="大小 (KB)")
        self.tree.heading("mod_time", text="修改时间")
        
        self.tree.column("path", width=400)
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("mod_time", width=150)

        self.tree_tools = attach_treeview_tools(self.tree)
        
        ysb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # --- 底部选项区 ---
        opt_frame = ttk.LabelFrame(self.parent, text="合并选项", padding=10)
        opt_frame.pack(fill="x", padx=10, pady=5)
        
        # 合并模式
        self.merge_mode = tk.StringVar(value="consolidate")
        ttk.Radiobutton(opt_frame, text="合并到一个工作表 (纵向追加)", variable=self.merge_mode, value="consolidate").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(opt_frame, text="合并到一个工作簿 (分Sheet存放)", variable=self.merge_mode, value="workbook").grid(row=0, column=1, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(opt_frame, text="智能自动分组 (结构相同合并，不同分表)", variable=self.merge_mode, value="auto_group").grid(row=0, column=2, sticky="w", padx=10, pady=5)
        
        # 详细选项 (Consolidate)
        self.add_source_col = tk.BooleanVar(value=True)
        self.align_columns = tk.BooleanVar(value=True)
        self.smart_header = tk.BooleanVar(value=True) # Default True
        
        cb_source = ttk.Checkbutton(opt_frame, text="添加来源文件名列", variable=self.add_source_col)
        cb_source.grid(row=1, column=0, sticky="w", padx=25)
        
        cb_align = ttk.Checkbutton(opt_frame, text="按列名自动对齐 (支持不同格式)", variable=self.align_columns)
        cb_align.grid(row=1, column=1, sticky="w", padx=25)
        
        cb_smart = ttk.Checkbutton(opt_frame, text="智能去重表头 (检查首行内容)", variable=self.smart_header)
        cb_smart.grid(row=2, column=0, columnspan=2, sticky="w", padx=25)
        
        # 联动控制
        def on_mode_change(*args):
            mode = self.merge_mode.get()
            if mode == "consolidate":
                cb_source.config(state="normal")
                cb_align.config(state="normal")
                cb_smart.config(state="normal")
            elif mode == "auto_group":
                cb_source.config(state="normal")
                cb_align.config(state="disabled") # 自动分组强制要求结构一致，对齐无意义
                cb_smart.config(state="normal")
            else:
                cb_source.config(state="disabled")
                cb_align.config(state="disabled")
                cb_smart.config(state="disabled")
        self.merge_mode.trace("w", on_mode_change)

        # --- 执行按钮 ---
        btn_frame = ttk.Frame(self.parent, padding=10)
        btn_frame.pack(fill="x")
        
        ttk.Button(btn_frame, text="开始合并", command=self._start_merge).pack(side="right", padx=10)
        ttk.Button(btn_frame, text="合并并转入凭证转换", command=lambda: self._start_merge(open_in_converter=True)).pack(side="right", padx=10)
        
        # 状态栏
        self.status_label = ttk.Label(self.parent, text="就绪")
        self.status_label.pack(fill="x", padx=10, pady=(0, 5))

    def _add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
        if files:
            for f in files:
                if f not in self.file_list:
                    self.file_list.append(f)
                    stats = os.stat(f)
                    size_kb = f"{stats.st_size / 1024:.1f}"
                    mtime = datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    self.tree.insert("", "end", values=(f, size_kb, mtime))
            self.status_label.config(text=f"已添加 {len(files)} 个文件，共 {len(self.file_list)} 个待合并")

    def _clear_files(self):
        self.file_list.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_label.config(text="列表已清空")

    def _remove_selected(self):
        selected = self.tree.selection()
        for item in selected:
            vals = self.tree.item(item, "values")
            path = vals[0]
            if path in self.file_list:
                self.file_list.remove(path)
            self.tree.delete(item)
        self.status_label.config(text=f"剩余 {len(self.file_list)} 个文件")

    def _start_merge(self, open_in_converter=False):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要合并的 Excel 文件")
            return
            
        out_path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not out_path:
            return
            
        try:
            self.status_label.config(text="正在合并中...")
            self.parent.update()
            
            mode = self.merge_mode.get()
            
            if mode == "consolidate":
                self._merge_consolidate(out_path)
            elif mode == "auto_group":
                self._merge_auto_group(out_path)
            else:
                self._merge_workbook(out_path)
                
            self.status_label.config(text=f"合并完成: {out_path}")
            messagebox.showinfo("成功", f"文件已合并至:\n{out_path}")
            if open_in_converter and self.open_in_converter:
                self.open_in_converter(out_path)
            
        except Exception as e:
            self.status_label.config(text="合并失败")
            messagebox.showerror("错误", f"合并过程中发生错误:\n{e}")
            import traceback
            traceback.print_exc()

    def _merge_consolidate(self, out_path):
        """模式1: 合并所有内容到一个 Sheet"""
        dfs = []
        master_header_row = None # 用于智能去重表头：存储第一个文件的表头行内容
        
        use_smart_header = self.smart_header.get()
        # 如果启用智能表头检查，我们必须用 header=None 读取，以便手动检查第一行
        # 如果未启用，并且 align_columns=True，我们用 standard header=0
        
        # 统一策略：为了支持 "Smart Header"，我们最好都用 header=None 读取，然后手动处理 columns
        # 只有当 !use_smart_header 且 align_columns=True 时，才回退到 header=0 以利用pandas的自动对齐
        
        # 但是，为了实现 "If header matches master, skip it"，header=None 是必须的。
        # 让我们统一用 header=None，然后根据 user preference 决定怎么处理 row 0
        
        read_kwargs = {}
        if not use_smart_header:
             # 原有逻辑：默认 header=0
             read_kwargs["header"] = 0
        else:
             # 智能逻辑：header=None (手动处理)
             read_kwargs["header"] = None

        for f_idx, f in enumerate(self.file_list):
            try:
                xls = pd.ExcelFile(f)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, **read_kwargs)
                    if df.empty:
                        continue
                    
                    # --- Smart Header Logic ---
                    if use_smart_header:
                        # 1. 确定Master Header (从第一个非空文件的第一行)
                        if master_header_row is None:
                            master_header_row = df.iloc[0].values
                            # 将第一行提升为表头
                            df.columns = master_header_row
                            df = df.iloc[1:] # Drop header row from data body
                            # Reset index is not strictly needed for concat but good for cleanliness
                            df.reset_index(drop=True, inplace=True)
                        else:
                            # 2. 检查当前文件第一行是否与 Master Header 匹配
                            current_first = df.iloc[0].values
                            
                            # 简单比对：长度相同且内容大致相同
                            is_header = False
                            if len(current_first) == len(master_header_row):
                                # 允许少量差异 (比如 "日期" vs "Date"?) 不，先做严格相等或包含
                                # 这里做宽松相等：如果 80% 的列名相同
                                match_count = np.sum(current_first == master_header_row)
                                if match_count / len(master_header_row) > 0.8:
                                    is_header = True
                            
                            if is_header:
                                # 是表头 -> 使用它作为列名 (对齐列)，并丢弃该行
                                df.columns = current_first # 使用自己的表头以便 align (如果 align_columns=True)
                                df = df.iloc[1:]
                            else:
                                # 不是表头 -> 是数据
                                # 强制使用 Master Header 作为列名 (Position Alignment)
                                # 只有当 align_columns=False 时才应该完全忽略列名?
                                # 如果 user 选了 align_columns=True，但这里没表头，那只能按位置了
                                if len(df.columns) == len(master_header_row):
                                     df.columns = master_header_row
                                else:
                                     # 列数不一致，这就麻烦了。
                                     # 如果 align_columns=True，我们希望按名字对齐。但如果没有名字...
                                     # 只能保留默认 index columns (0, 1, 2...)
                                     pass
                            
                            df.reset_index(drop=True, inplace=True)

                    # --- End Smart Header Logic ---

                    if self.add_source_col.get():
                        # 添加来源文件名
                        fname = os.path.basename(f)
                        df.insert(0, "来源文件", fname)
                        df.insert(1, "来源工作表", sheet_name)
                    
                    dfs.append(df)
            except Exception as e:
                print(f"读取文件失败 {f}: {e}")
                
        if not dfs:
            raise ValueError("没有读取到有效数据")
            
        # 合并
        # sort=False 禁止重排顺序
        # join="outer" (align columns) or "inner"
        
        if not self.align_columns.get():
            # 强制按位置对齐：重命名所有列为 0, 1, 2...
            # 如果使用了 Smart Header，master file 的 columns 已经被设为 text header 了
            # 我们需要重置它们
            # 但是，Smart Header 同时也意味着 "Identify header row".
            # 如果 align_columns=False，我们最终不需要表头? 或者我们只需要第一行的表头?
            # 通常 align_columns=False 意味着 "Trust column order, ignore names".
            
            # 为了保留第一个文件的表头 (User expectation)，我们不能简单地 range(len)
            # 我们应该：让所有 DF 的 columns = First DF's columns
            
            final_cols = list(dfs[0].columns)
            new_dfs = []
            for df in dfs:
                # Rename to 0..N
                df_reset = df.copy()
                df_reset.columns = range(len(df.columns))
                new_dfs.append(df_reset)
            dfs = new_dfs
            
            # Concat
            result = pd.concat(dfs, ignore_index=True, sort=False)
            
            # Restore header if possible
            if len(result.columns) == len(final_cols):
                result.columns = final_cols
                
        else:
            # Align by Name (Standard)
            result = pd.concat(dfs, ignore_index=True, sort=False)
            
        result.to_excel(out_path, index=False)

    def _merge_auto_group(self, out_path):
        """模式3: 智能自动分组 (结构相同合并，不同分表)"""
        
        # groups structure:
        # [
        #   {
        #     "signature": (col1, col2, ...), # columns tuple for comparison
        #     "dfs": [df1, df2, ...],
        #     "name": "SheetName"
        #   },
        #   ...
        # ]
        groups = []
        
        use_smart_header = self.smart_header.get()
        # Auto-group usually implies distinct structures, so aligning by name is implicit if we group by name
        # But here we group by "Structure Identity".
        # We will read with header=0 (or None if smart) to get columns.
        
        read_kwargs = {"header": None} if use_smart_header else {"header": 0}

        for f_idx, f in enumerate(self.file_list):
            try:
                xls = pd.ExcelFile(f)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name, **read_kwargs)
                    if df.empty:
                        continue
                        
                    # --- Determine Structure Signature ---
                    # Logic:
                    # 1. If Smart Header is on, treat row 0 as header.
                    # 2. Signature = tuple(row 0 values)
                    # 3. If matched, drop row 0 (except for the very first one in group) and append.
                    
                    if use_smart_header:
                        current_header = tuple(df.iloc[0].fillna("").astype(str).values)
                        data_df = df.iloc[1:].copy()
                        data_df.reset_index(drop=True, inplace=True)
                        data_df.columns = current_header # Assign header to data so concat works
                    else:
                        # Use pandas detected columns
                        current_header = tuple(df.columns.astype(str))
                        data_df = df
                    
                    # --- Find Matching Group ---
                    matched_group = None
                    for group in groups:
                        # Compare signatures
                        # Exact match required for "Structure Identity"
                        if group["signature"] == current_header:
                            matched_group = group
                            break
                    
                    # --- Add to Group ---
                    if matched_group:
                        # Add source info if needed
                        if self.add_source_col.get():
                            fname = os.path.basename(f)
                            data_df.insert(0, "来源文件", fname)
                            data_df.insert(1, "来源工作表", sheet_name)
                        matched_group["dfs"].append(data_df)
                    else:
                        # New Group
                        # Name strategy: Use file name + sheet name (but unique)
                        base_name = f"{os.path.splitext(os.path.basename(f))[0]}_{sheet_name}"
                        # Truncate
                        if len(base_name) > 25: base_name = base_name[:25]
                        
                        # Add source info
                        if self.add_source_col.get():
                            fname = os.path.basename(f)
                            data_df.insert(0, "来源文件", fname)
                            data_df.insert(1, "来源工作表", sheet_name)
                            
                        groups.append({
                            "signature": current_header,
                            "dfs": [data_df],
                            "base_name": base_name
                        })
                        
            except Exception as e:
                print(f"读取文件失败 {f}: {e}")

        if not groups:
            raise ValueError("没有读取到有效数据")

        # --- Write Output ---
        with pd.ExcelWriter(out_path) as writer:
            existing_sheet_names = set()
            
            for i, group in enumerate(groups):
                # Determine Sheet Name
                sheet_name = group["base_name"]
                
                # Deduplicate sheet name
                if sheet_name in existing_sheet_names:
                    idx = 1
                    while f"{sheet_name}_{idx}" in existing_sheet_names:
                        idx += 1
                    sheet_name = f"{sheet_name}_{idx}"
                
                existing_sheet_names.add(sheet_name)
                
                # Concat DFs
                merged_df = pd.concat(group["dfs"], ignore_index=True, sort=False)
                
                # Write
                merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Group {i}: {sheet_name} (Cols: {len(group['signature'])}, Rows: {len(merged_df)})")

    def _merge_workbook(self, out_path):
        """模式2: 每个文件/Sheet 作为一个 Sheet"""
        with pd.ExcelWriter(out_path) as writer:
            existing_sheets = set()
            
            for f in self.file_list:
                try:
                    xls = pd.ExcelFile(f)
                    fname = os.path.splitext(os.path.basename(f))[0]
                    
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        
                        # 生成唯一 Sheet 名
                        target_sheet_name = f"{fname}_{sheet_name}"
                        if len(xls.sheet_names) == 1:
                            target_sheet_name = fname
                        
                        # 防止名字太长 (Excel limit 31 chars)
                        if len(target_sheet_name) > 31:
                            target_sheet_name = target_sheet_name[:31]
                            
                        # 防止重复
                        base_name = target_sheet_name
                        idx = 1
                        while target_sheet_name in existing_sheets:
                            suffix = f"_{idx}"
                            trim_len = 31 - len(suffix)
                            target_sheet_name = base_name[:trim_len] + suffix
                            idx += 1
                        
                        existing_sheets.add(target_sheet_name)
                        df.to_excel(writer, sheet_name=target_sheet_name, index=False)
                        
                except Exception as e:
                     print(f"处理文件失败 {f}: {e}")
