# -*- coding: utf-8 -*-
"""
图片智能识别GUI模块
提供图片导入、识别、预览和导出功能
"""

import os
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Dict, Any, Optional
from datetime import datetime
from treeview_tools import attach_treeview_tools
from export_format_manager import (
    get_active_export_format_name,
    get_export_format_names,
    open_export_format_editor,
    set_active_export_format,
)

# 尝试导入图像处理库
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 尝试导入图片识别模块
try:
    from image_intelligence import ImageIntelligence, check_and_install_dependencies
    HAS_IMAGE_INTELLIGENCE = True
except ImportError:
    HAS_IMAGE_INTELLIGENCE = False


class ImageRecognitionWindow:
    """图片智能识别窗口"""

    def __init__(self, parent, api_key: str = "", ai_provider: str = "zhipu",
                 base_url: str = "http://localhost:1234/v1", model_name: str = "local-model",
                 template_path: str = "Template.xlsx",
                 default_engine: str = "auto",
                 tesseract_cmd: str = "",
                 tesseract_lang: str = "chi_sim+eng+por"):
        """
        初始化图片识别窗口

        Args:
            parent: 父窗口
            api_key: AI API密钥
            ai_provider: AI提供商
            base_url: LM Studio URL
            model_name: 模型名称
            template_path: 模板文件路径
            default_engine: 默认识别引擎 (auto/zhipu/lm_studio/tesseract/paddleocr/easyocr)
            tesseract_cmd: Tesseract可执行文件路径
            tesseract_lang: Tesseract语言包
        """
        self.parent = parent
        self.api_key = api_key
        self.ai_provider = ai_provider
        self.base_url = base_url
        self.model_name = model_name
        self.template_path = template_path
        self.default_engine = default_engine
        self.tesseract_cmd = tesseract_cmd
        self.tesseract_lang = tesseract_lang

        self.image_files: List[str] = []
        self.export_format_var = tk.StringVar(value=get_active_export_format_name("image_recognition"))
        self.recognition_results: List[Dict[str, Any]] = []
        self.merged_headers: List[str] = []
        self.merged_rows: List[List[str]] = []
        self.current_image_index = 0

        self.recognizer: Optional[ImageIntelligence] = None

        self._create_window()
        self._init_recognizer()

    def _create_window(self):
        """创建主窗口"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("图片智能识别")
        self.window.geometry("1200x800")
        self.window.minsize(900, 600)

        # 主框架
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 顶部：工具栏
        self._create_toolbar(main_frame)
        self._refresh_export_format_options()

        # 中部：主内容区
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill="both", expand=True, pady=10)

        # 左侧：图片列表和预览
        self._create_image_panel(content_frame)

        # 右侧：识别结果
        self._create_result_panel(content_frame)

        # 底部：状态栏和进度
        self._create_status_bar(main_frame)

    def _create_toolbar(self, parent):
        """创建工具栏"""
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 10))

        # 导入图片按钮
        ttk.Button(toolbar, text="导入图片", command=self._import_images, width=12).pack(side="left", padx=2)
        ttk.Button(toolbar, text="导入文件夹", command=self._import_folder, width=12).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)

        # 识别按钮
        ttk.Button(toolbar, text="识别当前", command=self._recognize_current, width=10).pack(side="left", padx=2)
        ttk.Button(toolbar, text="批量识别", command=self._batch_recognize, width=10).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)

        # 一键操作按钮（突出显示）
        one_click_btn = ttk.Button(toolbar, text="一键识别导出", command=self._one_click_recognize_export, width=14)
        one_click_btn.pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)

        # 导出按钮
        ttk.Button(toolbar, text="导出Excel", command=self._export_excel, width=10).pack(side="left", padx=2)
        ttk.Button(toolbar, text="按模板导出", command=self._export_with_template, width=12).pack(side="left", padx=2)
        ttk.Label(toolbar, text="导出格式:").pack(side="left", padx=(8, 2))
        self.export_format_combo = ttk.Combobox(
            toolbar,
            textvariable=self.export_format_var,
            values=get_export_format_names("image_recognition"),
            state="readonly",
            width=12
        )
        self.export_format_combo.pack(side="left", padx=2)
        self.export_format_combo.bind("<<ComboboxSelected>>", self._on_export_format_changed)
        ttk.Button(toolbar, text="设置", command=self._open_export_format_editor, width=6).pack(side="left", padx=2)

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10)

        # 识别设置
        ttk.Button(toolbar, text="识别设置", command=self._show_ai_settings, width=10).pack(side="left", padx=2)

        # 清空按钮
        ttk.Button(toolbar, text="清空列表", command=self._clear_all, width=10).pack(side="right", padx=2)

    def _on_export_format_changed(self, event=None):
        name = self.export_format_var.get().strip()
        set_active_export_format("image_recognition", name)

    def _open_export_format_editor(self):
        headers = list(self.merged_headers or [])
        open_export_format_editor(self.window, "image_recognition", headers, title="导出格式设置 - 图片识别")
        self._refresh_export_format_options()

    def _refresh_export_format_options(self):
        if hasattr(self, "export_format_combo"):
            names = get_export_format_names("image_recognition")
            self.export_format_combo["values"] = names
            active = get_active_export_format_name("image_recognition")
            if active:
                self.export_format_var.set(active)

    def _create_image_panel(self, parent):
        """创建图片面板"""
        left_frame = ttk.LabelFrame(parent, text="图片列表", width=400)
        left_frame.pack(side="left", fill="both", expand=False, padx=(0, 5))
        left_frame.pack_propagate(False)

        # 图片列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # 列表框
        self.image_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, width=40)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.image_listbox.yview)
        self.image_listbox.configure(yscrollcommand=scrollbar.set)

        self.image_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.image_listbox.bind("<<ListboxSelect>>", self._on_image_select)
        self.image_listbox.bind("<Double-1>", self._on_image_double_click)

        # 图片预览
        preview_frame = ttk.LabelFrame(left_frame, text="图片预览", height=300)
        preview_frame.pack(fill="x", padx=5, pady=5)
        preview_frame.pack_propagate(False)

        self.preview_label = ttk.Label(preview_frame, text="选择图片以预览", anchor="center")
        self.preview_label.pack(fill="both", expand=True, padx=5, pady=5)

        # 图片操作按钮
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", padx=5, pady=5)

        ttk.Button(btn_frame, text="上移", command=self._move_up, width=8).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="下移", command=self._move_down, width=8).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="删除", command=self._remove_selected, width=8).pack(side="left", padx=2)

    def _create_result_panel(self, parent):
        """创建识别结果面板"""
        right_frame = ttk.LabelFrame(parent, text="识别结果")
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # 创建Notebook用于显示不同视图
        self.result_notebook = ttk.Notebook(right_frame)
        self.result_notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # 标签页1：表格视图
        table_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(table_frame, text="表格预览")

        # 表格
        self.result_tree = ttk.Treeview(table_frame, show="headings")
        tree_scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.result_tree.yview)
        tree_scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=tree_scrollbar_y.set, xscrollcommand=tree_scrollbar_x.set)

        self.result_tree.pack(side="left", fill="both", expand=True)
        tree_scrollbar_y.pack(side="right", fill="y")
        tree_scrollbar_x.pack(side="bottom", fill="x")
        self.result_tree_tools = attach_treeview_tools(self.result_tree)

        # 标签页2：原始文本
        text_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(text_frame, text="识别文本")

        self.raw_text = tk.Text(text_frame, wrap="word")
        text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.raw_text.yview)
        self.raw_text.configure(yscrollcommand=text_scrollbar.set)

        self.raw_text.pack(side="left", fill="both", expand=True)
        text_scrollbar.pack(side="right", fill="y")

        # 标签页3：合并结果
        merged_frame = ttk.Frame(self.result_notebook)
        self.result_notebook.add(merged_frame, text="合并数据")

        self.merged_tree = ttk.Treeview(merged_frame, show="headings")
        merged_scrollbar_y = ttk.Scrollbar(merged_frame, orient="vertical", command=self.merged_tree.yview)
        merged_scrollbar_x = ttk.Scrollbar(merged_frame, orient="horizontal", command=self.merged_tree.xview)
        self.merged_tree.configure(yscrollcommand=merged_scrollbar_y.set, xscrollcommand=merged_scrollbar_x.set)

        self.merged_tree.pack(side="left", fill="both", expand=True)
        merged_scrollbar_y.pack(side="right", fill="y")
        merged_scrollbar_x.pack(side="bottom", fill="x")
        self.merged_tree_tools = attach_treeview_tools(self.merged_tree)

    def _create_status_bar(self, parent):
        """创建状态栏"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill="x", pady=(10, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)

        # 进度条
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, length=200)
        self.progress_bar.pack(side="right", padx=10)

    def _init_recognizer(self):
        """初始化识别器"""
        try:
            from image_intelligence import ImageIntelligence, check_and_install_dependencies
        except Exception as e:
            self.status_var.set(f"图片识别模块加载失败: {e}")
            return

        dep_status = check_and_install_dependencies(auto_install=False)
        missing = [k for k, ok in dep_status.items() if not ok]
        if missing:
            allow_install = messagebox.askyesno(
                "依赖缺失",
                "检测到以下依赖缺失：\n\n"
                f"{', '.join(missing)}\n\n"
                "是否尝试自动安装？"
            )
            if allow_install:
                self.status_var.set("正在安装依赖，请稍候...")
                self.window.update()
                try:
                    dep_status = check_and_install_dependencies(auto_install=True)
                    self.status_var.set(f"依赖安装完成: {dep_status}")
                except Exception as e:
                    self.status_var.set(f"依赖安装失败: {e}")
                    return
            else:
                self.status_var.set("依赖未安装，部分识别功能不可用")

        try:
            self.recognizer = ImageIntelligence(
                ai_provider=self.ai_provider,
                api_key=self.api_key,
                base_url=self.base_url,
                model_name=self.model_name,
                default_engine=self.default_engine,
                tesseract_cmd=self.tesseract_cmd,
                tesseract_lang=self.tesseract_lang,
                auto_install=False
            )
            # 显示可用引擎
            available = self.recognizer.get_available_engines()
            available_str = ", ".join(available) if available else "无"
            self.status_var.set(f"识别器初始化成功 (可用引擎: {available_str})")
        except Exception as e:
            self.status_var.set(f"识别器初始化失败: {e}")

    def _import_images(self):
        """导入图片"""
        filetypes = [
            ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"),
            ("JPEG", "*.jpg *.jpeg"),
            ("PNG", "*.png"),
            ("所有文件", "*.*")
        ]

        files = filedialog.askopenfilenames(
            title="选择图片",
            filetypes=filetypes
        )

        if files:
            for f in files:
                if f not in self.image_files:
                    self.image_files.append(f)
                    self.image_listbox.insert(tk.END, os.path.basename(f))
                    self.recognition_results.append(None)

            self.status_var.set(f"已导入 {len(files)} 张图片，共 {len(self.image_files)} 张")

    def _import_folder(self):
        """导入文件夹中的所有图片"""
        folder = filedialog.askdirectory(title="选择图片文件夹")

        if folder:
            image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
            count = 0

            for filename in sorted(os.listdir(folder)):
                ext = os.path.splitext(filename)[1].lower()
                if ext in image_extensions:
                    filepath = os.path.join(folder, filename)
                    if filepath not in self.image_files:
                        self.image_files.append(filepath)
                        self.image_listbox.insert(tk.END, filename)
                        self.recognition_results.append(None)
                        count += 1

            self.status_var.set(f"已导入 {count} 张图片，共 {len(self.image_files)} 张")

    def _on_image_select(self, event):
        """图片选择事件"""
        selection = self.image_listbox.curselection()
        if selection:
            index = selection[0]
            self.current_image_index = index
            self._show_preview(index)
            self._show_result(index)

    def _on_image_double_click(self, event):
        """双击图片触发识别"""
        self._recognize_current()

    def _show_preview(self, index):
        """显示图片预览"""
        if not HAS_PIL:
            self.preview_label.configure(text="需要安装Pillow库")
            return

        if 0 <= index < len(self.image_files):
            try:
                img = Image.open(self.image_files[index])

                # 缩放以适应预览区域
                max_size = (380, 280)
                img.thumbnail(max_size, Image.Resampling.LANCZOS)

                photo = ImageTk.PhotoImage(img)
                self.preview_label.configure(image=photo, text="")
                self.preview_label.image = photo  # 保持引用

            except Exception as e:
                self.preview_label.configure(text=f"无法加载图片: {e}", image="")

    def _show_result(self, index):
        """显示识别结果"""
        if 0 <= index < len(self.recognition_results):
            result = self.recognition_results[index]

            # 清空现有内容
            self.raw_text.delete(1.0, tk.END)
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)

            if result is None:
                self.raw_text.insert(tk.END, "尚未识别，请点击'识别当前'按钮")
                return

            # 显示原始文本
            raw_text = result.get("raw_text", "")
            self.raw_text.insert(tk.END, raw_text)

            # 显示表格
            headers = result.get("headers", [])
            rows = result.get("rows", [])

            if headers:
                self.result_tree["columns"] = headers
                for col in headers:
                    self.result_tree.heading(col, text=col)
                    self.result_tree.column(col, width=120, minwidth=80)

                for row in rows:
                    # 确保行数据与列数匹配
                    padded_row = list(row) + [""] * (len(headers) - len(row))
                    self.result_tree.insert("", tk.END, values=padded_row[:len(headers)])

    def _recognize_current(self):
        """识别当前选中的图片"""
        selection = self.image_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一张图片")
            return

        if not self.recognizer:
            messagebox.showerror("错误", "识别器未初始化")
            return

        index = selection[0]
        image_path = self.image_files[index]

        self.status_var.set(f"正在识别: {os.path.basename(image_path)}")
        self.progress_var.set(0)
        self.window.update()

        def do_recognize():
            try:
                result = self.recognizer.recognize_image(image_path, use_ai=True)
                self.recognition_results[index] = result

                # 更新UI（在主线程中）
                self.window.after(0, lambda: self._on_recognize_complete(index, result))
            except Exception as e:
                self.window.after(0, lambda: self._on_recognize_error(str(e)))

        # 在后台线程中执行识别
        thread = threading.Thread(target=do_recognize)
        thread.daemon = True
        thread.start()

    def _on_recognize_complete(self, index, result):
        """识别完成回调"""
        status = result.get("status", "error")
        if status == "success":
            self.status_var.set(f"识别成功: {len(result.get('rows', []))} 条数据")
        elif status == "partial":
            self.status_var.set(f"部分识别成功: {result.get('message', '')}")
        else:
            self.status_var.set(f"识别失败: {result.get('message', '未知错误')}")

        self.progress_var.set(100)
        self._show_result(index)
        self._update_merged_data()

    def _on_recognize_error(self, error_msg):
        """识别错误回调"""
        self.status_var.set(f"识别错误: {error_msg}")
        self.progress_var.set(0)

    def _batch_recognize(self):
        """批量识别所有图片"""
        if not self.image_files:
            messagebox.showwarning("提示", "请先导入图片")
            return

        if not self.recognizer:
            messagebox.showerror("错误", "识别器未初始化")
            return

        total = len(self.image_files)
        self.status_var.set(f"开始批量识别 (0/{total})")
        self.progress_var.set(0)
        self.window.update()

        def do_batch():
            for i, image_path in enumerate(self.image_files):
                try:
                    self.window.after(0, lambda idx=i: self.status_var.set(
                        f"正在识别 ({idx+1}/{total}): {os.path.basename(self.image_files[idx])}"
                    ))

                    result = self.recognizer.recognize_image(image_path, use_ai=True)
                    self.recognition_results[i] = result

                    progress = (i + 1) / total * 100
                    self.window.after(0, lambda p=progress: self.progress_var.set(p))

                except Exception as e:
                    self.recognition_results[i] = {"status": "error", "message": str(e)}

            self.window.after(0, self._on_batch_complete)

        thread = threading.Thread(target=do_batch)
        thread.daemon = True
        thread.start()

    def _on_batch_complete(self):
        """批量识别完成"""
        success_count = sum(1 for r in self.recognition_results if r and r.get("status") == "success")
        self.status_var.set(f"批量识别完成: {success_count}/{len(self.image_files)} 成功")
        self._update_merged_data()

        # 显示第一张图片的结果
        if self.image_files:
            self.image_listbox.selection_set(0)
            self._show_result(0)

    def _one_click_recognize_export(self):
        """一键批量识别并导出"""
        if not self.image_files:
            # 如果没有图片，先让用户选择文件夹
            folder = filedialog.askdirectory(title="选择包含图片的文件夹")
            if not folder:
                return

            # 导入文件夹中的图片
            image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp'}
            for filename in sorted(os.listdir(folder)):
                ext = os.path.splitext(filename)[1].lower()
                if ext in image_extensions:
                    filepath = os.path.join(folder, filename)
                    if filepath not in self.image_files:
                        self.image_files.append(filepath)
                        self.image_listbox.insert(tk.END, filename)
                        self.recognition_results.append(None)

            if not self.image_files:
                messagebox.showwarning("提示", "所选文件夹中没有找到图片文件")
                return

        if not self.recognizer:
            messagebox.showerror("错误", "识别器未初始化")
            return

        # 询问保存位置
        filename = filedialog.asksaveasfilename(
            title="选择保存位置（识别完成后自动导出）",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"图片识别合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not filename:
            return

        total = len(self.image_files)
        self.status_var.set(f"开始一键识别导出 (0/{total})")
        self.progress_var.set(0)
        self.window.update()

        def do_one_click():
            # 批量识别
            for i, image_path in enumerate(self.image_files):
                try:
                    self.window.after(0, lambda idx=i: self.status_var.set(
                        f"正在识别 ({idx+1}/{total}): {os.path.basename(self.image_files[idx])}"
                    ))

                    result = self.recognizer.recognize_image(image_path, use_ai=True)
                    self.recognition_results[i] = result

                    progress = (i + 1) / total * 90  # 留10%给导出
                    self.window.after(0, lambda p=progress: self.progress_var.set(p))

                except Exception as e:
                    self.recognition_results[i] = {"status": "error", "message": str(e)}

            # 合并并导出
            self.window.after(0, lambda: self.status_var.set("正在合并数据并导出..."))

            # 使用智能合并
            headers, rows = self.recognizer.merge_results_to_table(
                self.recognition_results, smart_merge=True
            )

            if rows:
                try:
                    success = self.recognizer.export_to_excel(headers, rows, filename)
                    self.window.after(0, lambda: self._on_one_click_complete(
                        len(rows), filename, success
                    ))
                except Exception as e:
                    self.window.after(0, lambda: self._on_one_click_complete(
                        0, filename, False, str(e)
                    ))
            else:
                self.window.after(0, lambda: self._on_one_click_complete(0, filename, False, "没有识别到数据"))

        thread = threading.Thread(target=do_one_click)
        thread.daemon = True
        thread.start()

    def _on_one_click_complete(self, row_count, filename, success, error_msg=None):
        """一键操作完成回调"""
        self.progress_var.set(100)
        self._update_merged_data()

        # 显示第一张图片的结果
        if self.image_files:
            self.image_listbox.selection_set(0)
            self._show_result(0)

        # 自动切换到合并数据标签页
        self.result_notebook.select(2)

        if success:
            success_count = sum(1 for r in self.recognition_results if r and r.get("status") == "success")
            self.status_var.set(f"一键识别导出完成: {success_count}/{len(self.image_files)} 成功, {row_count} 条数据")
            messagebox.showinfo(
                "一键识别导出完成",
                f"识别完成!\n\n"
                f"- 处理图片: {len(self.image_files)} 张\n"
                f"- 识别成功: {success_count} 张\n"
                f"- 合并数据: {row_count} 条\n\n"
                f"已导出到:\n{filename}"
            )
        else:
            self.status_var.set(f"导出失败: {error_msg or '未知错误'}")
            messagebox.showerror("导出失败", error_msg or "导出时发生错误")

    def _update_merged_data(self):
        """更新合并数据视图（使用智能合并）"""
        # 清空
        for item in self.merged_tree.get_children():
            self.merged_tree.delete(item)

        # 使用识别器的智能合并功能
        if self.recognizer:
            all_headers, all_rows = self.recognizer.merge_results_to_table(
                self.recognition_results, smart_merge=True
            )
        else:
            # 备用：简单合并
            all_headers = []
            all_rows = []
            for result in self.recognition_results:
                if result and result.get("status") in ["success", "partial"]:
                    headers = result.get("headers", [])
                    rows = result.get("rows", [])
                    if not all_headers and headers:
                        all_headers = headers
                    all_rows.extend(rows)

        self.merged_headers = all_headers
        self.merged_rows = all_rows

        # 更新表格
        if all_headers:
            self.merged_tree["columns"] = all_headers
            for col in all_headers:
                self.merged_tree.heading(col, text=col)
                self.merged_tree.column(col, width=120, minwidth=80)

            for row in all_rows:
                padded_row = list(row) + [""] * (len(all_headers) - len(row))
                self.merged_tree.insert("", tk.END, values=padded_row[:len(all_headers)])

        # 更新合并数据标签页标题显示数据条数
        if all_rows:
            self.result_notebook.tab(2, text=f"合并数据 ({len(all_rows)}条)")

    def _export_excel(self):
        """导出为Excel"""
        if not self.merged_rows:
            messagebox.showwarning("提示", "没有可导出的数据")
            return

        filename = filedialog.asksaveasfilename(
            title="导出Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"图片识别结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if filename:
            try:
                # 获取当前视觉显示的表头和行（尊重列管理中的排序和隐藏）
                if hasattr(self.merged_tree, "_treeview_tools"):
                    headers, rows = self.merged_tree._treeview_tools.get_visual_data()
                else:
                    headers, rows = self.merged_headers, self.merged_rows

                if self.recognizer:
                    success = self.recognizer.export_to_excel(
                        headers,
                        rows,
                        filename
                    )
                else:
                    # 备选方案：直接使用pandas
                    import pandas as pd
                    df = pd.DataFrame(rows, columns=headers)
                    df.to_excel(filename, index=False)
                    success = True

                if success:
                    self.status_var.set(f"已导出到: {filename}")
                    messagebox.showinfo("成功", f"已导出 {len(self.merged_rows)} 条数据到:\n{filename}")
                else:
                    messagebox.showerror("错误", "导出失败")

            except Exception as e:
                messagebox.showerror("导出失败", str(e))

    def _export_with_template(self):
        """按模板格式导出"""
        if not self.merged_rows:
            messagebox.showwarning("提示", "没有可导出的数据")
            return

        # 选择模板
        template = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=self.template_path
        )

        if not template:
            return

        # 选择保存位置
        filename = filedialog.asksaveasfilename(
            title="保存转换结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"按模板导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if filename:
            try:
                if self.recognizer:
                    success = self.recognizer.export_to_excel(
                        self.merged_headers,
                        self.merged_rows,
                        filename,
                        template_path=template
                    )
                else:
                    messagebox.showerror("错误", "识别器未初始化")
                    return

                if success:
                    self.status_var.set(f"已按模板导出到: {filename}")
                    messagebox.showinfo("成功", f"已按模板导出 {len(self.merged_rows)} 条数据到:\n{filename}")
                else:
                    messagebox.showerror("错误", "导出失败")

            except Exception as e:
                messagebox.showerror("导出失败", str(e))

    def _show_ai_settings(self):
        """显示AI和识别引擎设置对话框"""
        dialog = tk.Toplevel(self.window)
        dialog.title("识别设置")
        dialog.geometry("580x580")
        dialog.transient(self.window)
        dialog.grab_set()

        # 创建滚动框架
        canvas = tk.Canvas(dialog)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

        frame = scrollable_frame

        row_idx = 0

        # ========== 识别引擎设置 ==========
        ttk.Label(frame, text="识别引擎设置", font=("", 10, "bold")).grid(
            row=row_idx, column=0, columnspan=3, sticky="w", pady=(10, 5))
        row_idx += 1

        ttk.Separator(frame, orient="horizontal").grid(
            row=row_idx, column=0, columnspan=3, sticky="ew", pady=5)
        row_idx += 1

        # 默认识别引擎
        ttk.Label(frame, text="默认引擎:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        engine_var = tk.StringVar(value=self.default_engine)
        engine_values = ["auto", "zhipu", "lm_studio", "tesseract", "paddleocr", "easyocr"]
        engine_combo = ttk.Combobox(frame, textvariable=engine_var,
                                     values=engine_values, state="readonly", width=25)
        engine_combo.grid(row=row_idx, column=1, sticky="w", padx=5, pady=5)

        # 显示可用引擎
        available_engines = []
        if self.recognizer:
            available_engines = self.recognizer.get_available_engines()
        available_str = ", ".join(available_engines) if available_engines else "正在检测..."
        ttk.Label(frame, text=f"(可用: {available_str})", foreground="gray").grid(
            row=row_idx, column=2, sticky="w", padx=5, pady=5)
        row_idx += 1

        # 引擎说明
        engine_info = """• auto: 自动选择最佳引擎 (推荐)
• zhipu: 智谱AI云视觉模型 (最准确，需联网)
• lm_studio: 本地LM Studio (需配置)
• tesseract: Tesseract OCR (免费，需安装)
• paddleocr: PaddleOCR (中文效果好)
• easyocr: EasyOCR (多语言支持)"""
        ttk.Label(frame, text=engine_info, foreground="gray", justify="left").grid(
            row=row_idx, column=0, columnspan=3, sticky="w", padx=20, pady=5)
        row_idx += 1

        # ========== Tesseract 设置 ==========
        ttk.Label(frame, text="Tesseract OCR 设置", font=("", 10, "bold")).grid(
            row=row_idx, column=0, columnspan=3, sticky="w", pady=(15, 5))
        row_idx += 1

        ttk.Separator(frame, orient="horizontal").grid(
            row=row_idx, column=0, columnspan=3, sticky="ew", pady=5)
        row_idx += 1

        # Tesseract 路径
        ttk.Label(frame, text="Tesseract路径:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        tesseract_cmd_var = tk.StringVar(value=self.tesseract_cmd)
        tesseract_entry = ttk.Entry(frame, textvariable=tesseract_cmd_var, width=35)
        tesseract_entry.grid(row=row_idx, column=1, sticky="w", padx=5, pady=5)

        def browse_tesseract():
            filepath = filedialog.askopenfilename(
                title="选择Tesseract可执行文件",
                filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
            )
            if filepath:
                tesseract_cmd_var.set(filepath)

        ttk.Button(frame, text="浏览...", command=browse_tesseract, width=8).grid(
            row=row_idx, column=2, sticky="w", padx=5, pady=5)
        row_idx += 1

        ttk.Label(frame, text="(留空则自动检测)", foreground="gray").grid(
            row=row_idx, column=1, sticky="w", padx=5)
        row_idx += 1

        # Tesseract 语言
        ttk.Label(frame, text="语言包:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        tesseract_lang_var = tk.StringVar(value=self.tesseract_lang)
        lang_values = [
            "chi_sim+eng",
            "chi_sim+eng+por",
            "chi_sim+chi_tra+eng",
            "eng",
            "por",
            "chi_sim",
            "chi_tra"
        ]
        lang_combo = ttk.Combobox(frame, textvariable=tesseract_lang_var,
                                   values=lang_values, width=25)
        lang_combo.grid(row=row_idx, column=1, sticky="w", padx=5, pady=5)
        row_idx += 1

        lang_info = """常用语言包:
• chi_sim: 简体中文
• chi_tra: 繁体中文
• eng: 英语
• por: 葡萄牙语
用 + 号组合多语言"""
        ttk.Label(frame, text=lang_info, foreground="gray", justify="left").grid(
            row=row_idx, column=0, columnspan=3, sticky="w", padx=20, pady=5)
        row_idx += 1

        # ========== AI 服务设置 ==========
        ttk.Label(frame, text="AI 服务设置", font=("", 10, "bold")).grid(
            row=row_idx, column=0, columnspan=3, sticky="w", pady=(15, 5))
        row_idx += 1

        ttk.Separator(frame, orient="horizontal").grid(
            row=row_idx, column=0, columnspan=3, sticky="ew", pady=5)
        row_idx += 1

        # AI提供商选择
        ttk.Label(frame, text="AI提供商:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        provider_var = tk.StringVar(value=self.ai_provider)
        provider_combo = ttk.Combobox(frame, textvariable=provider_var,
                                       values=["zhipu", "lm_studio"], state="readonly", width=25)
        provider_combo.grid(row=row_idx, column=1, sticky="w", padx=5, pady=5)
        row_idx += 1

        # API Key
        ttk.Label(frame, text="API Key:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        key_var = tk.StringVar(value=self.api_key)
        ttk.Entry(frame, textvariable=key_var, width=40, show="*").grid(
            row=row_idx, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        row_idx += 1

        # Base URL (LM Studio)
        ttk.Label(frame, text="Base URL:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        url_var = tk.StringVar(value=self.base_url)
        ttk.Entry(frame, textvariable=url_var, width=40).grid(
            row=row_idx, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        row_idx += 1

        # 模型名称
        ttk.Label(frame, text="模型名称:").grid(row=row_idx, column=0, sticky="e", pady=5, padx=5)
        model_var = tk.StringVar(value=self.model_name)
        ttk.Entry(frame, textvariable=model_var, width=40).grid(
            row=row_idx, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        row_idx += 1

        ai_info = """• zhipu: 智谱AI云服务 (推荐glm-4v-flash)
• lm_studio: 本地LM Studio (需配置URL和模型)"""
        ttk.Label(frame, text=ai_info, foreground="gray", justify="left").grid(
            row=row_idx, column=0, columnspan=3, sticky="w", padx=20, pady=5)
        row_idx += 1

        # ========== 按钮 ==========
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row_idx, column=0, columnspan=3, pady=20)

        def save_settings():
            # 保存所有设置
            self.default_engine = engine_var.get()
            self.tesseract_cmd = tesseract_cmd_var.get()
            self.tesseract_lang = tesseract_lang_var.get()
            self.ai_provider = provider_var.get()
            self.api_key = key_var.get()
            self.base_url = url_var.get()
            self.model_name = model_var.get()

            # 重新初始化识别器
            self._init_recognizer()

            messagebox.showinfo("成功", "设置已保存并生效")
            dialog.destroy()

        def test_engine():
            """测试当前选择的引擎"""
            selected_engine = engine_var.get()
            test_msg = f"测试引擎: {selected_engine}\n\n"

            if self.recognizer:
                available = self.recognizer.get_available_engines()
                if selected_engine == "auto":
                    test_msg += f"auto模式将自动选择可用引擎\n可用引擎: {', '.join(available)}"
                elif selected_engine in available:
                    test_msg += f"引擎 '{selected_engine}' 可用！"
                else:
                    test_msg += f"警告: 引擎 '{selected_engine}' 不可用\n请安装相应依赖或选择其他引擎\n\n可用引擎: {', '.join(available)}"
            else:
                test_msg += "识别器未初始化，无法测试"

            messagebox.showinfo("引擎测试", test_msg)

        ttk.Button(btn_frame, text="保存", command=save_settings, width=10).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="测试引擎", command=test_engine, width=10).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side="left", padx=5)

    def _move_up(self):
        """上移选中项"""
        selection = self.image_listbox.curselection()
        if selection and selection[0] > 0:
            idx = selection[0]
            # 交换数据
            self.image_files[idx], self.image_files[idx-1] = self.image_files[idx-1], self.image_files[idx]
            self.recognition_results[idx], self.recognition_results[idx-1] = self.recognition_results[idx-1], self.recognition_results[idx]

            # 更新列表显示
            item = self.image_listbox.get(idx)
            self.image_listbox.delete(idx)
            self.image_listbox.insert(idx-1, item)
            self.image_listbox.selection_set(idx-1)

    def _move_down(self):
        """下移选中项"""
        selection = self.image_listbox.curselection()
        if selection and selection[0] < len(self.image_files) - 1:
            idx = selection[0]
            # 交换数据
            self.image_files[idx], self.image_files[idx+1] = self.image_files[idx+1], self.image_files[idx]
            self.recognition_results[idx], self.recognition_results[idx+1] = self.recognition_results[idx+1], self.recognition_results[idx]

            # 更新列表显示
            item = self.image_listbox.get(idx)
            self.image_listbox.delete(idx)
            self.image_listbox.insert(idx+1, item)
            self.image_listbox.selection_set(idx+1)

    def _remove_selected(self):
        """删除选中项"""
        selection = self.image_listbox.curselection()
        if selection:
            idx = selection[0]
            self.image_files.pop(idx)
            self.recognition_results.pop(idx)
            self.image_listbox.delete(idx)

            # 更新合并数据
            self._update_merged_data()

            self.status_var.set(f"已删除，剩余 {len(self.image_files)} 张图片")

    def _clear_all(self):
        """清空所有"""
        if self.image_files:
            if messagebox.askyesno("确认", "确定要清空所有图片吗？"):
                self.image_files.clear()
                self.recognition_results.clear()
                self.merged_headers.clear()
                self.merged_rows.clear()

                self.image_listbox.delete(0, tk.END)

                for item in self.result_tree.get_children():
                    self.result_tree.delete(item)
                for item in self.merged_tree.get_children():
                    self.merged_tree.delete(item)

                self.raw_text.delete(1.0, tk.END)
                self.preview_label.configure(text="选择图片以预览", image="")

                self.status_var.set("已清空")


def open_image_recognition_window(parent, api_key="", ai_provider="zhipu",
                                   base_url="http://localhost:1234/v1",
                                   model_name="local-model",
                                   template_path="Template.xlsx",
                                   default_engine="auto",
                                   tesseract_cmd="",
                                   tesseract_lang="chi_sim+eng+por"):
    """
    打开图片识别窗口的便捷函数

    Args:
        parent: 父窗口
        api_key: API密钥
        ai_provider: AI提供商
        base_url: LM Studio URL
        model_name: 模型名称
        template_path: 模板文件路径
        default_engine: 默认识别引擎
        tesseract_cmd: Tesseract可执行文件路径
        tesseract_lang: Tesseract语言包
    """
    return ImageRecognitionWindow(
        parent=parent,
        api_key=api_key,
        ai_provider=ai_provider,
        base_url=base_url,
        model_name=model_name,
        template_path=template_path,
        default_engine=default_engine,
        tesseract_cmd=tesseract_cmd,
        tesseract_lang=tesseract_lang
    )


if __name__ == "__main__":
    # 独立运行测试
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 打开图片识别窗口
    window = ImageRecognitionWindow(root)

    # 当图片识别窗口关闭时退出程序
    window.window.protocol("WM_DELETE_WINDOW", lambda: (root.quit(), root.destroy()))

    root.mainloop()
