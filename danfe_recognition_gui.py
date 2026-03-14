# -*- coding: utf-8 -*-
"""
巴西NF-e (DANFE) 文档识别GUI模块
支持PDF文件导入、数据提取、预览和导出到标准凭证格式
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
from danfe_recognition_module import DanfeRecognizer
from treeview_tools import attach_treeview_tools

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

class DanfeRecognitionWindow:
    """DANFE识别窗口"""

    def __init__(self, parent):
        self.parent = parent
        self.files: list = []
        self.results: list = []
        self.recognizer = DanfeRecognizer()
        
        self._create_window()

    def _create_window(self):
        self.window = tk.Toplevel(self.parent)
        self.window.title("巴西NF-e (DANFE) 文档智能识别")
        self.window.geometry("1000x700")
        
        # 工具栏
        toolbar = ttk.Frame(self.window, padding=5)
        toolbar.pack(fill="x")
        
        ttk.Button(toolbar, text="导入PDF文件", command=self._import_files).pack(side="left", padx=2)
        ttk.Button(toolbar, text="识别所有", command=self._recognize_all).pack(side="left", padx=2)
        ttk.Button(toolbar, text="导出到Excel", command=self._export_excel).pack(side="left", padx=2)
        ttk.Button(toolbar, text="导出所有明细", command=self._export_comprehensive).pack(side="left", padx=2)
        ttk.Button(toolbar, text="清空", command=self._clear_all).pack(side="right", padx=2)
        
        # 主内容
        paned = ttk.PanedWindow(self.window, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 左侧列表
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        self.file_listbox = tk.Listbox(left_frame)
        self.file_listbox.pack(fill="both", expand=True)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)
        
        # 右侧结果
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=3)
        
        self.tree = ttk.Treeview(right_frame, show="headings")
        self.tree.pack(fill="both", expand=True)
        
        cols = ["发票号码", "日期", "总金额", "ICMS", "PIS/COFINS", "发行人", "发行人CNPJ", "收件人", "收件人ID", "Access Key"]
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80 if "税" in col or "金额" in col or "日期" in col else 120)
        
        attach_treeview_tools(self.tree)
        
        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.window, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x")

    def _import_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")])
        if files:
            for f in files:
                if f not in self.files:
                    self.files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self.status_var.set(f"已导入 {len(self.files)} 个文件")

    def _recognize_all(self):
        if not self.files:
            messagebox.showwarning("提示", "请先导入PDF文件")
            return
        
        if not HAS_PDFPLUMBER:
            messagebox.showerror("错误", "未找到 pdfplumber 库，无法解析PDF内容。请运行: pip install pdfplumber")
            return

        self.status_var.set("正在识别中...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        results = []
        total = len(self.files)
        
        for i, file_path in enumerate(self.files):
            self.window.after(0, lambda idx=i: self.status_var.set(f"正在识别 ({idx+1}/{total}): {os.path.basename(file_path)}"))
            try:
                text = ""
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                
                if not text.strip():
                    # 如果PDF没有文本，可能是扫描件
                    res = {
                        'status': 'error', 
                        'message': 'PDF无文本，可能是扫描件，请使用图片识别模块',
                        'file_path': file_path
                    }
                else:
                    res = self.recognizer.recognize_from_text(text)
                    res['status'] = 'success'
                    res['file_path'] = file_path
                
                results.append(res)
            except Exception as e:
                results.append({
                    'status': 'error', 
                    'message': str(e), 
                    'file_path': file_path
                })
        
        self.results = results
        self.window.after(0, self._update_tree)

    def _update_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        error_count = 0
        for res in self.results:
            if res.get('status') == 'error':
                error_count += 1
                # 在列表中显示错误
                vals = [
                    "ERROR",
                    "",
                    0.0,
                    "",
                    "",
                    res.get('message', '未知错误'),
                    "",
                    "",
                    "",
                    os.path.basename(res.get('file_path', ''))
                ]
                self.tree.insert("", tk.END, values=vals, tags=('error',))
                continue
            
            icms = res.get('v_icms', 0.0)
            if res.get('v_icms_st', 0.0) > 0:
                icms = f"{icms} (+ST:{res.get('v_icms_st')})"
                
            pis_cofins = f"{res.get('v_pis', 0.0)} / {res.get('v_cofins', 0.0)}"

            vals = [
                res.get('numero_nota', ''),
                res.get('data_emissao', ''),
                res.get('valor_total', 0.0),
                icms,
                pis_cofins,
                res.get('emitente_nome', ''),
                res.get('emitente_cnpj', ''),
                res.get('destinatario_nome', ''),
                res.get('destinatario_cnpj_cpf', ''),
                res.get('chave_acesso', '')
            ]
            self.tree.insert("", tk.END, values=vals)
        
        self.tree.tag_configure('error', foreground='red')
        
        msg = f"识别完成，共 {len(self.results)} 个文件"
        if error_count > 0:
            msg += f" (其中 {error_count} 个失败)"
        self.status_var.set(msg)

    def _on_file_select(self, event):
        # 可以在此处添加文件详情预览
        pass

    def _export_excel(self):
        if not self.results:
            messagebox.showwarning("提示", "没有可导出的数据")
            return
        
        df = self.recognizer.to_standard_voucher([r for r in self.results if r.get('status') != 'error'])
        if df.empty:
            messagebox.showwarning("提示", "没有成功的识别结果")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"DANFE凭证导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if file_path:
            try:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"数据已导出到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")

    def _export_comprehensive(self):
        if not self.results:
            messagebox.showwarning("提示", "当前没有识别结果，请先点击“识别所有”")
            return
        
        success_results = [r for r in self.results if r.get('status') == 'success']
        if not success_results:
            messagebox.showwarning("导出失败", "没有成功的识别结果可供导出。请检查文件是否为扫描件或识别过程中是否有报错。")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"DANFE所有明细导出_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if file_path:
            try:
                # 显式检查 DataFrame 转换
                df = self.recognizer.to_comprehensive_dataframe(success_results)
                if df.empty:
                    messagebox.showwarning("提示", "转换后的数据集为空，请确认识别到的内容是否有效。")
                    return
                
                df.to_excel(file_path, index=False)
                messagebox.showinfo("成功", f"共 {len(success_results)} 个文件的详尽明细已导出到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")

    def _clear_all(self):
        self.files = []
        self.results = []
        self.file_listbox.delete(0, tk.END)
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("已清空")

def open_danfe_recognition_window(parent):
    return DanfeRecognitionWindow(parent)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    window = DanfeRecognitionWindow(root)
    window.window.protocol("WM_DELETE_WINDOW", lambda: (root.quit(), root.destroy()))
    root.mainloop()
