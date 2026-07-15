# -*- coding: utf-8 -*-
"""
文件夹递归处理模块 (原生拖拽支持版)
功能：将一个或多个文件夹及其子文件夹中的所有文件提取并汇总到一个统一的目录中。
支持：原生 Windows 拖拽、多路径选择、递归扫描、同名冲突自动重命名。
"""

import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import ctypes
from ctypes import wintypes

# --- Windows 原生拖拽支持 (无需第三方库) ---
class WindowsDropHandler:
    """利用 Windows API 实现拖放支持"""
    def __init__(self, widget, callback):
        self.widget = widget
        self.callback = callback
        self.old_window_proc = None
        
        # 定义必要的常量和结构
        self.WM_DROPFILES = 0x0233
        self.GWL_WNDPROC = -4
        
        # 获取 shell32 中的函数并设置类型
        self.shell32 = ctypes.windll.shell32
        self.DragAcceptFiles = self.shell32.DragAcceptFiles
        self.DragAcceptFiles.restype = None
        self.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        
        self.DragQueryFile = self.shell32.DragQueryFileW
        self.DragQueryFile.restype = wintypes.UINT
        self.DragQueryFile.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        
        self.DragFinish = self.shell32.DragFinish
        self.DragFinish.restype = None
        self.DragFinish.argtypes = [ctypes.c_void_p]
        
        # 绑定窗口
        self.hwnd = widget.winfo_id()
        self.DragAcceptFiles(self.hwnd, True)
        
        # 准备 API 函数及其参数类型 (解决 64 位兼容性问题)
        user32 = ctypes.windll.user32
        
        # 根据位数选择正确的函数名
        if ctypes.sizeof(ctypes.c_void_p) == 8:
            self.GetWindowLong = user32.GetWindowLongPtrW
            self.SetWindowLong = user32.SetWindowLongPtrW
        else:
            self.GetWindowLong = user32.GetWindowLongW
            self.SetWindowLong = user32.SetWindowLongW
            
        self.GetWindowLong.restype = ctypes.c_ssize_t
        self.GetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int]
        
        self.SetWindowLong.restype = ctypes.c_ssize_t
        self.SetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
        
        self.CallWindowProc = user32.CallWindowProcW
        self.CallWindowProc.restype = ctypes.c_ssize_t
        self.CallWindowProc.argtypes = [ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, ctypes.c_ssize_t, ctypes.c_ssize_t]
        
        # 定义 WNDPROC 回调类型 (LRESULT 是 c_ssize_t)
        self._callback_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, ctypes.c_ssize_t, ctypes.c_ssize_t)
        self.new_window_proc = self._callback_type(self._window_proc)
        
        # 替换窗口过程
        proc_ptr = ctypes.cast(self.new_window_proc, ctypes.c_void_p).value
        self.old_window_proc = self.GetWindowLong(self.hwnd, self.GWL_WNDPROC)
        if self.old_window_proc:
            self.SetWindowLong(self.hwnd, self.GWL_WNDPROC, proc_ptr)
        
        # 绑定销毁事件：关键修复，必须在窗口销毁时解除挂钩，否则 Python 3.13 会崩溃
        widget.bind("<Destroy>", lambda e: self.unregister(), add="+")
        
        # 额外加固：将引用挂载到 widget 上
        if not hasattr(widget, "_drop_handler_refs"):
            widget._drop_handler_refs = []
        widget._drop_handler_refs.append(self.new_window_proc)

    def unregister(self):
        """恢复原始窗口过程，防止非法回调"""
        if self.old_window_proc:
            try:
                self.SetWindowLong(self.hwnd, self.GWL_WNDPROC, self.old_window_proc)
                self.old_window_proc = None
            except:
                pass

    def _window_proc(self, hwnd, msg, wparam, lparam):
        if msg == self.WM_DROPFILES:
            try:
                hdrop = wparam
                num_files = self.DragQueryFile(hdrop, 0xFFFFFFFF, None, 0)
                paths = []
                for i in range(num_files):
                    length = self.DragQueryFile(hdrop, i, None, 0)
                    if length > 0:
                        buf = ctypes.create_unicode_buffer(length + 1)
                        self.DragQueryFile(hdrop, i, buf, length + 1)
                        paths.append(buf.value)
                self.DragFinish(hdrop)
                # 使用 after 异步处理，避免在 WndProc 回调中占用 GIL 太久 (Python 3.13 敏感)
                if self.widget.winfo_exists():
                    self.widget.after(10, lambda p=paths: self.callback(p))
                return 0
            except Exception as e:
                # 记录但不抛出，避免阻塞消息循环
                print(f"Drag-and-drop proc error: {e}")
                return 0
        
        # 如果 self.old_window_proc 为空，说明已经注销或出错
        if not self.old_window_proc:
            return 0
        
        try:
            return self.CallWindowProc(self.old_window_proc, hwnd, msg, wparam, lparam)
        except Exception:
            # 捕获异常以防止 Python 3.13 的致命错误
            return 0

class FolderProcessorGUI:
    def __init__(self, parent_frame):
        self.parent = parent_frame
        self.source_paths = []
        self.target_dir = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.operation_type = tk.StringVar(value="copy")
        
        self._build_ui()
        
        # 启用拖放功能
        try:
            self.drop_handler = WindowsDropHandler(self.path_listbox, self._handle_dropped_paths)
        except Exception as e:
            print(f"拖放功能启用失败 (可能非 Windows 系统): {e}")

    def _build_ui(self):
        main_frame = ttk.LabelFrame(self.parent, text="文件夹平铺汇总 (支持拖入文件夹)", padding=15)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        src_label_frame = ttk.LabelFrame(main_frame, text="待提取的源文件夹列表 (可将文件夹直接拖入此处)", padding=5)
        src_label_frame.pack(fill="both", expand=True, pady=5)

        list_frame = ttk.Frame(src_label_frame)
        list_frame.pack(fill="both", expand=True, side="left")
        
        self.path_listbox = tk.Listbox(list_frame, height=10, selectmode="extended", font=("微软雅黑", 9), bg="#fcfcfc")
        self.path_listbox.pack(fill="both", expand=True, side="left")
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.path_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.path_listbox.config(yscrollcommand=scrollbar.set)

        btn_sidebar = ttk.Frame(src_label_frame)
        btn_sidebar.pack(side="right", fill="y", padx=5)
        
        ttk.Button(btn_sidebar, text="添加单个文件夹...", command=self._add_single_source).pack(fill="x", pady=2)
        ttk.Button(btn_sidebar, text="批量添加子目录...", command=self._batch_add_subfolders).pack(fill="x", pady=2)
        ttk.Button(btn_sidebar, text="移除选中", command=self._remove_source).pack(fill="x", pady=5)
        ttk.Button(btn_sidebar, text="清空列表", command=self._clear_sources).pack(fill="x", pady=2)

        tgt_frame = ttk.Frame(main_frame)
        tgt_frame.pack(fill="x", pady=10)
        ttk.Label(tgt_frame, text="汇总存放目标位置:").pack(side="left")
        ttk.Entry(tgt_frame, textvariable=self.target_dir).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(tgt_frame, text="浏览...", command=self._select_target).pack(side="left")

        opt_frame = ttk.Frame(main_frame)
        opt_frame.pack(fill="x", pady=5)
        ttk.Checkbutton(opt_frame, text="包含子文件夹 (递归扫描)", variable=self.recursive).pack(side="left", padx=10)
        ttk.Separator(opt_frame, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Radiobutton(opt_frame, text="复制文件 (安全)", variable=self.operation_type, value="copy").pack(side="left", padx=5)
        ttk.Radiobutton(opt_frame, text="移动文件 (剪切)", variable=self.operation_type, value="move").pack(side="left", padx=5)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=15)
        self.process_btn = ttk.Button(btn_frame, text="开始执行批量提取", command=self._start_process)
        self.process_btn.pack(side="right", padx=5)
        
        self.status_var = tk.StringVar(value="提示: 您可以从 Windows 文件夹中直接拖入多个文件夹到列表中")
        ttk.Label(main_frame, textvariable=self.status_var, foreground="#336699", wraplength=600).pack(side="bottom", fill="x")

    def _handle_dropped_paths(self, paths):
        """处理拖入的文件/文件夹路径"""
        added_count = 0
        for path in paths:
            path = os.path.normpath(path)
            if os.path.isdir(path):
                if self._add_to_list(path, verbose=False):
                    added_count += 1
            else:
                # 如果拖入的是文件，则将其父目录加入列表（或者您可以改为忽略文件）
                parent = os.path.dirname(path)
                if self._add_to_list(parent, verbose=False):
                    added_count += 1
        
        if added_count > 0:
            self.status_var.set(f"通过拖拽添加了 {added_count} 个路径")

    def _add_single_source(self):
        path = filedialog.askdirectory(title="选择要添加的源文件夹")
        if path:
            self._add_to_list(os.path.normpath(path))

    def _batch_add_subfolders(self):
        parent_path = filedialog.askdirectory(title="选择包含多个子文件夹的父目录")
        if not parent_path:
            return
        parent_path = os.path.normpath(parent_path)
        added_count = 0
        try:
            for item in os.listdir(parent_path):
                full_path = os.path.join(parent_path, item)
                if os.path.isdir(full_path):
                    if self._add_to_list(full_path, verbose=False):
                        added_count += 1
            self.status_var.set(f"批量添加成功: 已加入 {added_count} 个文件夹")
        except Exception as e:
            messagebox.showerror("错误", f"读取目录失败: {e}")

    def _add_to_list(self, path, verbose=True):
        if path not in self.source_paths:
            self.source_paths.append(path)
            self.path_listbox.insert(tk.END, path)
            self.path_listbox.see(tk.END)
            if not self.target_dir.get():
                default_target = os.path.join(os.path.dirname(path), f"汇总提取_{datetime.now().strftime('%m%d%H%M')}")
                self.target_dir.set(os.path.normpath(default_target))
            return True
        elif verbose:
            messagebox.showinfo("提示", "该文件夹已在列表中")
        return False

    def _remove_source(self):
        selected = self.path_listbox.curselection()
        for index in reversed(selected):
            path = self.path_listbox.get(index)
            self.source_paths.remove(path)
            self.path_listbox.delete(index)

    def _clear_sources(self):
        if self.source_paths and messagebox.askyesno("确认", "清空列表？"):
            self.source_paths.clear()
            self.path_listbox.delete(0, tk.END)

    def _select_target(self):
        path = filedialog.askdirectory(title="选择目标目录")
        if path:
            self.target_dir.set(os.path.normpath(path))

    def _start_process(self):
        if not self.source_paths:
            messagebox.showwarning("错误", "列表为空")
            return
        dst = self.target_dir.get()
        if not dst:
            messagebox.showwarning("错误", "请选择目标位置")
            return

        dst_norm = os.path.normpath(dst)
        op_name = "复制" if self.operation_type.get() == "copy" else "移动"
        if not messagebox.askyesno("确认", f"执行 {op_name} 操作到: {dst}"):
            return

        try:
            self.process_btn.config(state="disabled")
            if not os.path.exists(dst):
                os.makedirs(dst)

            total_count = 0
            errors = 0
            for src_idx, src in enumerate(self.source_paths):
                self.status_var.set(f"处理中 ({src_idx+1}/{len(self.source_paths)}): {os.path.basename(src)}")
                self.parent.update_idletasks()
                for root, dirs, files in os.walk(src):
                    norm_root = os.path.normpath(root)
                    
                    # 避免处理目标目录本身（如果它在源目录内）
                    if norm_root == dst_norm:
                        dirs[:] = []  # 不再进入子目录
                        continue
                        
                    if not self.recursive.get() and norm_root != os.path.normpath(src):
                        dirs[:] = []
                        continue
                    
                    for file in files:
                        src_file = os.path.join(root, file)
                        target_file = self._get_unique_path(dst, file)
                        try:
                            if self.operation_type.get() == "copy":
                                shutil.copy2(src_file, target_file)
                            else:
                                shutil.move(src_file, target_file)
                            total_count += 1
                            if total_count % 5 == 0:
                                self.status_var.set(f"已处理 {total_count} 个文件...")
                                self.parent.update_idletasks()
                        except Exception as e:
                            print(f"Error processing {src_file}: {e}")
                            errors += 1

            messagebox.showinfo("完成", f"总计提取: {total_count}, 失败: {errors}")
        except Exception as e:
            messagebox.showerror("错误", f"{e}")
        finally:
            self.process_btn.config(state="normal")

    def _get_unique_path(self, folder, filename):
        base, ext = os.path.splitext(filename)
        counter = 1
        path = os.path.join(folder, filename)
        while os.path.exists(path):
            path = os.path.join(folder, f"{base}({counter}){ext}")
            counter += 1
        return path

if __name__ == "__main__":
    root = tk.Tk()
    root.title("拖拽汇总测试")
    root.geometry("700x550")
    app = FolderProcessorGUI(root)
    root.mainloop()
