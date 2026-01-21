#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX ↔ Markdown 批量转换工具
图形化界面，支持批量转换 Word 文档和 Markdown 文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import sys
import os

# 添加 src 目录到路径
# PyInstaller 打包后使用 sys._MEIPASS，否则使用当前文件目录
if getattr(sys, 'frozen', False):
    # 打包后的情况：模块在 sys._MEIPASS/docx2markdown 目录下
    base_path = sys._MEIPASS
    # 添加 docx2markdown 目录到路径
    sys.path.insert(0, os.path.join(base_path, 'docx2markdown'))
    # 也添加父目录，以防万一
    sys.path.insert(0, base_path)
else:
    # 开发环境：模块在 src 目录下
    base_path = os.path.dirname(__file__)
    src_path = os.path.join(base_path, 'src')
    sys.path.insert(0, src_path)

try:
    from docx2markdown import docx_to_markdown, markdown_to_docx
except ImportError as e:
    error_msg = f"无法导入 docx2markdown 模块: {str(e)}\n路径: {src_path}"
    try:
        messagebox.showerror("错误", error_msg)
    except:
        print(error_msg)
    sys.exit(1)


class ConversionTab:
    """转换标签页基类"""
    def __init__(self, parent, conversion_type="docx2md"):
        self.parent = parent
        self.conversion_type = conversion_type  # "docx2md" 或 "md2docx"
        self.file_list = []
        self.output_folder = ""
        self.create_widgets()
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.parent.columnconfigure(0, weight=1)
        self.parent.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="10")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)
        
        # 文件列表和滚动条
        listbox_frame = ttk.Frame(file_frame)
        listbox_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.file_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, height=15)
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.config(command=self.file_listbox.yview)
        
        # 文件操作按钮
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Button(button_frame, text="添加文件", command=self.add_files).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="移除选中", command=self.remove_file).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(button_frame, text="清空列表", command=self.clear_files).grid(row=0, column=2, padx=(0, 5))
        
        # 输出文件夹选择
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="10")
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="输出文件夹:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var, state="readonly")
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="选择文件夹", command=self.select_output_folder).grid(row=0, column=2)
        
        # 转换按钮和进度
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        action_frame.columnconfigure(0, weight=1)
        
        self.convert_button = ttk.Button(action_frame, text="开始转换", command=self.start_conversion)
        self.convert_button.grid(row=0, column=0, pady=5)
        
        # 进度条
        self.progress_var = tk.StringVar(value="就绪")
        progress_label = ttk.Label(action_frame, textvariable=self.progress_var)
        progress_label.grid(row=1, column=0, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(action_frame, mode='determinate')
        self.progress_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=3, column=0, sticky=(tk.W, tk.E))
    
    def add_files(self):
        """添加文件到列表"""
        if self.conversion_type == "docx2md":
            filetypes = [("Word 文档", "*.docx"), ("所有文件", "*.*")]
            title = "选择 DOCX 文件⭐"
        else:
            filetypes = [("Markdown 文件", "*.md"), ("所有文件", "*.*")]
            title = "选择 MD 文件🎈"
        
        files = filedialog.askopenfilenames(title=title, filetypes=filetypes)
        for file_path in files:
            if file_path not in self.file_list:
                self.file_list.append(file_path)
        self.update_file_listbox()
    
    def remove_file(self):
        """移除选中的文件"""
        selected = self.file_listbox.curselection()
        if selected:
            index = selected[0]
            del self.file_list[index]
            self.update_file_listbox()
    
    def clear_files(self):
        """清空文件列表"""
        self.file_list.clear()
        self.update_file_listbox()
    
    def update_file_listbox(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        for file_path in self.file_list:
            self.file_listbox.insert(tk.END, Path(file_path).name)
        self.status_var.set(f"已选择 {len(self.file_list)} 个文件")
    
    def select_output_folder(self):
        """选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_folder = folder
            self.output_path_var.set(folder)
    
    def start_conversion(self):
        """开始转换"""
        if not self.file_list:
            messagebox.showwarning("警告", "请先添加要转换的文件")
            return
        
        if not self.output_folder:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return
        
        # 在新线程中执行转换，避免界面冻结
        thread = threading.Thread(target=self.convert_files, daemon=True)
        thread.start()
    
    def convert_files(self):
        """批量转换文件"""
        total = len(self.file_list)
        success_count = 0
        fail_count = 0
        
        # 禁用转换按钮
        self.convert_button.config(state="disabled")
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = 0
        
        for index, input_file in enumerate(self.file_list, 1):
            try:
                input_path = Path(input_file)
                
                if self.conversion_type == "docx2md":
                    # DOCX to MD
                    output_filename = input_path.stem + ".md"
                    output_path = Path(self.output_folder) / output_filename
                    self.progress_var.set(f"正在转换 {index}/{total}: {input_path.name}")
                    docx_to_markdown(str(input_path), str(output_path))
                else:
                    # MD to DOCX
                    output_filename = input_path.stem + ".docx"
                    output_path = Path(self.output_folder) / output_filename
                    self.progress_var.set(f"正在转换 {index}/{total}: {input_path.name}")
                    markdown_to_docx(str(input_path), str(output_path))
                
                success_count += 1
                
            except Exception as e:
                fail_count += 1
                print(f"转换失败 {input_file}: {e}")
            
            # 更新进度条
            self.progress_bar['value'] = index
            self.parent.update_idletasks()
        
        # 转换完成
        self.progress_var.set(f"转换完成！成功: {success_count}, 失败: {fail_count}")
        self.convert_button.config(state="normal")
        # 打开输出文件夹
        os.startfile(self.output_folder)
        
        # 显示完成消息
        messagebox.showinfo("完成", 
                          f"转换完成！\n成功: {success_count}\n失败: {fail_count}")


class Docx2MarkdownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX ↔ Markdown 批量转换工具")
        self.root.geometry("900x700")
        
        # 设置窗口图标（如果存在）
        icon_path = Path(__file__).parent / "icon.ico"
        if icon_path.exists():
            try:
                self.root.iconbitmap(str(icon_path))
            except:
                pass  # 如果图标加载失败，继续运行
        
        # 支持拖拽（如果可用）
        self.drag_drop_enabled = False
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            self.root = TkinterDnD.Tk() if not isinstance(root, TkinterDnD.Tk) else root
            self.root.title("DOCX ↔ Markdown 批量转换工具")
            self.root.geometry("800x500")
            if icon_path.exists():
                try:
                    self.root.iconbitmap(str(icon_path))
                except:
                    pass
            self.drag_drop_enabled = True
        except ImportError:
            # 如果没有 tkinterdnd2，使用普通模式
            pass
        
        self.create_widgets()
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="DOCX ↔ Markdown 批量转换工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # 创建标签页
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # DOCX to MD 标签页
        tab1 = ttk.Frame(notebook, padding="10")
        notebook.add(tab1, text="DOCX → MD")
        self.tab1_converter = ConversionTab(tab1, "docx2md")
        
        # MD to DOCX 标签页
        tab2 = ttk.Frame(notebook, padding="10")
        notebook.add(tab2, text="MD → DOCX")
        self.tab2_converter = ConversionTab(tab2, "md2docx")
        
        # 设置拖拽功能
        if self.drag_drop_enabled:
            self.setup_drag_drop()
    
    def setup_drag_drop(self):
        """设置拖拽功能"""
        try:
            from tkinterdnd2 import DND_FILES
            # 为两个标签页的文件列表设置拖拽
            self.tab1_converter.file_listbox.drop_target_register(DND_FILES)
            self.tab1_converter.file_listbox.dnd_bind('<<Drop>>', 
                lambda e: self.on_drop(e, self.tab1_converter, '.docx'))
            
            self.tab2_converter.file_listbox.drop_target_register(DND_FILES)
            self.tab2_converter.file_listbox.dnd_bind('<<Drop>>', 
                lambda e: self.on_drop(e, self.tab2_converter, '.md'))
        except:
            pass
    
    def on_drop(self, event, converter, file_ext):
        """处理文件拖拽事件"""
        try:
            from tkinterdnd2 import DND_FILES
            files = self.root.tk.splitlist(event.data)
            for file_path in files:
                file_path = file_path.strip('{}')  # 移除可能的括号
                if file_path.lower().endswith(file_ext):
                    if file_path not in converter.file_list:
                        converter.file_list.append(file_path)
                        converter.update_file_listbox()
        except Exception as e:
            messagebox.showerror("错误", f"拖拽文件失败: {str(e)}")


def main():
    """主函数"""
    root = tk.Tk()
    app = Docx2MarkdownGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
