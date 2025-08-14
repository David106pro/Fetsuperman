import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import requests
from typing import List, Dict
import threading
import os
import webbrowser
import pandas as pd

class BatchAPIProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("批量接口处理工具")
        self.root.geometry("800x700")  # 调整窗口初始大小
        # 创建主框架并设置权重
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk. E, tk.N, tk.S))
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # 创建选项卡控件
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # 创建选项卡页面
        self.excel_frame = ttk.Frame(self.notebook, padding="10")
        self.media_frame = ttk.Frame(self.notebook, padding="10")
        self.batch_sync_frame = ttk.Frame(self.notebook, padding="10")
        
        self.notebook.add(self.excel_frame, text="Excel导入")
        self.notebook.add(self.media_frame, text="介质修改")
        self.notebook.add(self.batch_sync_frame, text="批量下发")
        
        # 项目选项
        self.projects = {
            "极智": "jz",
            "金胡桃": "kn",
            "测试": "abc",
            "宁夏": "nx",
            "福建": "fj",
            "江西": "jx",
            "陕西 线上": "sn",
            "甘肃OTT":"gs_ott",
            "甘肃OIPTV": "gs_iptv",
            "河南":"ha",
        }
        
        # Excel导入页面
        self.create_excel_page()
        
        # 媒资修改页面
        self.create_media_page()
        
        # 批量下发页面
        self.create_batch_sync_page()
        
        self.is_processing = False
        self.excel_data = None
        self.excel_file_handler = None
        self.media_file_handler = None
        self.batch_sync_file_handler = None

    def create_excel_page(self):
        # 配置Excel页面的网格权重
        self.excel_frame.grid_columnconfigure(0, weight=1)
        
        # Excel导入区域
        excel_import_frame = ttk.LabelFrame(self.excel_frame, text="Excel导入与选择", padding="10")
        excel_import_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        # Configure columns for layout
        excel_import_frame.grid_columnconfigure(1, weight=1) # Path entry expands
        excel_import_frame.grid_columnconfigure(4, weight=0) # Sheet combo doesn't expand excessively
        
        # --- File and Sheet Selection Row ---
        ttk.Label(excel_import_frame, text="文件:").grid(row=0, column=0, padx=(5, 0), pady=5, sticky=tk.W) # Label for path
        self.excel_path_var = tk.StringVar()
        path_entry = ttk.Entry(
            excel_import_frame, 
            textvariable=self.excel_path_var, 
            state='readonly',
            font=('Consolas', 10)
        )
        path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5) # Path entry
        
        select_button = ttk.Button(
            excel_import_frame,
            text="选择文件",
            command=self.select_excel_file,
            width=12
        )
        select_button.grid(row=0, column=2, padx=(0, 5), pady=5) # Select File button

        ttk.Label(excel_import_frame, text="Sheet:").grid(row=0, column=3, padx=(10, 0), pady=5, sticky=tk.W) # Label for sheet
        self.excel_sheet_var = tk.StringVar()
        self.excel_sheet_combo = ttk.Combobox(
            excel_import_frame, 
            textvariable=self.excel_sheet_var,
            state="disabled",
            width=20 # Adjust width as needed
        )
        self.excel_sheet_combo.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(5, 10), pady=5) # Sheet combobox
        
        # 预览区域
        preview_frame = ttk.LabelFrame(self.excel_frame, text="数据预览 (首个Sheet)", padding="10")
        # Update row index for preview, button, and progress bar
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # 预览文本框
        self.excel_preview_text = scrolledtext.ScrolledText(
            preview_frame,
            width=50,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 10)
        )
        self.excel_preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        self.excel_preview_text.config(state='disabled')
        
        # 控制按钮区域
        control_frame = ttk.Frame(self.excel_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10)) # Updated row index
        control_frame.grid_columnconfigure(0, weight=1)
        
        # Excel修改按钮
        self.excel_modify_button = ttk.Button(
            control_frame, 
            text="开始修改",
            command=lambda: self.start_modification('excel'),
            width=20,
            state='disabled'
        )
        self.excel_modify_button.grid(row=0, column=0)
        
        # 进度条区域
        progress_frame_excel = ttk.Frame(self.excel_frame)
        progress_frame_excel.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5) # Updated row index
        progress_frame_excel.grid_columnconfigure(0, weight=1)
        
        self.excel_progress_var = tk.DoubleVar()
        self.excel_progress = ttk.Progressbar(
            progress_frame_excel,
            variable=self.excel_progress_var,
            maximum=100
        )
        self.excel_progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.excel_progress_label = ttk.Label(progress_frame_excel, text="0%", width=5)
        self.excel_progress_label.grid(row=0, column=1)

    def create_media_page(self):
        # 配置媒资修改页面的网格权重
        self.media_frame.grid_columnconfigure(0, weight=1)
        
        # 文件与Sheet选择区域
        file_select_frame = ttk.LabelFrame(self.media_frame, text="文件与Sheet选择", padding="10")
        file_select_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        # Configure columns
        file_select_frame.grid_columnconfigure(1, weight=1) # Path expands
        file_select_frame.grid_columnconfigure(4, weight=0) # Sheet combo fixed width

        # --- File and Sheet Selection Row ---
        ttk.Label(file_select_frame, text="文件:").grid(row=0, column=0, padx=(5, 0), pady=5, sticky=tk.W)
        self.media_file_var = tk.StringVar()
        ttk.Entry(file_select_frame, textvariable=self.media_file_var, state='readonly').grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5
        )
        ttk.Button(
            file_select_frame, 
            text="选择文件", 
            command=lambda: self.select_file('media'),
            width=12
        ).grid(row=0, column=2, padx=(0, 5), pady=5)

        ttk.Label(file_select_frame, text="Sheet:").grid(row=0, column=3, padx=(10, 0), pady=5, sticky=tk.W)
        self.media_sheet_var = tk.StringVar()
        self.media_sheet_combo = ttk.Combobox(
            file_select_frame, 
            textvariable=self.media_sheet_var,
            state="disabled",
            width=20 # Adjust width
        )
        self.media_sheet_combo.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(self.media_frame, text="数据预览 (首个Sheet)", padding="10")
        # Update row index for subsequent elements
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)) 
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # 创建表格
        self.media_tree = ttk.Treeview(
            preview_frame,
            columns=("cid", "vid", "rate", "media_path", "status"),
            show="headings",
            height=10
        )
        
        # 设置列标题
        self.media_tree.heading("cid", text="CID")
        self.media_tree.heading("vid", text="VID")
        self.media_tree.heading("rate", text="码率")
        self.media_tree.heading("media_path", text="媒资路径")
        self.media_tree.heading("status", text="状态")
        
        # 设置列宽
        self.media_tree.column("cid", width=100)
        self.media_tree.column("vid", width=100)
        self.media_tree.column("rate", width=80)
        self.media_tree.column("media_path", width=250)
        self.media_tree.column("status", width=80)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.media_tree.yview)
        self.media_tree.configure(yscrollcommand=scrollbar.set)
        
        self.media_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 进度条
        progress_frame_media = ttk.Frame(self.media_frame)
        progress_frame_media.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10)) # Updated row
        progress_frame_media.grid_columnconfigure(0, weight=1)
        
        self.media_progress_var = tk.DoubleVar()
        self.media_progress_bar = ttk.Progressbar(
            progress_frame_media,
            variable=self.media_progress_var,
            maximum=100
        )
        self.media_progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.media_progress_label = ttk.Label(progress_frame_media, text="0%")
        self.media_progress_label.grid(row=0, column=1)
        
        # 修改按钮
        self.media_modify_button = ttk.Button(
            self.media_frame,
            text="开始修改",
            command=lambda: self.start_modification('media'),
            width=20,
            state='disabled'
        )
        self.media_modify_button.grid(row=3, column=0, pady=(0, 10)) # Updated row

    def create_batch_sync_page(self):
        # 配置批量下发页面的网格权重
        self.batch_sync_frame.grid_columnconfigure(0, weight=1)
        
        # 文件与Sheet选择区域
        file_select_frame = ttk.LabelFrame(self.batch_sync_frame, text="文件与Sheet选择", padding="10")
        file_select_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        # Configure columns
        file_select_frame.grid_columnconfigure(1, weight=1) # Path expands
        file_select_frame.grid_columnconfigure(4, weight=0) # Sheet combo fixed

        # --- File and Sheet Selection Row ---
        ttk.Label(file_select_frame, text="文件:").grid(row=0, column=0, padx=(5, 0), pady=5, sticky=tk.W)
        self.batch_sync_path_var = tk.StringVar()
        path_entry = ttk.Entry(
            file_select_frame, 
            textvariable=self.batch_sync_path_var, 
            state='readonly',
            font=('Consolas', 10)
        )
        path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)
        
        select_button = ttk.Button(
            file_select_frame,
            text="选择文件",
            command=self.select_batch_sync_file,
            width=12
        )
        select_button.grid(row=0, column=2, padx=(0, 5), pady=5)

        ttk.Label(file_select_frame, text="Sheet:").grid(row=0, column=3, padx=(10, 0), pady=5, sticky=tk.W)
        self.batch_sync_sheet_var = tk.StringVar()
        self.batch_sync_sheet_combo = ttk.Combobox(
            file_select_frame, 
            textvariable=self.batch_sync_sheet_var,
            state="disabled",
            width=20 # Adjust width
        )
        self.batch_sync_sheet_combo.grid(row=0, column=4, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)
        
        # 项目选择区域
        project_frame = ttk.LabelFrame(self.batch_sync_frame, text="项目选择", padding="10")
        # Update row index for subsequent elements
        project_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10)) 
        project_frame.grid_columnconfigure(1, weight=1)
        
        # 项目选择标签
        ttk.Label(project_frame, text="选择项目:").grid(row=0, column=0, padx=(5, 10), pady=5, sticky=tk.W)
        
        # 项目选择下拉框
        self.project_var = tk.StringVar()
        project_dropdown = ttk.Combobox(
            project_frame, 
            textvariable=self.project_var,
            state="readonly",
            width=30
        )
        project_dropdown['values'] = list(self.projects.keys())
        if self.projects:
            project_dropdown.current(0)
        project_dropdown.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=5)
        
        # 预览区域
        preview_frame = ttk.LabelFrame(self.batch_sync_frame, text="CID预览 (首个Sheet)", padding="10")
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)) # Updated row
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)
        
        # 预览文本框
        self.batch_sync_preview_text = scrolledtext.ScrolledText(
            preview_frame,
            width=50,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 10)
        )
        self.batch_sync_preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        self.batch_sync_preview_text.config(state='disabled')
        
        # 控制按钮区域
        control_frame = ttk.Frame(self.batch_sync_frame)
        control_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10)) # Updated row
        control_frame.grid_columnconfigure(0, weight=1)
        
        # 批量下发按钮
        self.batch_sync_button = ttk.Button(
            control_frame, 
            text="开始下发",
            command=lambda: self.start_modification('batch_sync'),
            width=20,
            state='disabled'
        )
        self.batch_sync_button.grid(row=0, column=0)
        
        # 进度条
        progress_frame_sync = ttk.Frame(self.batch_sync_frame)
        progress_frame_sync.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=5) # Updated row
        progress_frame_sync.grid_columnconfigure(0, weight=1)
        
        self.batch_sync_progress_var = tk.DoubleVar()
        self.batch_sync_progress = ttk.Progressbar(
            progress_frame_sync,
            variable=self.batch_sync_progress_var,
            maximum=100
        )
        self.batch_sync_progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.batch_sync_progress_label = ttk.Label(progress_frame_sync, text="0%", width=5)
        self.batch_sync_progress_label.grid(row=0, column=1)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            try:
                self.excel_file_handler = pd.ExcelFile(file_path)
                sheet_names = self.excel_file_handler.sheet_names
                
                if not sheet_names:
                    messagebox.showerror("错误", "Excel文件不包含任何Sheet")
                    self.excel_file_handler = None
                    self.excel_sheet_combo['values'] = []
                    self.excel_sheet_combo.set('')
                    self.excel_sheet_combo['state'] = 'disabled'
                    self.excel_modify_button['state'] = 'disabled'
                    return

                # Populate combobox
                self.excel_sheet_combo['values'] = sheet_names
                self.excel_sheet_combo.set(sheet_names[0])
                self.excel_sheet_combo['state'] = 'readonly'

                # Load first sheet for preview
                self.load_excel_preview(sheet_names[0])
                self.excel_modify_button['state'] = 'normal'
                messagebox.showinfo("成功", f"文件加载成功，找到 {len(sheet_names)} 个Sheet。请选择Sheet后开始修改。")

            except Exception as e:
                messagebox.showerror("错误", f"加载Excel文件失败: {str(e)}")
                self.excel_file_handler = None
                self.excel_path_var.set("")
                self.excel_sheet_combo['values'] = []
                self.excel_sheet_combo.set('')
                self.excel_sheet_combo['state'] = 'disabled'
                self.excel_modify_button['state'] = 'disabled'
                self.excel_preview_text.config(state='normal')
                self.excel_preview_text.delete(1.0, tk.END)
                self.excel_preview_text.config(state='disabled')

    def load_excel_preview(self, sheet_name):
        """Loads and previews data from the specified sheet for the Excel Import tab."""
        if not self.excel_file_handler:
            return
        try:
            df = self.excel_file_handler.parse(sheet_name)
            
            if df.empty or len(df.columns) < 1:
                preview_data = f"Sheet '{sheet_name}' 为空或无数据列."
                headers = "无有效列"
            else:
                # Check for CID column (assuming first column)
                if len(df.columns) > 0:
                    # Convert first column to string for preview consistency
                    df.iloc[:, 0] = df.iloc[:, 0].astype(str)
                
                # Prepare preview data
                headers = "字段列表：\n" + "\n".join([f"- {col}" for col in df.columns])
                preview_data = "数据预览（前5行）：\n"
                for i in range(min(5, len(df))):
                    row_data = [f"{col}: {df.iloc[i][col]}" for col in df.columns]
                    preview_data += f"行 {i+1}:\n" + "\n".join(row_data) + "\n\n"

            # Update preview text area
            self.excel_preview_text.config(state='normal')
            self.excel_preview_text.delete(1.0, tk.END)
            self.excel_preview_text.insert(tk.END, headers + "\n\n" + preview_data)
            self.excel_preview_text.config(state='disabled')

        except Exception as e:
            messagebox.showerror("错误", f"预览Sheet '{sheet_name}' 失败: {str(e)}")
            self.excel_preview_text.config(state='normal')
            self.excel_preview_text.delete(1.0, tk.END)
            self.excel_preview_text.insert(tk.END, f"无法加载Sheet '{sheet_name}' 的预览。")
            self.excel_preview_text.config(state='disabled')

    def select_file(self, mode):
        """Handles file selection for modes other than 'excel' and 'batch_sync' (currently just 'media')."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                if mode == 'media':
                    self.media_file_var.set(file_path)
                    self.media_file_handler = pd.ExcelFile(file_path)
                    sheet_names = self.media_file_handler.sheet_names

                    if not sheet_names:
                        messagebox.showerror("错误", "Excel文件不包含任何Sheet")
                        self.media_file_handler = None
                        self.media_sheet_combo['values'] = []
                        self.media_sheet_combo.set('')
                        self.media_sheet_combo['state'] = 'disabled'
                        self.media_modify_button['state'] = 'disabled'
                        for item in self.media_tree.get_children():
                            self.media_tree.delete(item)
                        return

                    # Populate combobox
                    self.media_sheet_combo['values'] = sheet_names
                    self.media_sheet_combo.set(sheet_names[0])
                    self.media_sheet_combo['state'] = 'readonly'

                    # Load preview from the first sheet
                    self.load_media_preview(sheet_names[0])
                    self.media_modify_button['state'] = 'normal'
                    messagebox.showinfo("成功", f"文件加载成功，找到 {len(sheet_names)} 个Sheet。请选择Sheet后开始修改。")

                else:
                    messagebox.showinfo("提示", "已选择文件: " + file_path)
                    
            except Exception as e:
                messagebox.showerror("错误", f"读取文件时出错：{str(e)}")
                if mode == 'media':
                    self.media_file_var.set("")
                    self.media_file_handler = None
                    self.media_sheet_combo['values'] = []
                    self.media_sheet_combo.set('')
                    self.media_sheet_combo['state'] = 'disabled'
                    self.media_modify_button['state'] = 'disabled'
                    for item in self.media_tree.get_children():
                        self.media_tree.delete(item)

    def load_media_preview(self, sheet_name):
        """Loads and previews data from the specified sheet for the Media tab."""
        if not self.media_file_handler:
            return
        
        # Clear existing preview
        for item in self.media_tree.get_children():
            self.media_tree.delete(item)
            
        try:
            df = self.media_file_handler.parse(sheet_name)
            
            # Check for required columns for preview (adjust as needed)
            preview_cols = ['cid', 'vid', 'rate', 'media_path', 'status']
            if not all(col in df.columns for col in preview_cols):
                missing = [col for col in preview_cols if col not in df.columns]
                messagebox.showwarning("预览警告", f"Sheet '{sheet_name}' 缺少预览所需的列: {', '.join(missing)}. 处理时仍将尝试使用可用数据。")
            else:
                # Populate treeview with data (handle potential missing columns gracefully)
                for index, row in df.iterrows():
                    values = (
                        str(row.get('cid', '')),
                        str(row.get('vid', '')),
                        str(row.get('rate', '')),
                        str(row.get('media_path', '')),
                        str(row.get('status', ''))
                    )
                    self.media_tree.insert('', 'end', values=values)
                
        except Exception as e:
            messagebox.showerror("错误", f"预览Sheet '{sheet_name}' 失败: {str(e)}")
            # Treeview is already cleared

    def select_batch_sync_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.batch_sync_path_var.set(file_path)
            try:
                self.batch_sync_file_handler = pd.ExcelFile(file_path)
                sheet_names = self.batch_sync_file_handler.sheet_names
                
                if not sheet_names:
                    messagebox.showerror("错误", "Excel文件不包含任何Sheet")
                    self.batch_sync_file_handler = None
                    self.batch_sync_sheet_combo['values'] = []
                    self.batch_sync_sheet_combo.set('')
                    self.batch_sync_sheet_combo['state'] = 'disabled'
                    self.batch_sync_button['state'] = 'disabled'
                    return

                # Populate combobox
                self.batch_sync_sheet_combo['values'] = sheet_names
                self.batch_sync_sheet_combo.set(sheet_names[0])
                self.batch_sync_sheet_combo['state'] = 'readonly'

                # Load first sheet for preview
                self.load_batch_sync_preview(sheet_names[0])
                self.batch_sync_button['state'] = 'normal'
                messagebox.showinfo("成功", f"文件加载成功，找到 {len(sheet_names)} 个Sheet。请选择Sheet后开始下发。")

            except Exception as e:
                messagebox.showerror("错误", f"加载Excel文件失败: {str(e)}")
                self.batch_sync_file_handler = None
                self.batch_sync_path_var.set("")
                self.batch_sync_sheet_combo['values'] = []
                self.batch_sync_sheet_combo.set('')
                self.batch_sync_sheet_combo['state'] = 'disabled'
                self.batch_sync_button['state'] = 'disabled'
                self.batch_sync_preview_text.config(state='normal')
                self.batch_sync_preview_text.delete(1.0, tk.END)
                self.batch_sync_preview_text.config(state='disabled')

    def load_batch_sync_preview(self, sheet_name):
        """Loads and previews CID data from the specified sheet for the Batch Sync tab."""
        if not self.batch_sync_file_handler:
            return
        try:
            df = self.batch_sync_file_handler.parse(sheet_name)
            preview_data = f"CID预览 (来自Sheet '{sheet_name}', 前10行):\n"
            
            if df.empty or len(df.columns) < 1:
                preview_data += "Sheet为空或无数据列."
                cids = []
            else:
                # Assume CIDs are in the first column
                cids = df.iloc[:, 0].astype(str).tolist()
                if not cids:
                    preview_data += "未找到CID数据 (第一列为空)."
                else:
                    for i in range(min(10, len(cids))):
                        preview_data += f"{i+1}. {str(cids[i]).strip()}\n"
            
            # Update preview text area
            self.batch_sync_preview_text.config(state='normal')
            self.batch_sync_preview_text.delete(1.0, tk.END)
            self.batch_sync_preview_text.insert(tk.END, preview_data)
            self.batch_sync_preview_text.config(state='disabled')

        except Exception as e:
            messagebox.showerror("错误", f"预览Sheet '{sheet_name}' 失败: {str(e)}")
            self.batch_sync_preview_text.config(state='normal')
            self.batch_sync_preview_text.delete(1.0, tk.END)
            self.batch_sync_preview_text.insert(tk.END, f"无法加载Sheet '{sheet_name}' 的预览。")
            self.batch_sync_preview_text.config(state='disabled')

    def process_modifications(self, mode):
        try:
            session = requests.Session()
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0',
                'Accept': 'application/json, text/plain, */*',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
                'Host': 'cms.enjoy-tv.cn',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache',
                'Cookie': '_enj="2|1:0|10:1752025383|4:_enj|36:MjFfZW5qZnVjaGFveWlAZW5qb3ktdHYuY24=|ea28d79c1f908fc2971a958e9ad63ba519d5fec2d802fb920befa6fcb4a7e25d"'
            }
            session.headers.update(headers)
            
            if mode == 'batch_sync':
                if not self.batch_sync_file_handler:
                    messagebox.showwarning("警告", "请先选择Excel文件")
                    return
                
                selected_sheet = self.batch_sync_sheet_var.get()
                if not selected_sheet:
                    messagebox.showwarning("警告", "请选择要处理的Sheet")
                    return
                    
                try:
                    df = self.batch_sync_file_handler.parse(selected_sheet)
                    if df.empty or len(df.columns) < 1:
                        messagebox.showwarning("警告", f"选择的Sheet '{selected_sheet}' 为空或无数据")
                        return
                    cids = df.iloc[:, 0].astype(str).tolist()
                except Exception as e:
                    messagebox.showerror("错误", f"读取Sheet '{selected_sheet}' 失败: {str(e)}")
                    return

                if not cids:
                    messagebox.showwarning("警告", f"Sheet '{selected_sheet}' 的第一列没有找到CID数据")
                    return
                
                selected_project = self.project_var.get()
                partner_code = self.projects.get(selected_project)
                
                if not partner_code:
                    messagebox.showwarning("警告", "请选择有效的项目")
                    return
                
                total_operations = len(cids)
                completed = 0
                success_count = 0
                error_count = 0
                
                for cid in cids:
                    if not self.is_processing:
                        break
                    
                    cid_str = str(cid).strip()
                    if not cid_str or cid_str.lower() == 'nan':
                        completed += 1
                        continue
                    
                    params = {
                        'cid': cid_str,
                        'partner_code': partner_code
                    }
                    
                    try:
                        url = "http://cms.enjoy-tv.cn/query/project/cover/batch_sync"
                        print(f"发送批量下发请求 (行 {completed+1}): {url}?{'&'.join([f'{k}={v}' for k, v in params.items()])}")
                        response = session.get(url, params=params)
                        response.encoding = 'utf-8'
                        
                        if response.status_code == 200:
                            success_count += 1
                            print(f"成功下发 CID={cid_str} 到项目={selected_project}")
                        else:
                            error_count += 1
                            print(f"下发处理出错: CID={cid_str}, 状态码={response.status_code}, 项目={selected_project}({partner_code}), 响应={response.text[:200]}")
                    except Exception as e:
                        error_count += 1
                        print(f"下发处理出错: CID={cid_str}, 错误: {str(e)}, 项目={selected_project}({partner_code})")
                    
                    completed += 1
                    progress = (completed / total_operations) * 100
                    self.batch_sync_progress_var.set(progress)
                    self.batch_sync_progress_label.config(text=f"{int(progress)}%")
                    self.root.update_idletasks()
                
                message = f"下发完成！\nSheet: {selected_sheet}\n项目: {selected_project}\n成功：{success_count}条\n失败：{error_count}条"
                messagebox.showinfo("完成", message)
                
            elif mode == 'media':
                if not self.media_file_handler:
                    messagebox.showwarning("警告", "请先选择Excel文件")
                    return

                selected_sheet = self.media_sheet_var.get()
                if not selected_sheet:
                    messagebox.showwarning("警告", "请选择要处理的Sheet")
                    return
                    
                try:
                    media_df = self.media_file_handler.parse(selected_sheet)
                    # Validate required columns (adjust as needed)
                    required_cols = ['cid', 'vid', 'rate', 'media_path', 'status']
                    if not all(col in media_df.columns for col in required_cols):
                        missing = [col for col in required_cols if col not in media_df.columns]
                        messagebox.showerror("错误", f"选择的Sheet '{selected_sheet}' 缺少列: {', '.join(missing)}")
                        return
                except Exception as e:
                    messagebox.showerror("错误", f"读取Sheet '{selected_sheet}' 失败: {str(e)}")
                    return

                total_operations = len(media_df)
                if total_operations == 0:
                    messagebox.showwarning("警告", f"Sheet '{selected_sheet}' 中没有数据")
                    return
                    
                completed = 0
                success_count = 0
                error_count = 0
                
                for index, row in media_df.iterrows():
                    if not self.is_processing:
                        break
                    
                    try:
                        # Ensure data types are correct (especially for API params)
                        cid_val = str(row['cid']).strip()
                        vid_val = str(row['vid']).strip()
                        rate_val = str(row['rate']).strip()
                        media_path_val = str(row['media_path']).strip()
                        status_val = str(row['status']).strip()

                        if not cid_val or not vid_val or not rate_val or cid_val.lower() == 'nan' or vid_val.lower() == 'nan':
                            print(f"跳过无效行 (Sheet: {selected_sheet}, 行: {index+2}): CID或VID为空")
                            error_count += 1
                            continue

                        # First modify media path
                        media_params = {
                            'cid': cid_val,
                            'vid': vid_val,
                            'rate': rate_val,
                            'media_path': media_path_val
                        }
                        # 过滤掉空值和NaN值
                        filtered_media_params = {k:v for k,v in media_params.items() if v and str(v).lower() != 'nan'}
                        media_url = "http://cms.enjoy-tv.cn/api/media/edit"
                        print(f"发送介质路径修改请求 (行 {index+2}): {media_url}?{'&'.join([f'{k}={v}' for k, v in filtered_media_params.items()])}")
                        media_response = session.get(media_url, params=filtered_media_params)
                        media_response.encoding = 'utf-8'
                        
                        # Then modify status
                        status_params = {
                            'cid': cid_val,
                            'vid': vid_val,
                            'rate': rate_val,
                            'status': status_val
                        }
                        # 过滤掉空值和NaN值
                        filtered_status_params = {k:v for k,v in status_params.items() if v and str(v).lower() != 'nan'}
                        status_url = "http://cms.enjoy-tv.cn/api/media/edit"
                        print(f"发送介质状态修改请求 (行 {index+2}): {status_url}?{'&'.join([f'{k}={v}' for k, v in filtered_status_params.items()])}")
                        status_response = session.get(status_url, params=filtered_status_params)
                        status_response.encoding = 'utf-8'
                        
                        # Check combined response success (adjust logic if needed)
                        if media_response.status_code == 200 and status_response.status_code == 200:
                            success_count += 1
                        else:
                            error_count += 1
                            print(f"介质处理出错: CID={cid_val}, VID={vid_val}, "
                                  f"路径状态码={media_response.status_code}, "
                                  f"状态修改状态码={status_response.status_code}")
                            
                    except Exception as e:
                        error_count += 1
                        print(f"介质处理异常: Sheet '{selected_sheet}', 行 {index+2}, CID={row.get('cid', 'N/A')}, VID={row.get('vid', 'N/A')}, 错误: {str(e)}")
                    
                    completed += 1
                    progress = (completed / total_operations) * 100
                    self.media_progress_var.set(progress)
                    self.media_progress_label.config(text=f"{int(progress)}%")
                    self.root.update_idletasks()
                
                message = f"介质修改完成！\nSheet: {selected_sheet}\n成功：{success_count}条\n失败：{error_count}条"
                messagebox.showinfo("完成", message)

            elif mode == 'excel':
                if not self.excel_file_handler:
                    messagebox.showwarning("警告", "请先导入Excel文件")
                    return

                selected_sheet = self.excel_sheet_var.get()
                if not selected_sheet:
                    messagebox.showwarning("警告", "请选择要处理的Sheet")
                    return

                try:
                    excel_df = self.excel_file_handler.parse(selected_sheet)
                    if excel_df.empty or len(excel_df.columns) < 1:
                        messagebox.showwarning("警告", f"选择的Sheet '{selected_sheet}' 为空或无数据")
                        return
                    cids = excel_df.iloc[:, 0].astype(str).tolist()
                    fields_data_dict = excel_df.iloc[:, 1:].to_dict(orient='list')
                    for field in fields_data_dict:
                        fields_data_dict[field] = [str(val) for val in fields_data_dict[field]]

                except Exception as e:
                    messagebox.showerror("错误", f"读取Sheet '{selected_sheet}' 失败: {str(e)}")
                    return

                if not cids:
                    messagebox.showwarning("警告", f"Sheet '{selected_sheet}' 的第一列没有找到CID数据")
                    return
                
                total_operations = len(cids)
                completed = 0
                success_count = 0
                error_count = 0
                
                for i, cid in enumerate(cids):
                    if not self.is_processing:
                        break
                    
                    cid_str = str(cid).strip()
                    if not cid_str or cid_str.lower() == 'nan':
                        print(f"跳过无效CID (行 {i+2}): {cid_str}")
                        completed += 1
                        continue

                    # 构造请求参数 - CID作为必需参数
                    params = {'cid': cid_str}
                    
                    # 添加其他字段参数
                    for field, values in fields_data_dict.items():
                        if i < len(values):
                            value_str = str(values[i]).strip()
                            # 包含所有非空值，包括空字符串
                            if value_str and value_str.lower() != 'nan':
                                params[field] = value_str
                    
                    # 只要有CID就发送请求（即使没有其他字段也要尝试）
                    try:
                        url = "http://cms.enjoy-tv.cn/api/cover/master_edit"
                        print(f"发送请求 (行 {i+2}): {url}?{'&'.join([f'{k}={v}' for k, v in params.items()])}")
                        response = session.get(url, params=params)
                        response.encoding = 'utf-8'
                        
                        if response.status_code == 200:
                            success_count += 1
                            print(f"成功处理 CID={cid_str}")
                        else:
                            error_count += 1
                            print(f"Excel导入处理出错: CID={cid_str}, 状态码={response.status_code}, 响应={response.text[:200]}")
                    except Exception as e:
                        error_count += 1
                        print(f"Excel导入处理异常: CID={cid_str}, 错误: {str(e)}")
                    
                    completed += 1
                    progress = (completed / total_operations) * 100
                    self.excel_progress_var.set(progress)
                    self.excel_progress_label.config(text=f"{int(progress)}%")
                    self.root.update_idletasks()
                
                message = f"Excel导入处理完成！\nSheet: {selected_sheet}\n成功：{success_count}条\n失败：{error_count}条"
                messagebox.showinfo("完成", message)
                
        except Exception as e:
            messagebox.showerror("错误", f"发生意外错误: {str(e)}")
            
        finally:
            self.is_processing = False
            if mode == 'excel':
                self.excel_modify_button['state'] = 'normal' if self.excel_file_handler else 'disabled'
                self.excel_progress_var.set(0)
                self.excel_progress_label.config(text="0%")
            elif mode == 'media':
                self.media_modify_button['state'] = 'normal' if self.media_file_handler else 'disabled'
                self.media_progress_var.set(0)
                self.media_progress_label.config(text="0%")
            elif mode == 'batch_sync':
                self.batch_sync_button['state'] = 'normal' if self.batch_sync_file_handler else 'disabled'
                self.batch_sync_progress_var.set(0)
                self.batch_sync_progress_label.config(text="0%")

    def start_modification(self, mode):
        if not self.is_processing:
            self.is_processing = True
            if mode == 'excel':
                self.excel_modify_button['state'] = 'disabled'
            elif mode == 'media':
                self.media_modify_button['state'] = 'disabled'
            elif mode == 'batch_sync':
                self.batch_sync_button['state'] = 'disabled'
            
            threading.Thread(target=self.process_modifications, args=(mode,), daemon=True).start()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = BatchAPIProcessor()
    app.run()
