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
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # 创建选项卡控件
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # 创建两个选项卡页面
        self.manual_frame = ttk.Frame(self.notebook, padding="10")
        self.excel_frame = ttk.Frame(self.notebook, padding="10")
        self.media_frame = ttk.Frame(self.notebook, padding="10")

        self.notebook.add(self.manual_frame, text="手动修改")
        self.notebook.add(self.excel_frame, text="Excel导入")
        self.notebook.add(self.media_frame, text="介质修改")

        # 手动修改页面
        self.create_manual_page()

        # Excel导入页面
        self.create_excel_page()

        # 媒资修改页面
        self.create_media_page()

        self.is_processing = False
        self.excel_data = None

    def create_manual_page(self):
        # 配置手动修改页面的网格权重
        self.manual_frame.grid_columnconfigure(0, weight=1)

        # CID输入区域
        cid_frame = ttk.LabelFrame(self.manual_frame, text="CID列表", padding="10")
        cid_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        cid_frame.grid_columnconfigure(0, weight=1)

        # CID输入框
        self.cid_text = scrolledtext.ScrolledText(
            cid_frame,
            width=50,  # 调整宽度
            height=8,  # 调整高度
            wrap=tk.WORD,
            font=('Consolas', 10)  # 设置等宽字体
        )
        self.cid_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)

        # 批量字段修改区域
        fields_frame = ttk.LabelFrame(self.manual_frame, text="批量字段修改", padding="10")
        fields_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        fields_frame.grid_columnconfigure(1, weight=1)
        fields_frame.grid_columnconfigure(3, weight=1)

        # 字段输入框
        self.field_vars = {}
        fields = ['is_effective', 'ott', 'iptv', 'sole', 'm4_status', 'm8_status', 'director', 'actor']
        fields_per_column = (len(fields) + 1) // 2

        for i, field in enumerate(fields):
            row = i % fields_per_column
            col = (i // fields_per_column) * 2

            # 字段标签
            ttk.Label(fields_frame, text=field, width=12).grid(
                row=row, column=col, padx=(10, 5), pady=5, sticky=tk.E
            )

            # 字段输入框
            var = tk.StringVar()
            self.field_vars[field] = var
            ttk.Entry(fields_frame, textvariable=var, width=30).grid(
                row=row, column=col + 1, padx=(0, 20), pady=5, sticky=(tk.W, tk.E)
            )

        # hs_id修改区域
        hsid_frame = ttk.LabelFrame(self.manual_frame, text="hs_id修改（最多10组）", padding="10")
        hsid_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        hsid_frame.grid_columnconfigure(0, weight=1)

        self.hsid_entries = []
        self.hsid_container = ttk.Frame(hsid_frame)
        self.hsid_container.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.hsid_container.grid_columnconfigure(1, weight=1)
        self.hsid_container.grid_columnconfigure(3, weight=1)

        # 添加按钮 - 保存为实例变量
        self.add_hsid_button = ttk.Button(hsid_frame, text="添加一组", command=self.add_hsid_entry)
        self.add_hsid_button.grid(row=1, column=0, pady=(5, 0))

        # 修改按钮 - 保存为实例变量
        self.manual_modify_button = ttk.Button(
            self.manual_frame,
            text="开始修改",
            command=lambda: self.start_modification('manual'),
            width=20
        )
        self.manual_modify_button.grid(row=3, column=0, pady=(0, 10))

        # 进度条区域
        self.create_progress_bar(self.manual_frame, 4)

    def create_excel_page(self):
        # 配置Excel页面的网格权重
        self.excel_frame.grid_columnconfigure(0, weight=1)

        # Excel导入区域
        excel_import_frame = ttk.LabelFrame(self.excel_frame, text="Excel导入", padding="10")
        excel_import_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        excel_import_frame.grid_columnconfigure(0, weight=1)

        # 文件选择区域
        file_frame = ttk.Frame(excel_import_frame)
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        file_frame.grid_columnconfigure(0, weight=1)

        # 文件路径显示
        self.excel_path_var = tk.StringVar()
        path_entry = ttk.Entry(
            file_frame,
            textvariable=self.excel_path_var,
            state='readonly',
            font=('Consolas', 10)
        )
        path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(5, 10))

        # 选择文件按钮
        select_button = ttk.Button(
            file_frame,
            text="选择文件",
            command=self.select_excel_file,
            width=15
        )
        select_button.grid(row=0, column=1, padx=(0, 5))

        # 预览区域
        preview_frame = ttk.LabelFrame(self.excel_frame, text="数据预览", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)

        # 预览文本框
        self.preview_text = scrolledtext.ScrolledText(
            preview_frame,
            width=50,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 10)
        )
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        self.preview_text.config(state='disabled')

        # 控制按钮区域
        control_frame = ttk.Frame(self.excel_frame)
        control_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
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

        # 进度条
        self.create_progress_bar(self.excel_frame, 3)

    def create_media_page(self):
        # 配置媒资修改页面的网格权重
        self.media_frame.grid_columnconfigure(0, weight=1)

        # Excel文件选择区域
        file_frame = ttk.Frame(self.media_frame)
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel文件：").grid(row=0, column=0, padx=(0, 5))
        self.media_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.media_file_var, state='readonly').grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(
            file_frame,
            text="选择文件",
            command=lambda: self.select_file('media')
        ).grid(row=0, column=2)

        # 预览区域
        preview_frame = ttk.LabelFrame(self.media_frame, text="数据预览", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(0, weight=1)

        # 创建表格
        self.media_tree = ttk.Treeview(
            preview_frame,
            columns=("cid", "vid", "rate", "media_path"),
            show="headings",
            height=10
        )

        # 设置列标题
        self.media_tree.heading("cid", text="CID")
        self.media_tree.heading("vid", text="VID")
        self.media_tree.heading("rate", text="码率")
        self.media_tree.heading("media_path", text="媒资路径")

        # 设置列宽
        self.media_tree.column("cid", width=100)
        self.media_tree.column("vid", width=100)
        self.media_tree.column("rate", width=100)
        self.media_tree.column("media_path", width=300)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.media_tree.yview)
        self.media_tree.configure(yscrollcommand=scrollbar.set)

        self.media_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # 进度条
        progress_frame = ttk.Frame(self.media_frame)
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.grid_columnconfigure(0, weight=1)

        self.media_progress_var = tk.DoubleVar()
        self.media_progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.media_progress_var,
            maximum=100
        )
        self.media_progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        self.media_progress_label = ttk.Label(progress_frame, text="0%")
        self.media_progress_label.grid(row=0, column=1)

        # 修改按钮
        self.media_modify_button = ttk.Button(
            self.media_frame,
            text="开始修改",
            command=lambda: self.start_modification('media'),
            width=20
        )
        self.media_modify_button.grid(row=3, column=0, pady=(0, 10))

    def create_cid_input(self, parent):
        # CID输入框标签
        cid_label = ttk.Label(parent, text="请输入CID列表（每行一个，最多100个）：")
        cid_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        # CID输入框
        self.cid_text = scrolledtext.ScrolledText(
            parent,
            width=70,
            height=10,
            wrap=tk.WORD  # 使用单词换行而不是字符换行
        )
        self.cid_text.grid(row=2, column=0, columnspan=2, pady=(0, 10))

    def create_batch_fields(self, parent):
        # 批量字段修改区域
        fields_frame = ttk.LabelFrame(parent, text="批量字段修改", padding="5")
        fields_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 字段输入框
        self.field_vars = {}
        fields = ['is_effective', 'ott', 'iptv', 'sole', 'm4_status', 'm8_status', 'director', 'actor']

        # 计算每列显示的字段数
        fields_per_column = (len(fields) + 1) // 2

        for i, field in enumerate(fields):
            # 计算行和列位置
            row = i % fields_per_column
            col = (i // fields_per_column) * 2

            # 添加标签
            ttk.Label(fields_frame, text=field).grid(
                row=row,
                column=col,
                padx=5,
                pady=2,
                sticky=tk.E
            )

            # 添加输入框
            var = tk.StringVar()
            self.field_vars[field] = var
            ttk.Entry(fields_frame, textvariable=var).grid(
                row=row,
                column=col + 1,
                sticky=(tk.W, tk.E),
                padx=5,
                pady=2
            )

        # 配置列的权重
        fields_frame.columnconfigure(1, weight=1)
        fields_frame.columnconfigure(3, weight=1)

    def create_hsid_section(self, parent):
        # hs_id修改区域
        self.hsid_frame = ttk.LabelFrame(parent, text="hs_id修改（最多10组）", padding="5")
        self.hsid_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        # 创建内部框架
        self.entries_frame = ttk.Frame(self.hsid_frame)
        self.entries_frame.pack(fill="x", expand=True)

        self.hsid_entries = []

        # 按钮框架
        button_frame = ttk.Frame(self.hsid_frame)
        button_frame.pack(fill="x", pady=5)

        # 添加按钮
        self.add_hsid_button = ttk.Button(
            button_frame,
            text="添加一组",
            command=self.add_hsid_entry
        )
        self.add_hsid_button.pack(side=tk.LEFT, padx=5)

    def add_hsid_entry(self):
        if len(self.hsid_entries) >= 10:
            messagebox.showwarning("警告", "最多只能添加10组hs_id修改")
            return

        row = len(self.hsid_entries)

        # 创建一行的容器
        entry_frame = ttk.Frame(self.hsid_container)
        entry_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=2)
        entry_frame.grid_columnconfigure(1, weight=1)
        entry_frame.grid_columnconfigure(3, weight=1)

        # CID输入框
        ttk.Label(entry_frame, text="CID:", width=8).grid(row=0, column=0, padx=(5, 0))
        cid_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=cid_var, width=30).grid(
            row=0, column=1, padx=5, sticky=(tk.W, tk.E)
        )

        # hs_id输入框
        ttk.Label(entry_frame, text="hs_id:", width=8).grid(row=0, column=2, padx=(10, 0))
        hsid_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=hsid_var, width=30).grid(
            row=0, column=3, padx=5, sticky=(tk.W, tk.E)
        )

        # 删除按钮
        ttk.Button(
            entry_frame,
            text="删除",
            command=lambda f=entry_frame, idx=row: self.delete_hsid_entry(f, idx),
            width=8
        ).grid(row=0, column=4, padx=5)

        self.hsid_entries.append({
            'frame': entry_frame,
            'cid': cid_var,
            'hsid': hsid_var
        })

    def delete_hsid_entry(self, frame, idx):
        frame.destroy()
        if 0 <= idx < len(self.hsid_entries):  # 添加索引检查
            self.hsid_entries.pop(idx)
        # 删除后总是启用添加按钮
        self.add_hsid_button['state'] = 'normal'

    def create_progress_bar(self, parent, row):
        progress_frame = ttk.Frame(parent)
        progress_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=5)
        progress_frame.grid_columnconfigure(0, weight=1)

        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        # 百分比标签
        self.progress_label = ttk.Label(progress_frame, text="0%", width=5)
        self.progress_label.grid(row=0, column=1)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.load_excel_data(file_path)

    def load_excel_data(self, file_path):
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)

            # 检查是否至少有一列数据
            if df.empty or len(df.columns) < 1:
                messagebox.showerror("错误", "Excel文件为空或格式不正确")
                return

            # 获取第一列作为CID
            cids = df.iloc[:, 0].astype(str).tolist()

            # 获取除第一列外的所有列作为修改字段
            self.excel_data = {
                'cids': cids,
                'fields': {}
            }

            # 保存字段数据
            for field in df.columns[1:]:
                self.excel_data['fields'][field] = df[field].astype(str).tolist()

            # 更新预览
            self.preview_text.config(state='normal')
            self.preview_text.delete(1.0, tk.END)

            # 添加表头预览
            headers = "字段列表：\n" + "\n".join([f"- {col}" for col in df.columns])
            self.preview_text.insert(tk.END, headers + "\n\n")

            # 添加数据预览
            preview_data = "数据预览（前5行）：\n"
            for i in range(min(5, len(df))):
                row_data = [f"{col}: {df.iloc[i][col]}" for col in df.columns]
                preview_data += f"行 {i + 1}:\n" + "\n".join(row_data) + "\n\n"

            self.preview_text.insert(tk.END, preview_data)
            self.preview_text.config(state='disabled')

            # 启用修改按钮
            self.excel_modify_button['state'] = 'normal'

            messagebox.showinfo("成功", "Excel数据导入成功")

        except Exception as e:
            messagebox.showerror("错误", f"导入Excel文件失败: {str(e)}")
            self.excel_data = None
            self.excel_modify_button['state'] = 'disabled'
            self.preview_text.config(state='normal')
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.config(state='disabled')

    def process_modifications(self, mode):
        try:
            session = requests.Session()
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0',
                'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5',
                'Accept-Encoding': 'gzip, deflate',
                'Connection': 'keep-alive',
                'Host': 'cms.enjoy-tv.cn',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache',
                'Cookie': '_enj="2|1:0|10:1737683592|4:_enj|36:MjFfZW5qZnVjaGFveWlAZW5qb3ktdHYuY24=|baca4c16df8c9ef8653c49e525641853a3206497eacd72d607e2290308e57211"'
            }
            session.headers.update(headers)

            if mode == 'media':
                if self.excel_data is None:
                    messagebox.showwarning("警告", "请先选择Excel文件")
                    return

                total_operations = len(self.excel_data)
                if total_operations == 0:
                    messagebox.showwarning("警告", "Excel文件中没有数据")
                    return

                completed = 0
                success_count = 0
                error_count = 0

                # 遍历Excel数据进行修改
                for _, row in self.excel_data.iterrows():
                    if not self.is_processing:
                        break

                    try:
                        # 首先修改媒资路径
                        media_params = {
                            'cid': str(row['cid']),
                            'vid': str(row['vid']),
                            'rate': str(row['rate']),
                            'media_path': str(row['media_path'])
                        }

                        # 发送媒资路径修改请求
                        media_url = "http://cms.enjoy-tv.cn/api/media/edit"
                        media_response = session.get(media_url, params=media_params)
                        media_response.encoding = 'utf-8'

                        # 然后修改状态
                        status_params = {
                            'cid': str(row['cid']),
                            'vid': str(row['vid']),
                            'rate': str(row['rate']),
                            'status': str(row['status'])
                        }

                        # 发送状态修改请求
                        status_url = "http://cms.enjoy-tv.cn/api/media/edit"
                        status_response = session.get(status_url, params=status_params)
                        status_response.encoding = 'utf-8'

                        # 检查响应
                        if media_response.status_code == 200 and status_response.status_code == 200:
                            success_count += 1
                        else:
                            error_count += 1
                            print(f"处理出错: CID={row['cid']}, VID={row['vid']}, "
                                  f"媒资状态码={media_response.status_code}, "
                                  f"状态修改状态码={status_response.status_code}")

                    except Exception as e:
                        error_count += 1
                        print(f"处理出错: CID={row['cid']}, VID={row['vid']}, 错误: {str(e)}")

                    # 更新进度
                    completed += 1
                    progress = (completed / total_operations) * 100
                    self.media_progress_var.set(progress)
                    self.media_progress_label.config(text=f"{int(progress)}%")
                    self.root.update_idletasks()

                # 显示处理结果
                message = f"处理完成！\n成功：{success_count}条\n失败：{error_count}条"
                messagebox.showinfo("完成", message)

            elif mode == 'excel':
                if not self.excel_data:
                    messagebox.showwarning("警告", "请先导入Excel文件")
                    return

                cids = self.excel_data['cids']
                fields_data = self.excel_data['fields']

                if not cids:
                    messagebox.showwarning("警告", "Excel文件中没有CID数据")
                    return

                total_operations = len(cids)
                completed = 0

                # 处理每一行数据
                for i, cid in enumerate(cids):
                    if not self.is_processing:
                        break

                    params = {'cid': cid}

                    # 添加该行的所有字段值
                    for field, values in fields_data.items():
                        if i < len(values) and values[i] != 'nan' and values[i].strip():
                            params[field] = values[i].strip()

                    # 如果有需要修改的字段
                    if len(params) > 1:  # params中至少包含cid和一个其他字段
                        try:
                            url = f"http://cms.enjoy-tv.cn/api/cover/master_edit"
                            response = session.get(url, params=params)
                            response.encoding = 'utf-8'
                        except Exception:
                            pass

                    completed += 1
                    progress = (completed / total_operations) * 100
                    self.progress_var.set(progress)
                    self.progress_label.config(text=f"{int(progress)}%")
                    self.root.update_idletasks()

                messagebox.showinfo("完成", "Excel数据处理完成！")

            else:
                # 手动模式处理逻辑
                cids = [cid.strip() for cid in self.cid_text.get("1.0", tk.END).splitlines() if cid.strip()]
                has_field_modifications = any(var.get().strip() for var in self.field_vars.values())

                # 检查CID列表
                if has_field_modifications:
                    if not cids:
                        messagebox.showwarning("警告", "请输入至少一个CID")
                        return

                    if len(cids) > 100:
                        messagebox.showwarning("警告", "CID数量不能超过100个")
                        return

                # 处理批量字段修改
                if has_field_modifications and cids:
                    total_operations = len(cids)
                    completed = 0

                    for cid in cids:
                        if not self.is_processing:
                            break

                        params = {'cid': cid}
                        for field, var in self.field_vars.items():
                            value = var.get().strip()
                            if value:
                                params[field] = value

                        if len(params) > 1:
                            try:
                                url = f"http://cms.enjoy-tv.cn/api/cover/master_edit"
                                response = session.get(url, params=params)
                                response.encoding = 'utf-8'
                            except Exception:
                                pass

                        completed += 1
                        progress = (completed / total_operations) * 100
                        self.progress_var.set(progress)
                        self.progress_label.config(text=f"{int(progress)}%")
                        self.root.update_idletasks()

                # 处理hs_id修改
                valid_hsid_entries = [(entry['cid'].get().strip(), entry['hsid'].get().strip())
                                      for entry in self.hsid_entries
                                      if entry['cid'].get().strip() and entry['hsid'].get().strip()]

                if valid_hsid_entries:
                    for cid, hsid in valid_hsid_entries:
                        try:
                            url = f"http://cms.enjoy-tv.cn/api/cover/master_edit"
                            response = session.get(url, params={'cid': cid, 'hs_id': hsid})
                            response.encoding = 'utf-8'
                        except Exception:
                            pass

                # 检查是否有任何修改
                if not (has_field_modifications or valid_hsid_entries):
                    messagebox.showwarning("警告", "没有需要处理的修改")
                else:
                    messagebox.showinfo("完成", "所有修改处理完成！")

        except Exception as e:
            messagebox.showerror("错误", str(e))

        finally:
            self.is_processing = False
            if mode == 'manual':
                self.manual_modify_button['state'] = 'normal'
            elif mode == 'excel':
                self.excel_modify_button['state'] = 'normal'
            elif mode == 'media':
                self.media_modify_button['state'] = 'normal'
                self.media_progress_var.set(0)
                self.media_progress_label.config(text="0%")

    def start_modification(self, mode='manual'):
        if not self.is_processing:
            self.is_processing = True
            if mode == 'manual':
                self.manual_modify_button['state'] = 'disabled'
            else:
                self.excel_modify_button['state'] = 'disabled'
            threading.Thread(target=self.process_modifications, args=(mode,), daemon=True).start()

    def select_file(self, mode):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                if mode == 'media':
                    self.media_file_var.set(file_path)
                    self.excel_data = pd.read_excel(file_path)

                    # 清空现有数据
                    for item in self.media_tree.get_children():
                        self.media_tree.delete(item)

                    # 添加新数据
                    for _, row in self.excel_data.iterrows():
                        self.media_tree.insert('', 'end', values=(
                            row['cid'],
                            row['vid'],
                            row['rate'],
                            row['media_path']
                        ))
                else:
                    # 这部分代码有问题，因为没有定义self.file_var和self.tree
                    # 应该根据实际情况修改或删除这部分代码
                    messagebox.showinfo("提示", "已选择文件: " + file_path)
                    # 如果需要处理其他模式的文件选择，请根据实际需求实现

            except Exception as e:
                messagebox.showerror("错误", f"读取文件时出错：{str(e)}")
                if mode == 'media':
                    self.media_file_var.set("")
                # 其他模式的错误处理也需要相应修改
                self.excel_data = None

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = BatchAPIProcessor()
    app.run()
