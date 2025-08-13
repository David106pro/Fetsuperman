import pandas as pd
from difflib import SequenceMatcher
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import sys

def create_gui():
    # 创建主窗口
    root = tk.Tk()
    root.title("影片路径匹配工具")
    root.geometry("1200x800")
    
    # 设置窗口样式
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 10))
    style.configure('TLabel', font=('Arial', 10))
    
    # 创建主框架
    main_frame = ttk.Frame(root, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 创建左侧控制面板
    left_frame = ttk.Frame(main_frame)
    left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
    
    # 创建右侧预览面板
    right_frame = ttk.Frame(main_frame)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
    
    # 左侧控制面板内容
    # Excel文件路径输入
    ttk.Label(left_frame, text="Excel文件路径:").grid(column=0, row=0, sticky=tk.W, pady=5)
    excel_path_var = tk.StringVar()
    excel_entry = ttk.Entry(left_frame, width=50, textvariable=excel_path_var)
    excel_entry.grid(column=0, row=1, sticky=(tk.W, tk.E), pady=5)
    
    # Sheet选择
    ttk.Label(left_frame, text="选择Sheet:").grid(column=0, row=2, sticky=tk.W, pady=5)
    sheet_var = tk.StringVar()
    sheet_combo = ttk.Combobox(left_frame, textvariable=sheet_var, width=47, state="readonly")
    sheet_combo.grid(column=0, row=3, sticky=(tk.W, tk.E), pady=5)
    
    # 列选择
    ttk.Label(left_frame, text="标题列名:").grid(column=0, row=4, sticky=tk.W, pady=5)
    column_var = tk.StringVar()
    column_combo = ttk.Combobox(left_frame, textvariable=column_var, width=47, state="readonly")
    column_combo.grid(column=0, row=5, sticky=(tk.W, tk.E), pady=5)
    
    # 4M/6M/8M码率勾选框
    bitrate_frame = ttk.Frame(left_frame)
    bitrate_frame.grid(column=0, row=6, sticky=(tk.W, tk.E), pady=5)
    bitrate_4m_var = tk.BooleanVar(value=True)
    bitrate_6m_var = tk.BooleanVar(value=False)
    bitrate_8m_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(bitrate_frame, text="4M", variable=bitrate_4m_var).pack(side=tk.LEFT, padx=5)
    ttk.Checkbutton(bitrate_frame, text="6M", variable=bitrate_6m_var).pack(side=tk.LEFT, padx=5)
    ttk.Checkbutton(bitrate_frame, text="8M", variable=bitrate_8m_var).pack(side=tk.LEFT, padx=5)

    # 4M路径输入和按钮同一行
    m4_row = ttk.Frame(left_frame)
    m4_row.grid(column=0, row=7, sticky=(tk.W, tk.E), pady=5)
    m4_path_var = tk.StringVar()
    m4_entry = ttk.Entry(m4_row, width=38, textvariable=m4_path_var)
    m4_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Button(m4_row, text="浏览", command=lambda: browse_m4()).pack(side=tk.LEFT, padx=3)

    # 6M路径输入和按钮同一行
    m6_row = ttk.Frame(left_frame)
    m6_row.grid(column=0, row=8, sticky=(tk.W, tk.E), pady=5)
    m6_path_var = tk.StringVar()
    m6_entry = ttk.Entry(m6_row, width=38, textvariable=m6_path_var)
    m6_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Button(m6_row, text="浏览", command=lambda: browse_m6()).pack(side=tk.LEFT, padx=3)

    # 8M路径输入和按钮同一行
    m8_row = ttk.Frame(left_frame)
    m8_row.grid(column=0, row=9, sticky=(tk.W, tk.E), pady=5)
    m8_path_var = tk.StringVar()
    m8_entry = ttk.Entry(m8_row, width=38, textvariable=m8_path_var)
    m8_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Button(m8_row, text="浏览", command=lambda: browse_m8()).pack(side=tk.LEFT, padx=3)

    # 其余控件行号顺延
    ttk.Button(left_frame, text="浏览Excel文件", command=lambda: browse_excel()).grid(column=0, row=10, pady=5, sticky=tk.W)
    execute_btn = ttk.Button(left_frame, text="开始处理", command=lambda: execute_process())
    execute_btn.grid(column=0, row=11, pady=20, sticky=tk.W)
    
    # 右侧预览面板内容
    preview_frame = ttk.LabelFrame(right_frame, text="表格预览", padding=5)
    preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
    
    # 创建表格预览
    preview_tree = ttk.Treeview(preview_frame, show="headings", height=15)
    preview_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    
    # 添加滚动条
    v_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=preview_tree.yview)
    v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    preview_tree.configure(yscrollcommand=v_scrollbar.set)
    
    h_scrollbar = ttk.Scrollbar(right_frame, orient="horizontal", command=preview_tree.xview)
    h_scrollbar.pack(fill=tk.X)
    preview_tree.configure(xscrollcommand=h_scrollbar.set)
    
    # 状态信息显示
    status_frame = ttk.LabelFrame(right_frame, text="处理状态", padding=5)
    status_frame.pack(fill=tk.X, pady=(10, 0))
    status_text = tk.Text(status_frame, height=8, wrap=tk.WORD)
    status_text.pack(fill=tk.BOTH, expand=True)
    
    # 状态文本滚动条
    status_scrollbar = ttk.Scrollbar(status_text, orient="vertical", command=status_text.yview)
    status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    status_text.configure(yscrollcommand=status_scrollbar.set)
    
    # 全局变量存储当前数据
    current_excel_data = {}
    
    # 用于更新状态信息的函数
    def update_status(message):
        status_text.insert(tk.END, message + "\n")
        status_text.see(tk.END)
        root.update_idletasks()
    
    # 更新表格预览
    def update_preview():
        try:
            # 清空现有内容
            for item in preview_tree.get_children():
                preview_tree.delete(item)
            
            selected_sheet = sheet_var.get()
            if not selected_sheet or selected_sheet not in current_excel_data:
                return
            
            df = current_excel_data[selected_sheet]
            
            # 设置列标题
            columns = list(df.columns)
            preview_tree["columns"] = columns
            
            for col in columns:
                preview_tree.heading(col, text=col)
                preview_tree.column(col, width=100, minwidth=50)
            
            # 添加数据（只显示前100行）
            for i, row in df.head(100).iterrows():
                values = [str(row[col]) if pd.notna(row[col]) else "" for col in columns]
                preview_tree.insert("", tk.END, values=values)
            
            # ====== 优化：标题列名优先级自动选择 ======
            # 优先级：title > 介质名称 > 片名 > 第一个包含"名称"的表头
            default_col = None
            for key in ['title', '介质名称', '片名']:
                if key in columns:
                    default_col = key
                    break
            if not default_col:
                for col in columns:
                    if '名称' in col:
                        default_col = col
                        break
            if default_col:
                column_var.set(default_col)
            else:
                column_var.set(columns[0] if columns else '')
            
            # 更新列选择下拉框
            column_combo['values'] = columns
            update_status(f"已加载Sheet: {selected_sheet}, 共{len(df)}行数据")
                
        except Exception as e:
            update_status(f"预览更新失败: {str(e)}")
    
    # 加载Excel文件
    def load_excel_file(file_path):
        try:
            update_status(f"正在加载Excel文件: {file_path}")
            
            # 读取所有sheet
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            current_excel_data.clear()
            for sheet_name in sheet_names:
                current_excel_data[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # 更新sheet选择下拉框
            sheet_combo['values'] = sheet_names
            if sheet_names:
                sheet_var.set(sheet_names[0])
                update_preview()
            
            update_status(f"Excel文件加载完成，共{len(sheet_names)}个Sheet")
            
        except Exception as e:
            update_status(f"Excel文件加载失败: {str(e)}")
            messagebox.showerror("错误", f"Excel文件加载失败: {str(e)}")
    
    # 浏览按钮点击事件
    def browse_excel():
        filename = filedialog.askopenfilename(
            title="选择Excel文件", 
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if filename:
            excel_path_var.set(filename)
            load_excel_file(filename)
    
    def browse_m4():
        filename = filedialog.askopenfilename(
            title="选择4M路径文件", 
            filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            m4_path_var.set(filename)
            update_status(f"已选择4M路径文件: {filename}")
    
    def browse_m6():
        filename = filedialog.askopenfilename(
            title="选择6M路径文件", 
            filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            m6_path_var.set(filename)
            update_status(f"已选择6M路径文件: {filename}")
    
    def browse_m8():
        filename = filedialog.askopenfilename(
            title="选择8M路径文件", 
            filetypes=[("Text Files", "*.txt")]
        )
        if filename:
            m8_path_var.set(filename)
            update_status(f"已选择8M路径文件: {filename}")
    
    # 执行处理的函数
    def execute_process():
        excel_path = excel_path_var.get().strip()
        selected_sheet = sheet_var.get().strip()
        selected_column = column_var.get().strip()
        m4_path = m4_path_var.get().strip()
        m6_path = m6_path_var.get().strip()
        m8_path = m8_path_var.get().strip()
        use_4m = bitrate_4m_var.get()
        use_6m = bitrate_6m_var.get()
        use_8m = bitrate_8m_var.get()

        if not excel_path:
            messagebox.showerror("错误", "请选择Excel文件")
            return
        if not selected_sheet:
            messagebox.showerror("错误", "请选择Sheet")
            return
        if not selected_column:
            messagebox.showerror("错误", "请选择标题列")
            return
        if use_4m and not m4_path:
            messagebox.showerror("错误", "请选择4M路径文件")
            return
        if use_6m and not m6_path:
            messagebox.showerror("错误", "请选择6M路径文件")
            return
        if use_8m and not m8_path:
            messagebox.showerror("错误", "请选择8M路径文件")
            return
        if not (use_4m or use_6m or use_8m):
            messagebox.showerror("错误", "请至少勾选一个码率")
            return
        try:
            update_status("开始处理文件...")
            output_path = process_files(
                excel_path, selected_sheet, selected_column,
                m4_path if use_4m else None,
                m6_path if use_6m else None,
                m8_path if use_8m else None,
                use_4m, use_6m, use_8m,
                update_status
            )
            update_status("处理完成!")
            if messagebox.askyesno("处理完成", f"处理完成，结果已保存至:\n{output_path}\n\n是否立即打开文件？"):
                open_file(output_path)
        except Exception as e:
            update_status(f"处理出错: {str(e)}")
            messagebox.showerror("错误", f"处理出错: {str(e)}")
    
    # Sheet选择变化事件
    def on_sheet_change(event):
        update_preview()
    
    # 绑定事件
    sheet_combo.bind('<<ComboboxSelected>>', on_sheet_change)
    
    # ====== 自动查找桌面最近的xlsx文件，作为默认Excel路径 ======
    def auto_set_excel_path():
        # 使用中文注释说明：自动查找桌面最近的xlsx文件
        desktop_dir = r"C:\Users\fucha\Desktop"
        if not os.path.exists(desktop_dir):
            return
        xlsx_files = [os.path.join(desktop_dir, f) for f in os.listdir(desktop_dir) if f.lower().endswith('.xlsx')]
        if not xlsx_files:
            return
        xlsx_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        recent_xlsx = xlsx_files[0]
        excel_path_var.set(recent_xlsx)
        update_status(f"自动选择最近的Excel文件: {recent_xlsx}")
    # 启动时自动设置Excel路径
    auto_set_excel_path()

    # ====== 自动扫描txt目录，按时间排序，自动填入4M/8M路径 ======
    def auto_set_txt_paths():
        # 使用中文注释说明：自动扫描目录，按时间排序，自动填入4M/8M路径
        txt_dir = r"D:\物料转接\批次txt"
        if not os.path.exists(txt_dir):
            return
        # 获取所有txt文件及其修改时间
        txt_files = [os.path.join(txt_dir, f) for f in os.listdir(txt_dir) if f.lower().endswith('.txt')]
        if not txt_files:
            return
        # 按修改时间降序排序
        txt_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        # 取最近的两个文件
        recent_files = txt_files[:2]
        # 遍历文件名，自动填入4M/8M路径
        for f in recent_files:
            fname = os.path.basename(f).lower()
            if '4m' in fname:
                m4_path_var.set(f)
                update_status(f"自动选择4M路径文件: {f}")
            elif '8m' in fname:
                m8_path_var.set(f)
                update_status(f"自动选择8M路径文件: {f}")
    # 启动时自动设置
    auto_set_txt_paths()

    # ====== 6M路径默认值保持不变 ======
    m6_path_var.set(r"D:\物料转接\批次txt\250114_6m.txt")
    
    # 设置列权重
    left_frame.columnconfigure(0, weight=1)
    
    # 如果有默认Excel文件，尝试加载
    if os.path.exists(excel_path_var.get()):
        load_excel_file(excel_path_var.get())
    
    # 启动主循环
    root.mainloop()

# 自动打开文件的函数
def open_file(file_path):
    try:
        if sys.platform == "win32":
            os.startfile(file_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", file_path])
        else:
            subprocess.run(["xdg-open", file_path])
    except Exception as e:
        print(f"无法打开文件: {e}")

# 修改process_files函数
def process_files(excel_path, sheet_name, column_name, m4_path, m6_path, m8_path, use_4m, use_6m, use_8m, update_status=None):
    if update_status:
        update_status(f"正在读取Excel文件: {excel_path}, Sheet: {sheet_name}")
    # 读取所有sheet
    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    if sheet_name not in all_sheets:
        raise ValueError(f"Sheet '{sheet_name}' 不存在于Excel文件中")
    df = all_sheets[sheet_name]
    if column_name not in df.columns:
        raise ValueError(f"指定的列 '{column_name}' 不存在于Sheet '{sheet_name}' 中")
    # 读取路径文件
    m4_paths = m6_paths = m8_paths = []
    if use_4m:
        if update_status:
            update_status(f"正在读取4M路径文件: {m4_path}")
        with open(m4_path, 'r', encoding='utf-8') as f:
            m4_paths = [line.strip() for line in f.readlines() if line.strip()]
        if update_status:
            update_status(f"4M路径文件包含 {len(m4_paths)} 个路径")
    if use_6m:
        if update_status:
            update_status(f"正在读取6M路径文件: {m6_path}")
        with open(m6_path, 'r', encoding='utf-8') as f:
            m6_paths = [line.strip() for line in f.readlines() if line.strip()]
        if update_status:
            update_status(f"6M路径文件包含 {len(m6_paths)} 个路径")
    if use_8m:
        if update_status:
            update_status(f"正在读取8M路径文件: {m8_path}")
        with open(m8_path, 'r', encoding='utf-8') as f:
            m8_paths = [line.strip() for line in f.readlines() if line.strip()]
        if update_status:
            update_status(f"8M路径文件包含 {len(m8_paths)} 个路径")
    # ========== 优化：智能识别表头 =============
    # 使用中文注释说明：查找4M/6M/8M路径相关表头
    def find_col(cols, m):
        # 优先找 *M路径，其次 *m_path
        m_upper = f"{m}M路径"
        m_lower = f"{m}m_path"
        col_M = None
        if m_upper in cols:
            col_M = m_upper
        elif m_lower in cols:
            col_M = m_lower
        return col_M
    # 记录各码率路径和匹配度列名
    col_4m = find_col(df.columns, 4) if use_4m else None
    col_6m = find_col(df.columns, 6) if use_6m else None
    col_8m = find_col(df.columns, 8) if use_8m else None
    # 匹配度列名
    def find_ratio_col(cols, m):
        # 优先找 *M匹配度，其次 *m_ratio
        m_upper = f"{m}M匹配度"
        m_lower = f"{m}m_ratio"
        col_ratio = None
        if m_upper in cols:
            col_ratio = m_upper
        elif m_lower in cols:
            col_ratio = m_lower
        return col_ratio
    col_4m_ratio = find_ratio_col(df.columns, 4) if use_4m else None
    col_6m_ratio = find_ratio_col(df.columns, 6) if use_6m else None
    col_8m_ratio = find_ratio_col(df.columns, 8) if use_8m else None
    # 没有则新建
    if use_4m and not col_4m:
        col_4m = '4M路径'
        df[col_4m] = ''
    if use_4m and not col_4m_ratio:
        col_4m_ratio = '4M匹配度'
        df[col_4m_ratio] = ''
    if use_6m and not col_6m:
        col_6m = '6M路径'
        df[col_6m] = ''
    if use_6m and not col_6m_ratio:
        col_6m_ratio = '6M匹配度'
        df[col_6m_ratio] = ''
    if use_8m and not col_8m:
        col_8m = '8M路径'
        df[col_8m] = ''
    if use_8m and not col_8m_ratio:
        col_8m_ratio = '8M匹配度'
        df[col_8m_ratio] = ''
    # ====== 强制将相关列转为字符串类型，避免pandas dtype警告 ======
    for col in [col_4m, col_4m_ratio, col_6m, col_6m_ratio, col_8m, col_8m_ratio]:
        if col and col in df.columns:
            df[col] = df[col].astype(str)
    title_path_map = {}
    if update_status:
        update_status(f"开始匹配视频标题与路径...")
    total_rows = len(df)
    matched_4m = matched_6m = matched_8m = 0
    
    # ====== 优化：支持多语言版本处理 ======
    rows_to_insert = []  # 存储需要插入的新行信息
    
    for idx, row in df.iterrows():
        title = str(row[column_name])
        if update_status and idx % max(1, total_rows // 10) == 0:
            update_status(f"正在处理: {idx+1}/{total_rows} ({(idx+1)/total_rows*100:.1f}%)")
        
        # 使用缓存避免重复计算
        if title in title_path_map:
            cache = title_path_map[title]
        else:
            cache = {}
            # 查找所有匹配的路径（支持多语言版本）
            if use_4m:
                all_4m_matches = find_all_matches(title, m4_paths)
                is_multi_4m, language_versions_4m = detect_multi_language_versions(all_4m_matches)
                if is_multi_4m:
                    cache['4M'] = {'is_multi': True, 'versions': language_versions_4m}
                    if update_status:
                        update_status(f"调试：标题 '{title}' 检测到4M多语言版本，共{len(language_versions_4m)}个")
                else:
                    # 使用原来的单一匹配逻辑
                    m4_match, m4_ratio = find_best_match(title, m4_paths)
                    cache['4M'] = {'is_multi': False, 'path': m4_match, 'ratio': m4_ratio}
            
            if use_6m:
                all_6m_matches = find_all_matches(title, m6_paths)
                is_multi_6m, language_versions_6m = detect_multi_language_versions(all_6m_matches)
                if is_multi_6m:
                    cache['6M'] = {'is_multi': True, 'versions': language_versions_6m}
                    if update_status:
                        update_status(f"调试：标题 '{title}' 检测到6M多语言版本，共{len(language_versions_6m)}个")
                else:
                    m6_match, m6_ratio = find_best_match(title, m6_paths)
                    cache['6M'] = {'is_multi': False, 'path': m6_match, 'ratio': m6_ratio}
            
            if use_8m:
                all_8m_matches = find_all_matches(title, m8_paths)
                is_multi_8m, language_versions_8m = detect_multi_language_versions(all_8m_matches)
                if is_multi_8m:
                    cache['8M'] = {'is_multi': True, 'versions': language_versions_8m}
                    if update_status:
                        update_status(f"调试：标题 '{title}' 检测到8M多语言版本，共{len(language_versions_8m)}个")
                else:
                    m8_match, m8_ratio = find_best_match(title, m8_paths)
                    cache['8M'] = {'is_multi': False, 'path': m8_match, 'ratio': m8_ratio}
            
            title_path_map[title] = cache
        
        # 处理匹配结果
        multi_versions_count = 0  # 记录最大语言版本数量
        
        # 检查是否有多语言版本，确定需要插入的行数
        if use_4m and cache['4M']['is_multi']:
            multi_versions_count = max(multi_versions_count, len(cache['4M']['versions']))
        if use_6m and cache['6M']['is_multi']:
            multi_versions_count = max(multi_versions_count, len(cache['6M']['versions']))
        if use_8m and cache['8M']['is_multi']:
            multi_versions_count = max(multi_versions_count, len(cache['8M']['versions']))
        
        if multi_versions_count > 1:
            # 有多语言版本，需要插入额外行
            if update_status:
                update_status(f"第{idx+1}行检测到多语言版本: {title}, 共{multi_versions_count}个版本")
            
            # 为每个语言版本准备行数据
            for version_idx in range(multi_versions_count):
                row_data = {
                    'original_idx': idx,
                    'version_idx': version_idx,
                    'is_first_row': version_idx == 0,
                    'title': title
                }
                
                # 填充各码率的路径和匹配度
                if use_4m:
                    if cache['4M']['is_multi'] and version_idx < len(cache['4M']['versions']):
                        version_info = cache['4M']['versions'][version_idx]
                        row_data['4m_path'] = version_info['path']
                        row_data['4m_ratio'] = f"{version_info['ratio']:.2f}"
                        row_data['4m_language'] = version_info['language']
                    elif not cache['4M']['is_multi'] and version_idx == 0:
                        row_data['4m_path'] = cache['4M']['path']
                        row_data['4m_ratio'] = f"{cache['4M']['ratio']:.2f}" if cache['4M']['ratio'] else ''
                
                if use_6m:
                    if cache['6M']['is_multi'] and version_idx < len(cache['6M']['versions']):
                        version_info = cache['6M']['versions'][version_idx]
                        row_data['6m_path'] = version_info['path']
                        row_data['6m_ratio'] = f"{version_info['ratio']:.2f}"
                        row_data['6m_language'] = version_info['language']
                    elif not cache['6M']['is_multi'] and version_idx == 0:
                        row_data['6m_path'] = cache['6M']['path']
                        row_data['6m_ratio'] = f"{cache['6M']['ratio']:.2f}" if cache['6M']['ratio'] else ''
                
                if use_8m:
                    if cache['8M']['is_multi'] and version_idx < len(cache['8M']['versions']):
                        version_info = cache['8M']['versions'][version_idx]
                        row_data['8m_path'] = version_info['path']
                        row_data['8m_ratio'] = f"{version_info['ratio']:.2f}"
                        row_data['8m_language'] = version_info['language']
                    elif not cache['8M']['is_multi'] and version_idx == 0:
                        row_data['8m_path'] = cache['8M']['path']
                        row_data['8m_ratio'] = f"{cache['8M']['ratio']:.2f}" if cache['8M']['ratio'] else ''
                
                rows_to_insert.append(row_data)
        else:
            # 单一版本，按原来的逻辑处理
            if use_4m and not cache['4M']['is_multi']:
                m4_match = cache['4M']['path']
                m4_ratio = cache['4M']['ratio']
                if m4_match:
                    df.at[idx, col_4m] = m4_match
                    df.at[idx, col_4m_ratio] = f"{m4_ratio:.2f}" if m4_ratio else ''
                    matched_4m += 1
                    if update_status:
                        update_status(f"第{idx+1}行4M匹配: {title} => {m4_match} 匹配度: {m4_ratio}")
            
            if use_6m and not cache['6M']['is_multi']:
                m6_match = cache['6M']['path']
                m6_ratio = cache['6M']['ratio']
                if m6_match:
                    df.at[idx, col_6m] = m6_match
                    df.at[idx, col_6m_ratio] = f"{m6_ratio:.2f}" if m6_ratio else ''
                    matched_6m += 1
                    if update_status:
                        update_status(f"第{idx+1}行6M匹配: {title} => {m6_match} 匹配度: {m6_ratio}")
            
            if use_8m and not cache['8M']['is_multi']:
                m8_match = cache['8M']['path']
                m8_ratio = cache['8M']['ratio']
                if m8_match:
                    df.at[idx, col_8m] = m8_match
                    df.at[idx, col_8m_ratio] = f"{m8_ratio:.2f}" if m8_ratio else ''
                    matched_8m += 1
                    if update_status:
                        update_status(f"第{idx+1}行8M匹配: {title} => {m8_match} 匹配度: {m8_ratio}")
    
    # ====== 处理多语言版本：插入新行 ======
    if rows_to_insert:
        if update_status:
            update_status(f"正在处理多语言版本，需要插入{len(rows_to_insert)}行数据...")
        
        # 按原始索引倒序排列，避免插入时索引偏移
        rows_to_insert.sort(key=lambda x: (x['original_idx'], x['version_idx']), reverse=True)
        
        # 分组处理每个原始行的多语言版本
        current_group = []
        current_original_idx = None
        
        for row_data in rows_to_insert:
            if current_original_idx != row_data['original_idx']:
                # 处理上一组
                if current_group:
                    insert_multi_language_rows(df, current_group, col_4m, col_4m_ratio, col_6m, col_6m_ratio, col_8m, col_8m_ratio)
                    # 更新匹配统计
                    for group_row in current_group:
                        if group_row.get('4m_path'):
                            matched_4m += 1
                        if group_row.get('6m_path'):
                            matched_6m += 1
                        if group_row.get('8m_path'):
                            matched_8m += 1
                
                # 开始新组
                current_group = [row_data]
                current_original_idx = row_data['original_idx']
            else:
                current_group.append(row_data)
        
        # 处理最后一组
        if current_group:
            insert_multi_language_rows(df, current_group, col_4m, col_4m_ratio, col_6m, col_6m_ratio, col_8m, col_8m_ratio)
            for group_row in current_group:
                if group_row.get('4m_path'):
                    matched_4m += 1
                if group_row.get('6m_path'):
                    matched_6m += 1
                if group_row.get('8m_path'):
                    matched_8m += 1
    
    # 用ExcelWriter覆盖当前sheet，保留其他sheet
    output_path = excel_path
    all_sheets[sheet_name] = df
    if update_status:
        stat = []
        if use_4m:
            stat.append(f"4M成功匹配 {matched_4m}/{total_rows}")
        if use_6m:
            stat.append(f"6M成功匹配 {matched_6m}/{total_rows}")
        if use_8m:
            stat.append(f"8M成功匹配 {matched_8m}/{total_rows}")
        update_status(f"匹配统计: {'，'.join(stat)}")
        update_status(f"保存结果到: {output_path}")
    with pd.ExcelWriter(output_path, mode='w', engine='openpyxl') as writer:
        for sname, sdf in all_sheets.items():
            sdf.to_excel(writer, sheet_name=sname, index=False)
    if update_status:
        update_status(f"处理完成，结果已保存至: {output_path}")
    else:
        print(f"处理完成，结果已保存至: {output_path}")
    return output_path

# 优化匹配函数
def find_best_match(title, path_list, threshold=0.5):
    """
    查找最佳匹配的路径
    
    参数:
        title: 要匹配的标题
        path_list: 路径列表
        threshold: 相似度阈值，低于此值不视为匹配
        
    返回:
        (最佳匹配的路径或None, 匹配度)
    """
    best_match = None
    max_ratio = 0
    
    # 清理标题，只保留数字、字母和中文字符
    clean_title = re.sub(r'[^\w\u4e00-\u9fff]', '', title.lower())
    
    for path in path_list:
        # 从路径中提取文件名，处理类似 /data1/hs4/250507/B115/4M/电视剧/3B的恋人_10 的格式
        file_name = path.split('/')[-1]  # 获取最后一部分
        
        # 去掉可能的扩展名
        if '.' in file_name:
            file_name = file_name.rsplit('.', 1)[0]
        
        # 清理文件名，只保留数字、字母和中文字符
        clean_file = re.sub(r'[^\w\u4e00-\u9fff]', '', file_name.lower())
        
        # 计算相似度
        ratio = SequenceMatcher(None, clean_title, clean_file).ratio()
        
        # 如果标题包含在文件名中或文件名包含在标题中，提高相似度
        if clean_title in clean_file or clean_file in clean_title:
            ratio = max(ratio, 0.8)
        
        # 更新最佳匹配
        if ratio > max_ratio and ratio >= threshold:
            max_ratio = ratio
            best_match = path
    
    return best_match, max_ratio if best_match else None

# 新增：查找所有匹配的路径（用于多语言版本处理）
def find_all_matches(title, path_list, threshold=0.3):
    """
    查找所有匹配的路径
    
    参数:
        title: 要匹配的标题
        path_list: 路径列表
        threshold: 相似度阈值，低于此值不视为匹配（降低到0.3）
        
    返回:
        [(路径, 匹配度), ...] 按匹配度降序排列
    """
    matches = []
    
    # 清理标题，只保留数字、字母和中文字符
    clean_title = re.sub(r'[^\w\u4e00-\u9fff]', '', title.lower())
    
    print(f"调试：查找标题 '{title}' 的所有匹配，清理后: '{clean_title}'")
    
    for path in path_list:
        # 从路径中提取文件名
        file_name = path.split('/')[-1]
        
        # 去掉可能的扩展名
        if '.' in file_name:
            file_name = file_name.rsplit('.', 1)[0]
        
        # 清理文件名，只保留数字、字母和中文字符
        clean_file = re.sub(r'[^\w\u4e00-\u9fff]', '', file_name.lower())
        
        # 计算相似度
        ratio = SequenceMatcher(None, clean_title, clean_file).ratio()
        
        # 如果标题包含在文件名中或文件名包含在标题中，提高相似度
        if clean_title in clean_file or clean_file in clean_title:
            ratio = max(ratio, 0.8)
        
        # 收集所有超过阈值的匹配
        if ratio >= threshold:
            matches.append((path, ratio))
            print(f"调试：匹配到文件 '{file_name}' 相似度: {ratio:.3f}")
    
    # 按匹配度降序排列
    matches.sort(key=lambda x: x[1], reverse=True)
    print(f"调试：总共找到{len(matches)}个匹配文件")
    return matches

# 新增：检测是否为多语言版本
def detect_multi_language_versions(matches):
    """
    检测匹配结果是否为多语言版本
    
    参数:
        matches: [(路径, 匹配度), ...] 匹配结果列表
        
    返回:
        bool: 是否为多语言版本
        list: 如果是多语言版本，返回按语言版本分组的路径列表
    """
    if len(matches) < 2:
        return False, []
    
    # 提取所有文件名（不含扩展名）
    filenames = []
    for path, ratio in matches:
        filename = path.split('/')[-1]
        if '.' in filename:
            filename = filename.rsplit('.', 1)[0]
        filenames.append(filename)
    
    # 调试输出：显示所有匹配的文件名
    print(f"调试：检测多语言版本，找到{len(filenames)}个匹配文件名: {filenames}")
    
    # 查找共同前缀
    if not filenames:
        return False, []
    
    # 优化：使用更智能的共同前缀查找
    # 先尝试找到最短的文件名作为基准
    base_name = min(filenames, key=len)
    common_prefix = ""
    
    # 逐字符检查共同前缀
    for i in range(len(base_name)):
        char = base_name[i]
        if all(i < len(name) and name[i] == char for name in filenames):
            common_prefix += char
        else:
            break
    
    print(f"调试：找到共同前缀: '{common_prefix}' (长度: {len(common_prefix)})")
    
    # 降低共同前缀长度要求，从3改为2
    if len(common_prefix) < 2:
        return False, []
    
    # 检查是否所有文件名都以共同前缀开头，并且后面有下划线分隔的语言版本
    language_versions = []
    for i, filename in enumerate(filenames):
        if filename.startswith(common_prefix):
            # 提取语言版本部分
            suffix = filename[len(common_prefix):]
            print(f"调试：文件名 '{filename}' 的后缀: '{suffix}'")
            
            # 优化：不仅检查下划线开头，也检查直接的语言标识
            if suffix.startswith('_') or len(suffix) > 0:
                language_part = suffix[1:] if suffix.startswith('_') else suffix
                language_versions.append({
                    'path': matches[i][0],
                    'ratio': matches[i][1],
                    'language': language_part if language_part else '默认版本'
                })
                print(f"调试：识别语言版本: '{language_part}'")
    
    print(f"调试：识别到{len(language_versions)}个语言版本")
    
    # 如果有多个语言版本，则认为是多语言版本
    if len(language_versions) >= 2:
        print(f"调试：确认为多语言版本，共{len(language_versions)}个版本")
        return True, language_versions
    
    return False, []

# 新增：插入多语言版本行
def insert_multi_language_rows(df, rows_group, col_4m, col_4m_ratio, col_6m, col_6m_ratio, col_8m, col_8m_ratio):
    """
    在DataFrame中插入多语言版本行
    
    参数:
        df: DataFrame对象
        rows_group: 同一原始行的多语言版本数据列表
        col_4m, col_4m_ratio, col_6m, col_6m_ratio, col_8m, col_8m_ratio: 各码率的列名
    """
    if not rows_group:
        return
    
    # 按version_idx排序，确保第一行是原始行
    rows_group.sort(key=lambda x: x['version_idx'])
    
    original_idx = rows_group[0]['original_idx']
    
    # 处理第一行（原始行）- 更新现有行
    first_row = rows_group[0]
    if first_row.get('4m_path'):
        df.at[original_idx, col_4m] = first_row['4m_path']
        if col_4m_ratio:
            df.at[original_idx, col_4m_ratio] = first_row['4m_ratio']
    if first_row.get('6m_path'):
        df.at[original_idx, col_6m] = first_row['6m_path']
        if col_6m_ratio:
            df.at[original_idx, col_6m_ratio] = first_row['6m_ratio']
    if first_row.get('8m_path'):
        df.at[original_idx, col_8m] = first_row['8m_path']
        if col_8m_ratio:
            df.at[original_idx, col_8m_ratio] = first_row['8m_ratio']
    
    # 处理后续行（新增行）- 在DataFrame末尾添加
    for row_data in rows_group[1:]:
        # 复制原始行的所有数据
        new_row = df.iloc[original_idx].copy()
        
        # 清空路径相关列，只保留基础信息
        if col_4m:
            new_row[col_4m] = ''
        if col_4m_ratio:
            new_row[col_4m_ratio] = ''
        if col_6m:
            new_row[col_6m] = ''
        if col_6m_ratio:
            new_row[col_6m_ratio] = ''
        if col_8m:
            new_row[col_8m] = ''
        if col_8m_ratio:
            new_row[col_8m_ratio] = ''
        
        # 填充当前语言版本的路径
        if row_data.get('4m_path'):
            new_row[col_4m] = row_data['4m_path']
            if col_4m_ratio:
                new_row[col_4m_ratio] = row_data['4m_ratio']
        if row_data.get('6m_path'):
            new_row[col_6m] = row_data['6m_path']
            if col_6m_ratio:
                new_row[col_6m_ratio] = row_data['6m_ratio']
        if row_data.get('8m_path'):
            new_row[col_8m] = row_data['8m_path']
            if col_8m_ratio:
                new_row[col_8m_ratio] = row_data['8m_ratio']
        
        # 在DataFrame末尾添加新行
        df.loc[len(df)] = new_row

# 启动程序
if __name__ == "__main__":
    create_gui() 