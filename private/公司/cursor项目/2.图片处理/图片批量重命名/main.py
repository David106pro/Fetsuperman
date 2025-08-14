#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
图片批量重命名工具
功能：根据Excel表格中的对应关系，批量重命名图片文件，并按横竖图分类存储
"""

import os
import sys
import shutil
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image
import subprocess
from pathlib import Path
import glob

class ImageRenamer:
    def __init__(self):
        """初始化图片重命名工具"""
        self.root = tk.Tk()
        self.root.title("图片批量重命名工具")
        self.root.geometry("800x600")
        
        # 数据存储
        self.excel_data = None
        self.excel_path = ""
        self.image_folder_path = ""
        self.selected_sheet = ""
        self.original_column = ""
        self.rename_column = ""
        
        # 默认图片文件夹路径
        self.default_image_folder = r"D:\物料转接\海报\重庆图片\重庆海报"
        
        # 创建UI界面
        self.create_ui()
        
        # 设置默认图片文件夹路径
        self.set_default_image_folder()
        
        # 自动选择桌面最后修改的Excel文件
        self.auto_select_excel()
    
    def create_ui(self):
        """创建用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置行列权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Excel文件选择
        ttk.Label(main_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.excel_path_var = tk.StringVar()
        excel_entry = ttk.Entry(main_frame, textvariable=self.excel_path_var, width=50)
        excel_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_excel_file).grid(row=0, column=2, padx=5, pady=5)
        
        # Sheet选择
        ttk.Label(main_frame, text="Sheet选择:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(main_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.load_sheet_data)
        
        # 图片文件夹选择
        ttk.Label(main_frame, text="图片文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.image_folder_var = tk.StringVar()
        folder_entry = ttk.Entry(main_frame, textvariable=self.image_folder_var, width=50)
        folder_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="浏览", command=self.select_image_folder).grid(row=2, column=2, padx=5, pady=5)
        
        # 列选择框架
        column_frame = ttk.LabelFrame(main_frame, text="列选择", padding="5")
        column_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        column_frame.columnconfigure(1, weight=1)
        column_frame.columnconfigure(3, weight=1)
        
        ttk.Label(column_frame, text="原名列:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.original_col_var = tk.StringVar()
        self.original_col_combo = ttk.Combobox(column_frame, textvariable=self.original_col_var, state="readonly")
        self.original_col_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(column_frame, text="重命名列:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.rename_col_var = tk.StringVar()
        self.rename_col_combo = ttk.Combobox(column_frame, textvariable=self.rename_col_var, state="readonly")
        self.rename_col_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=5)
        
        # 数据预览
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览", padding="5")
        preview_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview用于数据预览
        self.tree = ttk.Treeview(preview_frame)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=tree_scroll_y.set)
        
        tree_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=tree_scroll_x.set)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # 开始按钮
        start_button = ttk.Button(main_frame, text="开始处理", command=self.start_processing)
        start_button.grid(row=7, column=0, columnspan=3, pady=10)
        
        # 配置主框架行权重
        main_frame.rowconfigure(4, weight=1)
    
    def set_default_image_folder(self):
        """设置默认图片文件夹路径"""
        if os.path.exists(self.default_image_folder):
            self.image_folder_var.set(self.default_image_folder)
            self.image_folder_path = self.default_image_folder
            self.status_var.set(f"已设置默认图片文件夹: {self.default_image_folder}")
        else:
            print(f"默认图片文件夹不存在: {self.default_image_folder}")
    
    def process_rename_id(self, rename_id):
        """处理重命名列ID长度，确保不超过14位"""
        try:
            # 移除所有空格和特殊字符
            clean_id = str(rename_id).strip()
            
            # 如果ID长度超过14位，取前14位
            if len(clean_id) > 14:
                processed_id = clean_id[:14]
                print(f"ID长度超过14位，已截取前14位: {clean_id} -> {processed_id}")
                return processed_id
            else:
                return clean_id
        except Exception as e:
            print(f"处理重命名ID时出错: {str(e)}")
            return str(rename_id)
    
    def auto_select_excel(self):
        """自动选择桌面最后修改的Excel文件"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            excel_files = []
            
            # 查找Excel文件
            for ext in ['*.xlsx', '*.xls']:
                excel_files.extend(glob.glob(os.path.join(desktop_path, ext)))
            
            if excel_files:
                # 按修改时间排序，选择最新的
                latest_file = max(excel_files, key=os.path.getmtime)
                self.excel_path_var.set(latest_file)
                self.load_excel_file(latest_file)
        except Exception as e:
            print(f"自动选择Excel文件失败: {str(e)}")
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.load_excel_file(file_path)
    
    def load_excel_file(self, file_path):
        """加载Excel文件"""
        try:
            # 先设置excel_path，这样load_sheet_data才能使用
            self.excel_path = file_path
            
            # 读取Excel文件的所有sheet名称
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # 更新sheet下拉框
            self.sheet_combo['values'] = sheet_names
            if sheet_names:
                self.sheet_combo.set(sheet_names[0])
                self.load_sheet_data()
            
            self.status_var.set("Excel文件加载成功")
        except Exception as e:
            messagebox.showerror("错误", f"加载Excel文件失败: {str(e)}")
            self.status_var.set("Excel文件加载失败")
    
    def load_sheet_data(self, event=None):
        """加载选定sheet的数据"""
        try:
            if not hasattr(self, 'excel_path') or not self.excel_path or not self.sheet_var.get():
                return
            
            # 读取选定的sheet
            self.excel_data = pd.read_excel(self.excel_path, sheet_name=self.sheet_var.get())
            
            # 更新列选择下拉框
            columns = list(self.excel_data.columns)
            self.original_col_combo['values'] = columns
            self.rename_col_combo['values'] = columns
            
            # 清空之前的选择
            self.original_col_var.set("")
            self.rename_col_var.set("")
            
            # 更新数据预览
            self.update_preview()
            
            self.status_var.set(f"已加载sheet: {self.sheet_var.get()}")
        except Exception as e:
            messagebox.showerror("错误", f"加载sheet数据失败: {str(e)}")
            self.status_var.set("加载sheet数据失败")
    
    def update_preview(self):
        """更新数据预览"""
        try:
            # 清空现有数据
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            if self.excel_data is None:
                return
            
            # 设置列
            columns = list(self.excel_data.columns)
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'
            
            # 设置列标题
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)
            
            # 插入数据（只显示前100行）
            for index, row in self.excel_data.head(100).iterrows():
                values = [str(row[col]) for col in columns]
                self.tree.insert('', 'end', values=values)
            
        except Exception as e:
            messagebox.showerror("错误", f"更新预览失败: {str(e)}")
    
    def select_image_folder(self):
        """选择图片文件夹"""
        folder_path = filedialog.askdirectory(title="选择图片文件夹")
        if folder_path:
            self.image_folder_var.set(folder_path)
            self.image_folder_path = folder_path
            self.status_var.set("图片文件夹选择成功")
    
    def remove_duplicate_markers_and_deduplicate(self, landscape_folder, portrait_folder):
        """删除文件名中的(2)标识并进行去重处理"""
        try:
            for folder in [landscape_folder, portrait_folder]:
                folder_type = "横图" if folder.endswith("横图") else "竖图"
                print(f"正在处理{folder_type}文件夹去重...")
                
                # 获取文件夹中的所有文件
                files = os.listdir(folder)
                
                # 创建文件映射字典：基础名称 -> 文件列表
                file_groups = {}
                
                for filename in files:
                    if not filename.lower().endswith(('.jpg', '.jpeg')):
                        continue
                    
                    file_path = os.path.join(folder, filename)
                    
                    # 提取基础名称（去掉(2)标识）
                    base_name, ext = os.path.splitext(filename)
                    # 去掉 " (2)" 标识
                    clean_base_name = base_name.replace(' (2)', '')
                    
                    if clean_base_name not in file_groups:
                        file_groups[clean_base_name] = []
                    
                    file_groups[clean_base_name].append({
                        'original_path': file_path,
                        'filename': filename,
                        'base_name': base_name,
                        'clean_name': clean_base_name,
                        'extension': ext,
                        'has_marker': ' (2)' in base_name
                    })
                
                # 处理每个文件组
                for clean_name, file_list in file_groups.items():
                    if len(file_list) == 1:
                        # 只有一个文件，直接重命名（去掉(2)标识）
                        file_info = file_list[0]
                        if file_info['has_marker']:
                            old_path = file_info['original_path']
                            new_filename = f"{clean_name}{file_info['extension']}"
                            new_path = os.path.join(folder, new_filename)
                            
                            # 如果目标文件不存在，直接重命名
                            if not os.path.exists(new_path):
                                os.rename(old_path, new_path)
                                print(f"重命名: {file_info['filename']} -> {new_filename}")
                    else:
                        # 多个文件，需要去重处理
                        print(f"发现重复文件组: {clean_name} (共{len(file_list)}个文件)")
                        
                        # 选择保留的文件（优先保留没有(2)标识的文件）
                        files_without_marker = [f for f in file_list if not f['has_marker']]
                        files_with_marker = [f for f in file_list if f['has_marker']]
                        
                        if files_without_marker:
                            # 如果有没有(2)标识的文件，保留第一个
                            keep_file = files_without_marker[0]
                            remove_files = files_without_marker[1:] + files_with_marker
                        else:
                            # 如果都有(2)标识，保留第一个并重命名
                            keep_file = files_with_marker[0]
                            remove_files = files_with_marker[1:]
                        
                        # 确保保留的文件名是正确的（没有(2)标识）
                        if keep_file['has_marker']:
                            old_path = keep_file['original_path']
                            new_filename = f"{clean_name}{keep_file['extension']}"
                            new_path = os.path.join(folder, new_filename)
                            
                            if not os.path.exists(new_path) or os.path.samefile(old_path, new_path):
                                if not os.path.samefile(old_path, new_path):
                                    os.rename(old_path, new_path)
                                    print(f"重命名保留文件: {keep_file['filename']} -> {new_filename}")
                        
                        # 删除重复文件
                        for remove_file in remove_files:
                            try:
                                os.remove(remove_file['original_path'])
                                print(f"删除重复文件: {remove_file['filename']}")
                            except Exception as e:
                                print(f"删除文件失败 {remove_file['filename']}: {str(e)}")
                
                print(f"{folder_type}文件夹去重完成")
                
        except Exception as e:
            print(f"去重处理失败: {str(e)}")
            messagebox.showerror("警告", f"文件去重处理失败: {str(e)}")
    
    def is_landscape_image(self, image_path):
        """判断图片是否为横图"""
        try:
            with Image.open(image_path) as img:
                width, height = img.size
                return width > height
        except Exception:
            return False
    
    def find_image_files(self, base_name):
        """根据基础名称查找图片文件（包括带(2)的配对文件）"""
        found_files = []
        
        # 查找可能的文件扩展名
        extensions = ['.jpg', '.jpeg', '.JPG', '.JPEG']
        
        for ext in extensions:
            # 查找主文件
            main_file = os.path.join(self.image_folder_path, f"{base_name}{ext}")
            if os.path.exists(main_file):
                found_files.append(main_file)
            
            # 查找配对文件（带(2)）
            pair_file = os.path.join(self.image_folder_path, f"{base_name} (2){ext}")
            if os.path.exists(pair_file):
                found_files.append(pair_file)
        
        return found_files
    
    def start_processing(self):
        """开始处理图片重命名"""
        # 验证输入
        if not self.excel_data is not None:
            messagebox.showerror("错误", "请先选择并加载Excel文件")
            return
        
        if not self.image_folder_path:
            messagebox.showerror("错误", "请选择图片文件夹")
            return
        
        if not self.original_col_var.get() or not self.rename_col_var.get():
            messagebox.showerror("错误", "请选择原名列和重命名列")
            return
        
        try:
            # 创建输出文件夹
            base_folder_name = os.path.basename(self.image_folder_path)
            output_folder = os.path.join(os.path.dirname(self.image_folder_path), f"{base_folder_name}（已修改）")
            
            # 如果输出文件夹已存在，先删除
            if os.path.exists(output_folder):
                shutil.rmtree(output_folder)
            
            os.makedirs(output_folder)
            
            # 创建横图和竖图文件夹
            landscape_folder = os.path.join(output_folder, "横图")
            portrait_folder = os.path.join(output_folder, "竖图")
            os.makedirs(landscape_folder)
            os.makedirs(portrait_folder)
            
            # 获取重命名映射
            original_col = self.original_col_var.get()
            rename_col = self.rename_col_var.get()
            
            # 创建映射字典
            rename_map = {}
            for index, row in self.excel_data.iterrows():
                original_name = str(row[original_col]).strip()
                new_name = str(row[rename_col]).strip()
                if original_name != 'nan' and new_name != 'nan':
                    # 处理重命名列ID长度，确保不超过14位
                    processed_new_name = self.process_rename_id(new_name)
                    rename_map[original_name] = processed_new_name
            
            total_items = len(rename_map)
            processed_count = 0
            success_count = 0
            
            self.status_var.set("开始处理图片...")
            
            # 处理每个重命名项
            for original_name, new_name in rename_map.items():
                # 更新进度
                processed_count += 1
                progress = (processed_count / total_items) * 100
                self.progress_var.set(progress)
                self.root.update()
                
                # 查找对应的图片文件
                image_files = self.find_image_files(original_name)
                
                if not image_files:
                    print(f"未找到图片文件: {original_name}")
                    continue
                
                # 处理找到的图片文件
                for i, image_file in enumerate(image_files):
                    try:
                        # 判断是横图还是竖图
                        is_landscape = self.is_landscape_image(image_file)
                        target_folder = landscape_folder if is_landscape else portrait_folder
                        
                        # 确定新文件名
                        file_ext = os.path.splitext(image_file)[1]
                        if len(image_files) > 1:
                            # 如果有多个文件，给第二个文件添加(2)标识
                            if i == 0:
                                new_filename = f"{new_name}{file_ext}"
                            else:
                                new_filename = f"{new_name} (2){file_ext}"
                        else:
                            new_filename = f"{new_name}{file_ext}"
                        
                        # 复制并重命名文件
                        target_path = os.path.join(target_folder, new_filename)
                        shutil.copy2(image_file, target_path)
                        success_count += 1
                        
                    except Exception as e:
                        print(f"处理文件失败 {image_file}: {str(e)}")
            
            # 处理文件去重和重命名（删除(2)标识）
            self.status_var.set("正在处理文件去重...")
            self.remove_duplicate_markers_and_deduplicate(landscape_folder, portrait_folder)
            
            # 创建ZIP压缩包
            self.status_var.set("正在创建ZIP压缩包...")
            zip_path = f"{output_folder}.zip"
            if os.path.exists(zip_path):
                os.remove(zip_path)
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(output_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_path = os.path.relpath(file_path, os.path.dirname(output_folder))
                        zipf.write(file_path, arc_path)
            
            # 完成处理
            self.progress_var.set(100)
            self.status_var.set(f"处理完成！成功处理 {success_count} 个文件")
            
            # 自动打开输出文件夹
            if os.name == 'nt':  # Windows
                os.startfile(output_folder)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.run(['open', output_folder])
            
            messagebox.showinfo("完成", f"图片重命名完成！\n成功处理: {success_count} 个文件\n输出文件夹: {output_folder}\nZIP文件: {zip_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误: {str(e)}")
            self.status_var.set("处理失败")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    # 创建并运行应用程序
    app = ImageRenamer()
    app.run()
