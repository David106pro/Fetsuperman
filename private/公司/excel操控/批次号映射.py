#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批次号映射脚本
功能：根据映射表将W列的批次号转换为X列的批次名称
作者：AI助手
日期：2025-01-30
"""

import pandas as pd
import os
import logging
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('batch_mapping.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def create_batch_mapping():
    """
    创建批次号映射字典
    返回：批次号到批次名称的映射字典
    """
    batch_mapping = {
        '240101': '新片批次',
        '240626': '第一批',
        '240725': '第二批',
        '240816': '第三批',
        '240905': '第四批',
        '241015': '第五批',
        '241122': '第六批',
        '241212': '第七批',
        '250114': '第八批',
        '250307': '第九批',
        '250507': '第十批',
        '250613': '第十一批',
        '250624': '第十二批'
    }
    return batch_mapping

def validate_file_path(file_path):
    """
    验证文件路径是否有效
    参数：file_path - 文件路径
    返回：布尔值，True表示有效
    """
    if not file_path:
        return False
    if not os.path.exists(file_path):
        return False
    if not file_path.lower().endswith(('.xlsx', '.xls')):
        return False
    return True

def find_column_index(df, column_names):
    """
    查找列索引
    参数：df - DataFrame对象，column_names - 可能的列名列表
    返回：列索引，如果找不到返回None
    """
    for col_name in column_names:
        if col_name in df.columns:
            return df.columns.get_loc(col_name)
    return None

def process_batch_mapping(file_path, save_path=None):
    """
    处理批次号映射的主函数
    参数：file_path - 输入文件路径，save_path - 输出文件路径（可选）
    返回：处理结果信息
    """
    try:
        # 验证文件路径
        if not validate_file_path(file_path):
            raise ValueError("无效的文件路径或文件不存在")
        
        logging.info(f"开始处理文件：{file_path}")
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        logging.info(f"成功读取文件，共{len(df)}行数据")
        
        # 获取批次映射字典
        batch_mapping = create_batch_mapping()
        
        # 查找W列（batch列）
        w_column_names = ['batch', 'Batch', 'BATCH', 'W']
        w_column_index = find_column_index(df, w_column_names)
        
        if w_column_index is None:
            # 如果找不到指定列名，尝试使用第23列（W列）
            if len(df.columns) > 22:
                w_column_index = 22  # W列是第23列（索引从0开始）
                logging.warning("未找到batch列，使用第23列（W列）作为batch列")
            else:
                raise ValueError("无法找到batch列，请检查文件格式")
        
        # 查找X列（批次列）
        x_column_names = ['批次', '批次名称', 'X']
        x_column_index = find_column_index(df, x_column_names)
        
        if x_column_index is None:
            # 如果找不到批次列，创建新列
            if len(df.columns) > 23:
                x_column_index = 23  # X列是第24列（索引从0开始）
                logging.warning("未找到批次列，使用第24列（X列）作为批次列")
            else:
                # 在末尾添加新列
                df['批次'] = ''
                x_column_index = len(df.columns) - 1
                logging.info("创建新的批次列")
        
        # 获取列名
        w_column_name = df.columns[w_column_index]
        x_column_name = df.columns[x_column_index]
        
        logging.info(f"使用列：{w_column_name}（索引{w_column_index}）-> {x_column_name}（索引{x_column_index}）")
        
        # 处理映射
        successful_mappings = 0
        failed_mappings = 0
        
        for index, row in df.iterrows():
            batch_code = str(row[w_column_name]).strip()
            
            if batch_code in batch_mapping:
                df.at[index, x_column_name] = batch_mapping[batch_code]
                successful_mappings += 1
            elif batch_code and batch_code != 'nan':
                # 如果有值但不在映射表中，记录但不处理
                logging.warning(f"第{index+2}行的批次号'{batch_code}'不在映射表中")
                failed_mappings += 1
        
        # 保存文件
        if save_path is None:
            # 生成输出文件名
            file_dir = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            name, ext = os.path.splitext(file_name)
            save_path = os.path.join(file_dir, f"{name}_批次映射{ext}")
        
        df.to_excel(save_path, index=False)
        
        result_msg = f"""
处理完成！
输入文件：{file_path}
输出文件：{save_path}
成功映射：{successful_mappings} 条
失败映射：{failed_mappings} 条
总行数：{len(df)} 行
"""
        
        logging.info(result_msg)
        return True, result_msg
        
    except Exception as e:
        error_msg = f"处理过程中发生错误：{str(e)}"
        logging.error(error_msg)
        return False, error_msg

class BatchMappingGUI:
    """批次号映射GUI界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("批次号映射工具")
        self.root.geometry("600x400")
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        """创建GUI组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(file_frame, text="输入文件：").grid(row=0, column=0, sticky=tk.W)
        self.file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).grid(row=0, column=2, padx=5)
        
        # 映射表显示区域
        mapping_frame = ttk.LabelFrame(main_frame, text="批次号映射表", padding="10")
        mapping_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 创建映射表显示
        self.create_mapping_display(mapping_frame)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="开始处理", command=self.process_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="处理结果", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.result_text = tk.Text(result_frame, height=8, width=60)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 设置默认文件路径
        self.file_var.set(r"C:\Users\fucha\Desktop\全量批次分类表.xlsx")
        
    def create_mapping_display(self, parent):
        """创建映射表显示"""
        mapping_text = tk.Text(parent, height=10, width=60)
        mapping_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=mapping_text.yview)
        mapping_text.configure(yscrollcommand=mapping_scrollbar.set)
        
        # 添加映射表内容
        batch_mapping = create_batch_mapping()
        mapping_content = "批次号 -> 批次名称\n" + "-" * 30 + "\n"
        for code, name in batch_mapping.items():
            mapping_content += f"{code} -> {name}\n"
        
        mapping_text.insert(tk.END, mapping_content)
        mapping_text.configure(state=tk.DISABLED)
        
        mapping_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mapping_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
    def browse_file(self):
        """浏览文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_var.set(filename)
    
    def process_file(self):
        """处理文件"""
        file_path = self.file_var.get()
        
        if not file_path:
            messagebox.showerror("错误", "请选择输入文件")
            return
        
        # 清空结果显示
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "正在处理...\n")
        self.root.update()
        
        # 处理文件
        success, result_msg = process_batch_mapping(file_path)
        
        # 显示结果
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result_msg)
        
        if success:
            messagebox.showinfo("成功", "批次号映射处理完成！")
        else:
            messagebox.showerror("错误", result_msg)

def main():
    """主函数"""
    try:
        # 检查是否有命令行参数
        if len(sys.argv) > 1:
            # 命令行模式
            file_path = sys.argv[1]
            success, result_msg = process_batch_mapping(file_path)
            print(result_msg)
            return 0 if success else 1
        else:
            # GUI模式
            root = tk.Tk()
            app = BatchMappingGUI(root)
            root.mainloop()
            return 0
            
    except Exception as e:
        logging.error(f"程序运行错误：{str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
