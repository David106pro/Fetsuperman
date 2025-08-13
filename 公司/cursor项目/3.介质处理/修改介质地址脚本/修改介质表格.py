import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class MediaPathEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("介质路径修改工具")
        self.root.geometry("1000x700")
        
        # 设置整体样式
        style = ttk.Style()
        style.configure("TButton", padding=6, font=('微软雅黑', 10))
        style.configure("TLabel", font=('微软雅黑', 10))
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 创建顶部操作区域
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.grid(row=0, column=0, pady=(0, 10), sticky=(tk.W, tk.E))
        
        # 选择文件按钮
        self.select_btn = ttk.Button(self.top_frame, text="选择Excel文件", command=self.select_file)
        self.select_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 文件路径显示
        self.file_path_var = tk.StringVar()
        self.file_path_label = ttk.Label(self.top_frame, textvariable=self.file_path_var)
        self.file_path_label.grid(row=0, column=1, sticky=tk.W)
        
        # 创建表格显示区域
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # 创建表格和滚动条
        self.tree = ttk.Treeview(self.tree_frame)
        self.tree_scroll_y = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        
        # 布局表格和滚动条
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.rowconfigure(0, weight=1)
        
        # 创建底部操作区域
        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.grid(row=2, column=0, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # 开始处理按钮
        self.process_btn = ttk.Button(self.bottom_frame, text="开始处理", command=self.process_file)
        self.process_btn.grid(row=0, column=0, pady=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(self.bottom_frame, length=400, mode='determinate', 
                                      variable=self.progress_var)
        self.progress.grid(row=0, column=1, padx=(20, 0), pady=5)
        
        self.df = None
        
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path_var.set(file_path)
            self.load_excel(file_path)
    
    def load_excel(self, file_path):
        self.df = pd.read_excel(file_path)
        
        # 清除现有表格内容
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 设置列
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        
        for column in self.df.columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=100)
        
        # 添加数据
        for i, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))
    
    def process_file(self):
        if self.df is None:
            return
        
        total_rows = len(self.df)
        for i, row in self.df.iterrows():
            # 更新进度条
            self.progress_var.set((i + 1) / total_rows * 100)
            self.root.update()
            
            if 'title' in self.df.columns and 'rate' in self.df.columns and 'media_path' in self.df.columns:
                title = str(row['title'])
                if '_' in title:
                    name, episode = title.rsplit('_', 1)
                    # 找到相同名称的最大集数
                    max_episode = max([int(str(t).split('_')[-1]) for t in self.df['title'] 
                                    if str(t).startswith(name + '_')])
                    
                    # 构建新的media_path
                    rate = str(row['rate']) if 'rate' in self.df.columns else '8'
                    new_path_parts = [
                        "",
                        "data1",
                        "hs4",
                        "baiduyun",
                        "zongbu",
                        f"{rate}M",
                        f"{name}_{max_episode}",
                        f"{name}_{max_episode}_{episode}.ts"  # 修改最后一项的格式
                    ]
                    
                    # 更新DataFrame中的media_path
                    self.df.at[i, 'media_path'] = '/'.join(new_path_parts)
        
        # 直接保存到原文件
        self.df.to_excel(self.file_path_var.get(), index=False)
        tk.messagebox.showinfo("完成", "处理完成！文件已更新")
        
        # 重新加载显示修改后的数据
        self.load_excel(self.file_path_var.get())

if __name__ == "__main__":
    root = tk.Tk()
    app = MediaPathEditor(root)
    root.mainloop()
