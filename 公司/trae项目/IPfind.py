import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import re

class IPFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IP整理工具")
        self.root.geometry("600x200")
        
        # 创建文件选择框和按钮
        self.file_frame = tk.Frame(root)
        self.file_frame.pack(pady=20)
        
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_path, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        
        self.browse_button = tk.Button(self.file_frame, text="选择文件", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT)
        
        # 创建执行按钮
        self.execute_button = tk.Button(root, text="开始执行", command=self.process_file)
        self.execute_button.pack(pady=20)
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if filename:
            self.file_path.set(filename)
            
    def extract_ip(self, titles):
        if not titles:
            return {}
        
        def clean_title(title):
            # 移除季数、版本等信息
            title = re.sub(r'第[一二三四五六七八九十\d]+[季部].*$', '', title)
            title = re.sub(r'系列.*$', '', title)
            title = re.sub(r'中文版$|英文版$', '', title)
            return title.strip()
        
        # 清理所有标题并按清理后的标题分组
        title_groups = {}
        for title in titles:
            cleaned = clean_title(title)
            if cleaned not in title_groups:
                title_groups[cleaned] = []
            title_groups[cleaned].append(title)
        
        # 只保留出现2次以上的组
        final_groups = {ip: group for ip, group in title_groups.items() if len(group) >= 2}
        
        return final_groups
    def process_file(self):
        try:
            file_path = self.file_path.get()
            if not file_path:
                messagebox.showerror("错误", "请选择文件！")
                return
                
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 查找"介质名称"列
            media_column = None
            for col in df.columns:
                if "介质名称" in str(col):
                    media_column = col
                    break
                    
            if media_column is None:
                messagebox.showerror("错误", "未找到'介质名称'列！")
                return
                
            # 对介质名称列进行分组并提取IP
            grouped_titles = df[media_column].dropna().tolist()
            ip_groups = self.extract_ip(grouped_titles)
            
            # 准备结果数据
            result_data = []
            # 处理已匹配到IP的数据
            processed_titles = set()
            for ip, titles in ip_groups.items():
                for title in titles:
                    # 获取原始行的所有需要的字段
                    row = df[df[media_column] == title].iloc[0]
                    result_data.append({
                        "介质名称": title,
                        "IP": ip,
                        "同IP数量": len(titles),
                        "cid": row.get('cid', ''),
                        "出品年代": row.get('出品年代', ''),
                        "出品地区": row.get('出品地区', ''),
                        "频道属性": row.get('频道属性', '')
                    })
                    processed_titles.add(title)
            
            # 处理未匹配到IP的数据
            for title in grouped_titles:
                if title not in processed_titles:
                    # 获取原始行的所有需要的字段
                    row = df[df[media_column] == title].iloc[0]
                    result_data.append({
                        "介质名称": title,
                        "IP": title,
                        "同IP数量": 1,
                        "cid": row.get('cid', ''),
                        "出品年代": row.get('出品年代', ''),
                        "出品地区": row.get('出品地区', ''),
                        "频道属性": row.get('频道属性', '')
                    })
            
            # 创建新的DataFrame，设置列顺序
            result_df = pd.DataFrame(result_data)
            columns_order = ["介质名称", "IP", "同IP数量", "cid", "出品年代", "出品地区", "频道属性"]
            result_df = result_df[columns_order]
            
            # 按IP分组排序
            result_df = result_df.sort_values(['IP', '介质名称'])
            
            # 生成新文件名
            file_name = os.path.splitext(file_path)[0]
            new_file_path = f"{file_name}_ip整理.xlsx"
            
            # 保存结果
            result_df.to_excel(new_file_path, index=False)
            
            messagebox.showinfo("成功", f"IP整理完成！\n文件保存在：{new_file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = IPFinderApp(root)
    root.mainloop()