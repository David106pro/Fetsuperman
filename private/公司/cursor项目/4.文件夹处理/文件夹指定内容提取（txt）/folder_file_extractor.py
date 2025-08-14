import os
import shutil
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from pathlib import Path
import re

class FolderFileExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("文件夹指定内容提取")
        self.root.geometry("800x600")
        
        # 默认txt路径
        self.default_txt_path = r"C:\Users\fucha\PycharmProjects\python\公司\cursor项目\4.文件夹处理\文件夹指定内容提取（txt）\test.txt"
        
        # 创建界面组件
        self.create_widgets()
        
        # 如果默认txt存在，加载它
        if os.path.exists(self.default_txt_path):
            self.txt_path_entry.delete(0, tk.END)
            self.txt_path_entry.insert(0, self.default_txt_path)
            self.load_txt_content()
    
    def create_widgets(self):
        # 文件夹路径选择
        tk.Label(self.root, text="源文件夹路径:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.folder_path_entry = tk.Entry(self.root, width=60)
        self.folder_path_entry.grid(row=0, column=1, padx=5, pady=10)
        tk.Button(self.root, text="浏览...", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=10)
        
        # TXT文件路径选择
        tk.Label(self.root, text="TXT文件路径:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.txt_path_entry = tk.Entry(self.root, width=60)
        self.txt_path_entry.grid(row=1, column=1, padx=5, pady=10)
        self.txt_path_entry.insert(0, self.default_txt_path)
        tk.Button(self.root, text="浏览...", command=self.browse_txt).grid(row=1, column=2, padx=5, pady=10)
        tk.Button(self.root, text="加载TXT内容", command=self.load_txt_content).grid(row=1, column=3, padx=5, pady=10)
        
        # TXT内容显示区域
        tk.Label(self.root, text="TXT文件内容:").grid(row=2, column=0, sticky="nw", padx=10, pady=10)
        self.txt_content = scrolledtext.ScrolledText(self.root, width=80, height=20)
        self.txt_content.grid(row=3, column=0, columnspan=4, padx=10, pady=10)
        
        # 开始按钮
        self.start_button = tk.Button(self.root, text="开始提取", command=self.start_extraction, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2)
        self.start_button.grid(row=4, column=0, columnspan=4, padx=10, pady=20)
    
    def browse_folder(self):
        folder_path = filedialog.askdirectory(title="选择源文件夹")
        if folder_path:
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, folder_path)
    
    def browse_txt(self):
        txt_path = filedialog.askopenfilename(title="选择TXT文件", filetypes=[("Text files", "*.txt")])
        if txt_path:
            self.txt_path_entry.delete(0, tk.END)
            self.txt_path_entry.insert(0, txt_path)
            self.load_txt_content()
    
    def load_txt_content(self):
        txt_path = self.txt_path_entry.get()
        if not os.path.exists(txt_path):
            messagebox.showerror("错误", f"文件不存在: {txt_path}")
            return
        
        try:
            with open(txt_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            self.txt_content.delete(1.0, tk.END)
            self.txt_content.insert(tk.END, content)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def start_extraction(self):
        folder_path = self.folder_path_entry.get()
        txt_path = self.txt_path_entry.get()
        
        if not folder_path or not os.path.exists(folder_path):
            messagebox.showerror("错误", "请选择有效的源文件夹路径")
            return
        
        if not txt_path or not os.path.exists(txt_path):
            messagebox.showerror("错误", "请选择有效的TXT文件路径")
            return
        
        try:
            # 读取TXT文件内容
            with open(txt_path, 'r', encoding='utf-8') as f:
                file_prefixes = [line.strip() for line in f.readlines() if line.strip()]
            
            if not file_prefixes:
                messagebox.showinfo("提示", "TXT文件中没有文件名")
                return
            
            # 获取源文件夹中的所有文件
            all_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
            
            # 根据前缀匹配文件
            matched_files = []
            not_found = []
            
            for prefix in file_prefixes:
                found = False
                # 查找所有匹配的文件
                for file in all_files:
                    name_part = os.path.splitext(file)[0]
                    if name_part == prefix:
                        matched_files.append(file)
                        found = True
                
                if not found:
                    not_found.append(prefix)
            
            # 只输出没有匹配到的内容
            if not_found:
                messagebox.showwarning("警告", f"以下{len(not_found)}个文件未找到:\n" + 
                                      "\n".join(not_found[:10]) + 
                                      (f"\n...等{len(not_found)-10}个" if len(not_found) > 10 else ""))
            
            if not matched_files:
                messagebox.showerror("错误", "未找到任何匹配的文件，请检查文件名")
                return
            
            # 确定文件类型
            file_extensions = []
            for file_name in matched_files:
                if '.' in file_name:
                    ext = os.path.splitext(file_name)[1].lower().replace('.', '')
                    if ext and ext not in file_extensions:
                        file_extensions.append(ext)
            
            # 如果有多种文件类型，使用"files"，否则使用具体的文件类型
            if len(file_extensions) > 1:
                file_type = "files"
            elif len(file_extensions) == 1:
                file_type = file_extensions[0]
            else:
                file_type = "files"
            
            # 创建新文件夹
            new_folder_name = f"{len(matched_files)}{file_type}"
            new_folder_path = os.path.join(folder_path, new_folder_name)
            
            # 确保新文件夹不存在，如果存在则加上序号
            if os.path.exists(new_folder_path):
                counter = 1
                while os.path.exists(f"{new_folder_path}_{counter}"):
                    counter += 1
                new_folder_path = f"{new_folder_path}_{counter}"
            
            os.makedirs(new_folder_path, exist_ok=True)
            
            # 复制文件
            copied_count = 0
            for file_name in matched_files:
                source_path = os.path.join(folder_path, file_name)
                dest_path = os.path.join(new_folder_path, file_name)
                try:
                    shutil.copy2(source_path, dest_path)
                    copied_count += 1
                except Exception:
                    pass
            
            if copied_count == 0:
                messagebox.showerror("错误", "没有文件被复制，请检查文件权限或路径")
            else:
                messagebox.showinfo("完成", f"已成功复制 {copied_count}/{len(matched_files)} 个文件到 {new_folder_path}")
        
        except Exception as e:
            messagebox.showerror("错误", f"提取过程中出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderFileExtractor(root)
    root.mainloop() 