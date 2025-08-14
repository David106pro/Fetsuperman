import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import zipfile
import tempfile

class FileProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件夹处理工具")
        self.root.geometry("600x400")

        # 创建输入框和按钮
        self.create_widgets()

    def create_widgets(self):
        # 源文件夹选择
        tk.Label(self.root, text="源文件夹路径:").pack(pady=5)
        self.source_frame = tk.Frame(self.root)
        self.source_frame.pack(fill=tk.X, padx=5)
        
        self.source_path = tk.StringVar()
        tk.Entry(self.source_frame, textvariable=self.source_path).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(self.source_frame, text="浏览", command=self.select_source).pack(side=tk.RIGHT)

        # 目标文件夹选择
        tk.Label(self.root, text="目标文件夹路径:").pack(pady=5)
        self.target_frame = tk.Frame(self.root)
        self.target_frame.pack(fill=tk.X, padx=5)
        
        self.target_path = tk.StringVar()
        tk.Entry(self.target_frame, textvariable=self.target_path).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(self.target_frame, text="浏览", command=self.select_target).pack(side=tk.RIGHT)

        # 目标文件夹名称
        tk.Label(self.root, text="目标文件夹名称:").pack(pady=5)
        self.folder_name = tk.StringVar()
        tk.Entry(self.root, textvariable=self.folder_name).pack(fill=tk.X, padx=5)

        # 执行按钮
        tk.Button(self.root, text="开始处理", command=self.process_files).pack(pady=20)

    def select_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_path.set(folder)

    def select_target(self):
        folder = filedialog.askdirectory()
        if folder:
            self.target_path.set(folder)

    def process_files(self):
        source = self.source_path.get()
        target_base = self.target_path.get()
        folder_name = self.folder_name.get()

        if not source or not target_base or not folder_name:
            messagebox.showerror("错误", "请填写所有必要信息！")
            return

        # 创建目标文件夹
        target = os.path.join(target_base, folder_name)
        if not os.path.exists(target):
            os.makedirs(target)

        try:
            # 创建临时目录用于解压文件
            with tempfile.TemporaryDirectory() as temp_dir:
                # 首先处理压缩文件
                for root, dirs, files in os.walk(source):
                    for file in files:
                        if file.lower().endswith(('.zip', '.rar', '.7z')):
                            source_file = os.path.join(root, file)
                            try:
                                if file.lower().endswith('.zip'):
                                    with zipfile.ZipFile(source_file, 'r') as zip_ref:
                                        # 修改解压部分，正确处理中文文件名
                                        for zip_info in zip_ref.filelist:
                                            zip_info.filename = zip_info.filename.encode('cp437').decode('gbk')
                                            try:
                                                zip_ref.extract(zip_info, temp_dir)
                                            except:
                                                # 如果gbk解码失败，尝试使用utf-8
                                                zip_info.filename = zip_info.filename.encode('cp437').decode('utf-8')
                                                try:
                                                    zip_ref.extract(zip_info, temp_dir)
                                                except:
                                                    messagebox.showwarning("警告", f"无法正确解码文件名: {zip_info.filename}")
                            except Exception as e:
                                messagebox.showwarning("警告", f"解压文件 {file} 时出错：{str(e)}")

                # 处理所有文件（包括解压的文件）
                for root, dirs, files in os.walk(source):
                    for file in files:
                        source_file = os.path.join(root, file)
                        target_file = os.path.join(target, file)
                        
                        # 如果目标文件已存在，添加数字后缀
                        base, ext = os.path.splitext(target_file)
                        counter = 1
                        while os.path.exists(target_file):
                            target_file = f"{base}_{counter}{ext}"
                            counter += 1
                        
                        # 复制文件
                        shutil.copy2(source_file, target_file)

                # 处理临时目录中解压的文件
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        source_file = os.path.join(root, file)
                        target_file = os.path.join(target, file)
                        
                        # 如果目标文件已存在，添加数字后缀
                        base, ext = os.path.splitext(target_file)
                        counter = 1
                        while os.path.exists(target_file):
                            target_file = f"{base}_{counter}{ext}"
                            counter += 1
                        
                        # 复制文件
                        shutil.copy2(source_file, target_file)

            messagebox.showinfo("成功", "文件处理完成！")
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时出错：{str(e)}")

def main():
    root = tk.Tk()
    app = FileProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
