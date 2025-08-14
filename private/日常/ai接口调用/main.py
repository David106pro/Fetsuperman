import os
import pandas as pd  # 这行代码本身没有问题。pandas是一个常用的数据处理库，用于读取和处理Excel文件。从代码上下文可以看出,这个程序需要处理Excel文件,所以导入pandas是必要的。
from api_client import shorten_text
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread
import time
from googletrans import Translator

class InfoShortenerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("内容简化工具")
        self.root.geometry("600x500")
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 步骤标签
        self.steps_label = ttk.Label(
            self.main_frame, 
            text="使用步骤：\n1. 选择Excel文件\n2. 选择要处理的列\n3. 点击开始处理",
            justify=tk.LEFT
        )
        self.steps_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        # 文件选择部分
        self.file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="5")
        self.file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.file_path = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=0, column=0, padx=5)
        
        self.file_button = ttk.Button(self.file_frame, text="浏览", command=self.select_file)
        self.file_button.grid(row=0, column=1, padx=5)
        
        # 列选择部分
        self.column_frame = ttk.LabelFrame(self.main_frame, text="列选择", padding="5")
        self.column_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(self.column_frame, textvariable=self.column_var, state='disabled', width=47)
        self.column_combo.grid(row=0, column=0, padx=5)
        
        # 添加控制按钮框架
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.grid(row=4, column=0, columnspan=3, pady=5)
        
        self.process_button = ttk.Button(self.control_frame, text="开始处理", command=self.start_processing, state='disabled')
        self.process_button.grid(row=0, column=0, padx=5)
        
        self.pause_button = ttk.Button(self.control_frame, text="暂停", command=self.pause_processing, state='disabled')
        self.pause_button.grid(row=0, column=1, padx=5)
        
        self.stop_button = ttk.Button(self.control_frame, text="停止", command=self.stop_processing, state='disabled')
        self.stop_button.grid(row=0, column=2, padx=5)
        
        # 添加控制标志
        self.is_processing = False
        self.is_paused = False
        self.should_stop = False
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            self.main_frame, 
            variable=self.progress_var, 
            maximum=100,
            length=400
        )
        self.progress.grid(row=5, column=0, columnspan=3, pady=5)
        
        # 状态信息
        self.status_var = tk.StringVar(value="请选择Excel文件开始")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var)
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # 日志文本框
        self.log_frame = ttk.LabelFrame(self.main_frame, text="处理日志", padding="5")
        self.log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 创建水平滚动条
        self.h_scrollbar = ttk.Scrollbar(self.log_frame, orient=tk.HORIZONTAL)
        self.h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        self.log_text = tk.Text(self.log_frame, height=8, width=50, wrap=tk.NONE,
                               xscrollcommand=self.h_scrollbar.set)
        self.log_text.grid(row=0, column=0, padx=5)
        
        # 垂直滚动条
        self.v_scrollbar = ttk.Scrollbar(self.log_frame, orient=tk.VERTICAL, 
                                        command=self.log_text.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置滚动条
        self.h_scrollbar.config(command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=self.v_scrollbar.set)
        
        # 初始化变量
        self.df = None
        
        # 初始化翻译器
        self.translator = Translator()
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title='选择Excel文件',
            filetypes=[('Excel files', '*.xlsx *.xls')]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.load_columns(file_path)
            
    def load_columns(self, file_path):
        try:
            self.df = pd.read_excel(file_path)
            self.column_combo['values'] = list(self.df.columns)
            self.column_combo['state'] = 'readonly'
            self.column_var.set(self.df.columns[0])
            self.process_button['state'] = 'normal'
            self.log("成功加载文件，请选择要处理的列")
        except Exception as e:
            self.log(f"加载文件失败: {str(e)}")
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
            
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        
    def start_processing(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_button['text'] = "暂停"
        else:
            self.is_processing = True
            self.should_stop = False
            self.process_button['state'] = 'disabled'
            self.pause_button['state'] = 'normal'
            self.stop_button['state'] = 'normal'
            Thread(target=self.process_file).start()
        
    def pause_processing(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_button['text'] = "暂停"
        else:
            self.is_paused = True
            self.pause_button['text'] = "继续"
        
    def stop_processing(self):
        self.should_stop = True
        self.is_paused = False
        
    def normalize_punctuation(self, text):
        # 将英文符号转换为中文符号
        punctuation_map = {
            ',': '，',
            '.': '。',
            '!': '！',
            '?': '？',
            ';': '；',
            ':': '：',
            '(': '（',
            ')': '）'
        }
        for en, cn in punctuation_map.items():
            text = text.replace(en, cn)
        
        # 确保文本以正确的中文符号结尾
        if not text.endswith(('。', '？', '！')):
            text += '。'
        return text

    def translate_text(self, text):
        """使用googletrans将英文翻译为中文"""
        try:
            # 如果文本全是中文，直接返回
            if all('\u4e00' <= char <= '\u9fff' or char in '，。！？；：（）' or char.isspace() for char in text):
                return text

            # 尝试翻译
            result = self.translator.translate(text, dest='zh-cn')
            return result.text
                
        except Exception as e:
            self.log(f"翻译出错: {str(e)}")
            return text  # 如果翻译失败，返回原文

    def process_text(self, text, max_retries=3):
        retry_count = 0
        while retry_count < max_retries:
            try:
                # 生成普通简介的提示词
                normal_prompt = """请处理以下内容：
                1. 如果内容不是中文，请先将其翻译为中文
                2. 然后将内容改写为简洁的中文简介
                3. 确保使用中文标点符号
                """
                
                # 生成标题的提示词
                title_prompt = """请处理以下内容：
                1. 如果内容不是中文，请先将其翻译为中文
                2. 然后将内容改写为15字以内的吸引人标题，遵循以下规则：
                   - 如有知名演员，使用"演员A和演员B做了什么"的格式
                   - 如无知名演员，使用"人物A和人物B做了什么"的格式
                   - 如故事不以人物为中心，使用"什么事情怎么样"的格式
                3. 确保标题简洁有力，富有吸引力
                4. 确保使用中文标点符号
                """
                
                # 调用 AI 接口时附加提示词
                shortened = shorten_text(normal_prompt + "\n" + str(text), mode="normal")
                title = shorten_text(title_prompt + "\n" + str(text), mode="title")
                
                # 规范化标点符号
                shortened = self.normalize_punctuation(shortened)
                title = self.normalize_punctuation(title)
                
                return shortened, title
                
            except Exception as e:
                retry_count += 1
                if retry_count == max_retries:
                    raise e
                self.log(f"处理失败，正在重试 ({retry_count}/{max_retries})...")
                time.sleep(1)

    def process_file(self):
        try:
            input_file = self.file_path.get()
            column_name = self.column_var.get()
            
            # 创建新的 DataFrame 来存储结果
            result_df = self.df.copy()
            new_column = f"{column_name}_简化"
            title_column = f"{column_name}_标题"  # 新增标题列
            result_df[new_column] = ''
            result_df[title_column] = ''  # 初始化标题列
            
            # 处理每一行
            total_rows = len(self.df)
            
            # 先处理第一条数据
            original_text = self.df.iloc[0][column_name]
            # 检查是否为空值或 NaN
            if pd.isna(original_text) or str(original_text).strip() == '':
                result_df.iloc[0, result_df.columns.get_loc(new_column)] = ''
                result_df.iloc[0, result_df.columns.get_loc(title_column)] = ''
                self.log(f"第 1 行为空，已跳过")
                self.progress_var.set((1 / total_rows) * 100)
                self.status_var.set(f"处理进度: {(1 / total_rows) * 100:.1f}% (1/{total_rows})")
            else:
                try:
                    # 生成两种长度的摘要
                    shortened, title = self.process_text(str(original_text))
                    
                    result_df.iloc[0, result_df.columns.get_loc(new_column)] = shortened
                    result_df.iloc[0, result_df.columns.get_loc(title_column)] = title
                    
                    self.log(f"已处理第 1 行")
                    self.progress_var.set((1 / total_rows) * 100)
                    self.status_var.set(f"处理进度: {(1 / total_rows) * 100:.1f}% (1/{total_rows})")
                except Exception as e:
                    error_msg = f"处理第 1 行时出错: {str(e)}"
                    self.log(error_msg)
                    self.status_var.set("处理失败，请检查错误信息")
                    messagebox.showerror("错误", f"处理第一条数据失败，流程已停止\n错误信息：{str(e)}")
                    return
            
            # 如果第一条数据处理成功，继续处理剩余数据
            for index in range(1, len(self.df)):
                if self.should_stop:
                    self.log("处理已停止")
                    break
                    
                while self.is_paused:
                    time.sleep(0.1)
                    if self.should_stop:
                        break
                
                progress = (index + 1) / total_rows * 100
                self.progress_var.set(progress)
                self.status_var.set(f"处理进度: {progress:.1f}% ({index + 1}/{total_rows})")
                
                original_text = self.df.iloc[index][column_name]
                if pd.isna(original_text) or str(original_text).strip() == '':
                    result_df.iloc[index, result_df.columns.get_loc(new_column)] = ''
                    result_df.iloc[index, result_df.columns.get_loc(title_column)] = ''
                    self.log(f"第 {index + 1} 行为空，已跳过")
                    continue
                
                try:
                    shortened, title = self.process_text(str(original_text))
                    result_df.iloc[index, result_df.columns.get_loc(new_column)] = shortened
                    result_df.iloc[index, result_df.columns.get_loc(title_column)] = title
                    self.log(f"已处理第 {index + 1} 行")
                except Exception as e:
                    error_msg = f"处理第 {index + 1} 行时出错: {str(e)}"
                    self.log(error_msg)
                    result_df.iloc[index, result_df.columns.get_loc(new_column)] = f"处理错误: {str(e)}"
                    result_df.iloc[index, result_df.columns.get_loc(title_column)] = f"处理错误: {str(e)}"
            
            # 保存结果
            output_dir = os.path.dirname(input_file)
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_file = os.path.join(output_dir, f"{base_name}_简化.xlsx")
            result_df.to_excel(output_file, index=False)
            self.log(f"处理完成，结果已保存到: {output_file}")
            self.status_var.set("处理完成！")
            messagebox.showinfo("成功", f"处理完成！\n输出文件保存在：{output_file}")
            
        except Exception as e:
            self.log(f"处理失败: {str(e)}")
            messagebox.showerror("错误", str(e))
            
        finally:
            self.is_processing = False
            self.is_paused = False
            self.should_stop = False
            self.process_button['state'] = 'normal'
            self.pause_button['state'] = 'disabled'
            self.stop_button['state'] = 'disabled'
            
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = InfoShortenerGUI()
    app.run() 