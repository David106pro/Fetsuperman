import json
import time
import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime
import hmac
import hashlib
import base64
import re

class DoubanCrawler:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("豆瓣评分爬取工具")
        self.root.geometry("500x500")
        self.is_running = False
        self.is_paused = False
        self.setup_ui()

    def get_auth_header(self, secret_id, secret_key, host, timestamp, payload):
        """生成腾讯云 API 认证头"""
        date = datetime.utcfromtimestamp(int(timestamp)).strftime('%Y-%m-%d')
        
        http_request_method = "POST"
        canonical_uri = "/"
        canonical_querystring = ""
        canonical_headers = (
            "content-type:application/json\n"
            f"host:{host}\n"
            "x-tc-action:chatcompletions\n"
        )
        signed_headers = "content-type;host;x-tc-action"
        
        payload_hash = hashlib.sha256(json.dumps(payload).encode('utf-8')).hexdigest()
        
        canonical_request = (
            f"{http_request_method}\n"
            f"{canonical_uri}\n"
            f"{canonical_querystring}\n"
            f"{canonical_headers}\n"
            f"{signed_headers}\n"
            f"{payload_hash}"
        )
        
        algorithm = "TC3-HMAC-SHA256"
        credential_scope = f"{date}/hunyuan/tc3_request"
        string_to_sign = (
            f"{algorithm}\n"
            f"{timestamp}\n"
            f"{credential_scope}\n"
            f"{hashlib.sha256(canonical_request.encode('utf-8')).hexdigest()}"
        )
        
        def sign(key, msg):
            return hmac.new(key, msg.encode('utf-8'), hashlib.sha256).digest()
        
        secret_date = sign(("TC3" + secret_key).encode('utf-8'), date)
        secret_service = sign(secret_date, "hunyuan")
        secret_signing = sign(secret_service, "tc3_request")
        signature = hmac.new(secret_signing, string_to_sign.encode('utf-8'), hashlib.sha256).hexdigest()
        
        authorization = (
            f"{algorithm} "
            f"Credential={secret_id}/{credential_scope}, "
            f"SignedHeaders={signed_headers}, "
            f"Signature={signature}"
        )
        
        return authorization

    def get_douban_rating(self, name):
        """使用腾讯混元大模型获取豆瓣评分"""
        try:
            secret_id = "AKIDFEMTUM6odGiGi3HJSSJ8xhIgroW6RZT8"
            secret_key = "ayRxX5hGm8hDbmnfO1XryyoW3SineYUH"
            
            host = "hunyuan.tencentcloudapi.com"
            timestamp = str(int(time.time()))
            
            prompt = f"请帮我找出《{name}》对应的豆瓣评分，若是没有，则回复没有，如果有评分则告诉我他的评分"
            
            payload = {
                "Model": "hunyuan-large",
                "Messages": [
                    {
                        "Role": "user",
                        "Content": prompt
                    }
                ],
                "Temperature": 0.1,  # 降低随机性
                "TopP": 0.1,
                "Stream": False
            }
            
            headers = {
                "Content-Type": "application/json",
                "Host": host,
                "X-TC-Action": "ChatCompletions",
                "X-TC-Version": "2023-09-01",
                "X-TC-Timestamp": timestamp,
                "X-TC-Region": "",
                "Authorization": self.get_auth_header(secret_id, secret_key, host, timestamp, payload)
            }

            response = requests.post(
                "https://hunyuan.tencentcloudapi.com/",
                headers=headers,
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            result = response.json()
            
            if 'Response' in result and 'Choices' in result['Response']:
                ai_response = result['Response']['Choices'][0]['Message']['Content'].strip()
                print(f"AI响应: {ai_response}")
                
                # 使用正则表达式提取评分
                rating_match = re.search(r'(\d+\.\d+|\d+)', ai_response)
                if rating_match and '没有' not in ai_response:
                    return rating_match.group(1)
                return None
            else:
                raise Exception(f"API返回格式错误: {json.dumps(result)}")
                
        except Exception as e:
            print(f"获取评分失败: {str(e)}")
            return None
        finally:
            time.sleep(1)  # 避免请求过于频繁

    def setup_ui(self):
        # 文件选择框
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=20, padx=20, fill='x')
        
        self.file_path = tk.StringVar()
        tk.Label(file_frame, text="Excel文件：").pack(side='left')
        tk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side='left', padx=5)
        tk.Button(file_frame, text="选择文件", command=self.select_file).pack(side='left')

        # 表头选择框
        header_frame = tk.Frame(self.root)
        header_frame.pack(pady=20, padx=20, fill='x')
        
        tk.Label(header_frame, text="选择名称列：").pack(side='left')
        self.header_combobox = ttk.Combobox(header_frame, width=47)
        self.header_combobox.pack(side='left', padx=5)

        # 按钮框架
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)
        
        self.start_button = tk.Button(button_frame, text="开始爬取", command=self.start_crawling)
        self.start_button.pack(side='left', padx=5)
        
        self.pause_button = tk.Button(button_frame, text="暂停", command=self.pause_crawling, state='disabled')
        self.pause_button.pack(side='left', padx=5)
        
        self.stop_button = tk.Button(button_frame, text="结束", command=self.stop_crawling, state='disabled')
        self.stop_button.pack(side='left', padx=5)

        # 进度显示
        self.progress_var = tk.StringVar()
        self.progress_var.set("等待开始...")
        tk.Label(self.root, textvariable=self.progress_var).pack(pady=10)

        # 进度条
        self.progress_bar = ttk.Progressbar(self.root, length=400, mode='determinate')
        self.progress_bar.pack(pady=10)

        # 结果显示文本框
        self.result_text = tk.Text(self.root, height=8, width=50)
        self.result_text.pack(pady=10)
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.root)
        scrollbar.pack(side='right', fill='y')
        self.result_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.result_text.yview)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            try:
                # 读取Excel文件的表头
                df = pd.read_excel(file_path)
                columns = df.columns.tolist()
                self.header_combobox['values'] = columns
                
                # 自动选择可能的名称列
                for col in columns:
                    if any(keyword in col for keyword in ['名称', '片名', '标题', 'name', 'title']):
                        self.header_combobox.set(col)
                        break
                if not self.header_combobox.get():
                    self.header_combobox.set(columns[0])
                    
            except Exception as e:
                messagebox.showerror("错误", f"读取Excel文件失败：{str(e)}")

    def pause_crawling(self):
        """暂停或继续爬取"""
        self.is_paused = not self.is_paused
        self.pause_button.config(text="继续" if self.is_paused else "暂停")

    def stop_crawling(self):
        """结束爬取"""
        self.is_running = False
        self.save_current_results()
        self.driver.quit()  # 关闭浏览器

    def save_current_results(self):
        """保存当前已爬取的结果"""
        if hasattr(self, 'df'):
            output_file = self.file_path.get().rsplit('.', 1)[0] + '_with_ratings.xlsx'
            self.df.to_excel(output_file, index=False)
            messagebox.showinfo("已保存", f"已处理的结果已保存至：\n{output_file}")

    def start_crawling(self):
        file_path = self.file_path.get()
        name_column = self.header_combobox.get()
        
        if not file_path or not name_column:
            messagebox.showerror("错误", "请选择文件和名称列！")
            return
            
        try:
            # 读取Excel文件
            self.df = pd.read_excel(file_path)
            
            # 添加评分列，并设置为float类型
            if '豆瓣评分' not in self.df.columns:
                self.df['豆瓣评分'] = pd.Series(dtype='float64')
            
            # 设置进度条
            total_rows = len(self.df)
            self.progress_bar['maximum'] = total_rows
            
            # 更新按钮状态
            self.start_button.config(state='disabled')
            self.pause_button.config(state='normal')
            self.stop_button.config(state='normal')
            
            # 清空结果显示
            self.result_text.delete(1.0, tk.END)
            
            self.is_running = True
            self.is_paused = False
            
            # 遍历每一行获取评分
            for index, row in self.df.iterrows():
                if not self.is_running:
                    break
                    
                while self.is_paused:
                    self.root.update()
                    time.sleep(0.1)
                    continue
                    
                name = row[name_column]
                self.progress_var.set(f"正在处理 {index+1}/{total_rows}: {name}")
                self.progress_bar['value'] = index + 1
                
                rating = self.get_douban_rating(name)
                
                # 将评分转换为浮点数
                try:
                    rating_float = float(rating) if rating else None
                    self.df.at[index, '豆瓣评分'] = rating_float
                except (ValueError, TypeError):
                    self.df.at[index, '豆瓣评分'] = None
                
                # 显示结果
                result_text = f"《{name}》 {rating if rating else '无'}\n"
                self.result_text.insert(tk.END, result_text)
                self.result_text.see(tk.END)  # 自动滚动到最新结果
                
                # 保存当前进度
                self.df.to_excel(file_path.rsplit('.', 1)[0] + '_with_ratings.xlsx', index=False)
                
                self.root.update()
                time.sleep(random.uniform(1, 3))
            
            if self.is_running:
                messagebox.showinfo("完成", "所有数据处理完成！")
            
            # 重置按钮状态
            self.start_button.config(state='normal')
            self.pause_button.config(state='disabled')
            self.stop_button.config(state='disabled')
            self.is_running = False
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出错：{str(e)}")
            self.progress_var.set("处理出错！")
            self.start_button.config(state='normal')
            self.pause_button.config(state='disabled')
            self.stop_button.config(state='disabled')

    def run(self):
        self.root.mainloop()

    def __del__(self):
        """确保程序结束时关闭浏览器"""
        if hasattr(self, 'driver'):
            self.driver.quit()

if __name__ == "__main__":
    app = DoubanCrawler()
    app.run()