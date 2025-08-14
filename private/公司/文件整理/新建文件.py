import os
import re

# 输入路径选择
path_choice = input("请输入批次号：")

# 定义文件夹路径
base_paths = [r'E:\{}\4M'.format(path_choice), r'E:\{}\8M'.format(path_choice)]

# 新的文件名和ID
data = []
while True:
    input_str = input("请输入文件名和ID（用下划线隔开），回车确认，输入“好了”结束输入：")
    if input_str == "好了":
        break
    else:
        file_name, id = input_str.split('_')
        data.append((file_name, id))

# 处理节目名称，去除括号中的内容和所有符号
def clean_name(name):
    # 去掉括号及其中内容
    name = re.sub(r'\（.*?\）|\(.*?\)', '', name)
    # 去掉所有符号
    name = re.sub(r'[^\w\s]', '', name)
    # 去掉多余的空格
    name = name.strip()
    return name

# 创建文件夹
for base_path in base_paths:
    for name, id in data:
        clean_folder_name = f"{clean_name(name)}_{id}"
        folder_path = os.path.join(base_path, clean_folder_name)
        try:
            os.makedirs(folder_path, exist_ok=True)
            print(f"创建文件夹: {folder_path}")
        except Exception as e:
            print(f"创建文件夹失败: {folder_path}, 错误: {e}")
