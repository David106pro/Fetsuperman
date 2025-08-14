# -*- coding: utf-8 -*-
import os
import pandas as pd

excel_path = "C:\\Users\\fucha\\Desktop\\爬虫\\爬取结果去重\\猫眼爬取_手动去重.xlsx"
poster_folder = "D:\\猫眼海报"

# 读取 Excel 文件
df = pd.read_excel(excel_path)

# 获取 id 列
ids = df['媒资号'].tolist()

# 遍历海报文件夹，删除对应的海报
for filename in os.listdir(poster_folder):
    if filename.endswith('.jpg'):
        file_id = filename.split('.')[0]
        if file_id in ids:
            file_path = os.path.join(poster_folder, filename)
            os.remove(file_path)