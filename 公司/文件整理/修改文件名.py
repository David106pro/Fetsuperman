# -*- coding: utf-8 -*-
import os
import re

paths = [r'E:\2409\4M', r'E:\2409\8M']
original_name_format = r'奇特益智工程车第2季_20_(\d+)(.*).ts'
new_name_format = r'奇特益智工程车第二季_20_$1.ts'
folder_path = r'E:\2409\4M\奇特益智工程车第二季_20'

def rename_num(paths):
    for path in paths:
        for folder_name in os.listdir(path):
            folder_path = os.path.join(path, folder_name)
            if os.path.isdir(folder_path):
                file_count = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
                if file_count > 1:
                    new_folder_name = f"{folder_name}_{file_count}"
                    new_folder_path = os.path.join(path, new_folder_name)
                    os.rename(folder_path, new_folder_path)

def rename_files(folder_path, original_pattern, new_pattern):
    for filename in os.listdir(folder_path):
        if re.match(original_pattern, filename):
            print(f"匹配到的文件：{filename}")
            new_filename = re.sub(original_pattern, new_pattern, filename)

            old_file_path = os.path.join(folder_path, filename)
            new_file_path = os.path.join(folder_path, new_filename)
            os.rename(old_file_path, new_file_path)

rename_files(folder_path, original_name_format, new_name_format)
# rename_num(paths)