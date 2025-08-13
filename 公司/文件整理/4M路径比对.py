file_path_1 = "C:\\Users\\fucha\\Desktop\\240905_4m.txt"
file_path_2 = "C:\\Users\\fucha\\Desktop\\m4_path.txt"

# 读取文件 2 的内容并存入列表，指定编码方式为 utf-8（可根据实际情况调整）
with open(file_path_2, 'r', encoding='utf-8') as file2:
    content_file2 = [line.strip() for line in file2.readlines()]

# 找出文件 2 中的重复内容
duplicates = [item for item in content_file2 if content_file2.count(item) > 1]

# 去除 duplicates 中的所有 '0'
duplicates = [item for item in duplicates if item!= '0']

# # 读取文件 1 的内容
# with open(file_path_1, 'r', encoding='utf-8') as file1:
#     content_file1 = [line.strip() for line in file1.readlines()]
#
# # 从文件 1 中删除与文件 2 相同的内容
# content_file1 = [line for line in content_file1 if line not in content_file2]
#
# # 将处理后的内容写回文件 1
# with open(file_path_1, 'w', encoding='utf-8') as file1:
#     for line in content_file1:
#         file1.write(line + '\n')

# 逐行打印重复内容
for duplicate in duplicates:
    print(duplicate)