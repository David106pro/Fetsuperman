# -*- coding: utf-8 -*-
content = input("请输入内容：")

lines = content.split('\n')
formatted_content = []
for line in lines:
    parts = line.split('\t')
    if len(parts) == 2:
        formatted_content.append(f"{parts[1]}_{parts[0]}")
    else:
        print(f"该行无法正确分割：{line}")

for item in formatted_content:
    print(item)
    print()