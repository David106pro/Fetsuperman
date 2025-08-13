# -*- coding: utf-8 -*-
import re

s1 = '《错过你的那些年》【正片介质链接】https://pan.baidu.com/s/1PAH6sU10PAaY8xwQ6udqXw?pwd=ktq2'
s2 = '《错过你的那些年》【8M】https://pan.baidu.com/s/14MeMsCIp2NdF2CEpHHBVtA?pwd=gnn1'

chinese_pattern = r'《([^《》]+)》'
size_pattern = r'【([^【】]+)】'

for s in [s1, s2]:
    title = re.search(chinese_pattern, s).group(1)
    size = re.search(size_pattern, s).group(1)
    parts = s.split('】')
    link = parts[1].strip()
    print(f"标题：{title}")
    print(f"格式：{size}")
    print(f"链接：{link}")
    print()