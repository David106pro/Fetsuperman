#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试数据生成脚本
用于创建批次号映射测试数据
"""

import pandas as pd
import os

def create_test_data():
    """创建测试数据Excel文件"""
    # 创建测试数据
    test_data = {
        'ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
        'name': ['电影1', '电影2', '电影3', '电影4', '电影5', '电影6', '电影7', '电影8', '电影9', '电影10', '电影11', '电影12', '电影13'],
        'batch': ['240101', '240626', '240725', '240816', '240905', '241015', '241122', '241212', '250114', '250307', '250507', '250613', '250624'],
        '批次': ['', '', '', '', '', '', '', '', '', '', '', '', '']
    }
    
    # 创建DataFrame
    df = pd.DataFrame(test_data)
    
    # 保存为Excel文件
    output_path = os.path.join(os.path.dirname(__file__), '测试数据.xlsx')
    df.to_excel(output_path, index=False)
    print(f"测试数据文件已创建：{output_path}")
    return output_path

if __name__ == "__main__":
    create_test_data() 