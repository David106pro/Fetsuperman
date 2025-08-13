#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件 - 柱子区域统计分析脚本

用于配置文件路径和区域定义等参数
"""

# DWG/DXF文件路径配置
# 现在支持DWG文件自动转换，也可以直接使用DXF文件
DXF_FILE_PATH = r"E:\code\CursorCode\公司\cursor项目\6.老爹脚本\找柱子\S-2#地下车库结构图.dwg"

# 区域定义配置
# 格式：每个区域包含负责人、X轴范围、Y轴范围
REGIONS_CONFIG = [
    {
        'name': '吴晨料',
        'x_range': '17-25',
        'y_range': 'AA-N',
        'description': '吴晨料: (17-25/AA-N)'
    },
    {
        'name': '吴晨料',
        'x_range': '27-35',
        'y_range': 'AA-N',
        'description': '吴晨料: (27-35/AA-N)'
    },
    {
        'name': '吴晨料',
        'x_range': '21-27',
        'y_range': 'CA-N',
        'description': '吴晨料: (21-27/CA-N)'
    },
    {
        'name': '王姣',
        'x_range': '9-15',
        'y_range': 'E-M',
        'description': '王姣: (9-15/E-M)'
    },
    {
        'name': '王姣',
        'x_range': '15-21',
        'y_range': 'E-N',
        'description': '王姣: (15-21/E-N)'
    },
    {
        'name': '王姣',
        'x_range': '1-55',
        'y_range': 'F-H',
        'description': '王姣: (1-55/F-H)'
    },
    {
        'name': '王姣',
        'x_range': '5-9',
        'y_range': 'E-M',
        'description': '王姣: (5-9/E-M)'
    },
    {
        'name': '度小彤',
        'x_range': '15-21',
        'y_range': 'X-CA',
        'description': '度小彤: (15-21/X-CA)'
    },
    {
        'name': '度小彤',
        'x_range': '9-15',
        'y_range': 'X-CA',
        'description': '度小彤: (9-15/X-CA)'
    },
    {
        'name': '度小彤',
        'x_range': '5-10',
        'y_range': 'AA-M',
        'description': '度小彤: (5-10/AA-M)'
    },
    {
        'name': '度小彤',
        'x_range': '1-5',
        'y_range': 'X-CA',
        'description': '度小彤: (1-5/X-CA)'
    }
]

# 柱子识别配置
# 用于识别柱子的关键字符
COLUMN_IDENTIFIER = 'Z'

# 日志配置
LOG_FILE = '柱子统计分析.log'
LOG_LEVEL = 'INFO'  # 可选: DEBUG, INFO, WARNING, ERROR

# 坐标容差配置（用于处理轴网坐标的微小差异）
COORDINATE_TOLERANCE = 0.1
