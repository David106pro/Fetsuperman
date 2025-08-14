@echo off
chcp 65001 >nul
title 豆瓣爬虫启动器

echo.
echo ==========================================
echo          豆瓣影视媒资信息爬取工具
echo          增强版 v2.0 - 启动器
echo ==========================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python未安装或未添加到PATH环境变量
    echo 请先安装Python 3.8+
    pause
    exit /b 1
)

echo ✅ Python环境检查通过

echo.
echo 正在检查依赖包...
python -c "import requests, bs4, openpyxl, selenium, tkinter" >nul 2>&1
if errorlevel 1 (
    echo ❌ 缺少必要的依赖包
    echo 正在尝试安装...
    python 环境安装工具.py
    if errorlevel 1 (
        echo ❌ 依赖包安装失败，请手动运行: python 环境安装工具.py
        pause
        exit /b 1
    )
)

echo ✅ 依赖包检查通过

echo.
echo 🚀 启动豆瓣爬虫主程序...
echo.
echo 功能特性:
echo • 智能反爬检测与规避
echo • GUI界面，操作简单
echo • 支持Excel批量处理
echo • 实时进度显示
echo • 自动保存和断点续传
echo.

python test.py

echo.
echo 程序已退出，按任意键关闭窗口...
pause >nul 