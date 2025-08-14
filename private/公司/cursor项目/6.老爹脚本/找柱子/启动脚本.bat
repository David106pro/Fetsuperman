@echo off
chcp 65001 >nul
title 柱子区域统计分析工具

echo.
echo ===============================================
echo           建筑柱子区域统计分析工具
echo ===============================================
echo.
echo 请选择功能：
echo 1. DWG转DXF格式转换工具
echo 2. 柱子区域统计分析（需要DXF文件）
echo 3. 退出
echo.

set /p choice="请输入选项 (1-3): "

if "%choice%"=="1" (
    echo.
    echo 正在启动DWG转换工具...
    python "DWG转换工具.py"
) else if "%choice%"=="2" (
    echo.
    echo 正在启动柱子分析脚本...
    python main.py
) else if "%choice%"=="3" (
    echo 退出程序
    exit /b
) else (
    echo 无效选项，请重新运行
)

echo.
pause
