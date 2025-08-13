#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
环境配置脚本：自动安装依赖、初始化配置
"""
import sys
import subprocess
import os
import json

REQUIRED_PACKAGES = [
    "customtkinter",
    "requests",
    "pillow",
    "pandas",
    "openpyxl",
    "bs4"
]

def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"✅ {package} 安装成功")
    except Exception as e:
        print(f"❌ {package} 安装失败: {e}")

def check_and_install():
    for pkg in REQUIRED_PACKAGES:
        try:
            __import__(pkg if pkg != "pillow" else "PIL")
        except ImportError:
            print(f"未检测到 {pkg}，正在安装...")
            install_package(pkg)

def get_desktop_path():
    # 获取桌面路径
    return os.path.join(os.path.expanduser("~"), "Desktop")

def init_config():
    config_dir = os.path.join(os.path.dirname(__file__), "setting")
    os.makedirs(config_dir, exist_ok=True)
    config_path = os.path.join(config_dir, "config.json")
    if not os.path.exists(config_path):
        config = {
            "default_download_path": get_desktop_path(),
            "filename_format": "{标题}_{类型}_{图片尺寸}"
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"✅ 配置文件已创建: {config_path}")
    else:
        print(f"配置文件已存在: {config_path}")

def main():
    print("==== 环境配置开始 ====")
    if sys.version_info < (3, 7):
        print("❌ Python版本过低，请升级到3.7及以上")
        sys.exit(1)
    check_and_install()
    init_config()
    print("==== 环境配置完成 ====")

if __name__ == "__main__":
    main() 