#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
北京运营表格生成工具 - 环境配置脚本
自动安装依赖包、配置环境、设置默认路径
"""

import os
import sys
import subprocess
import platform
import json
from pathlib import Path

def print_banner():
    """打印程序横幅"""
    print("=" * 60)
    print("北京运营表格生成工具 - 环境配置脚本")
    print("=" * 60)
    print()

def check_python_version():
    """检查Python版本"""
    print("🔍 检查Python版本...")
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"❌ Python版本过低: {version.major}.{version.minor}")
        print("   需要Python 3.8或更高版本")
        return False
    else:
        print(f"✅ Python版本: {version.major}.{version.minor}.{version.micro}")
        return True

def install_package(package):
    """安装Python包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False

def install_required_packages():
    """安装必需的Python包"""
    print("\n📦 安装必需的Python包...")
    
    required_packages = [
        "pandas>=1.5.0",
        "openpyxl>=3.0.0",
        "psutil>=5.8.0"
    ]
    
    success_count = 0
    for package in required_packages:
        print(f"   正在安装 {package}...")
        if install_package(package):
            print(f"   ✅ {package} 安装成功")
            success_count += 1
        else:
            print(f"   ❌ {package} 安装失败")
    
    print(f"\n📊 安装结果: {success_count}/{len(required_packages)} 个包安装成功")
    return success_count == len(required_packages)

def get_desktop_path():
    """获取桌面路径"""
    if platform.system() == "Windows":
        return os.path.join(os.path.expanduser("~"), "Desktop")
    else:
        return os.path.join(os.path.expanduser("~"), "Desktop")

def create_default_config():
    """创建默认配置文件"""
    print("\n⚙️ 创建默认配置文件...")
    
    desktop_path = get_desktop_path()
    
    # 创建setting目录
    setting_dir = Path("setting")
    setting_dir.mkdir(exist_ok=True)
    
    # 默认配置
    default_config = {
        "总表路径": os.path.join(desktop_path, "【北京OTT总表】提审&发上线总表.xlsx"),
        "专辑注入路径": os.path.join(desktop_path, "北京项目", "北京专辑注入.csv"),
        "剧集注入路径": os.path.join(desktop_path, "北京项目", "北京子集注入.csv"),
        "审核表路径": os.path.join(desktop_path, "北京项目", "3、审核结果"),
        "送审表输出路径": os.path.join(desktop_path, "北京项目", "1、送审"),
        "上线表输出路径": os.path.join(desktop_path, "北京项目", "2、上线"),
        "是否备份总表": False,
        "审核状态格式统一": False
    }
    
    # 保存配置文件
    config_file = setting_dir / "config.json"
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print(f"   ✅ 配置文件已创建: {config_file}")
        return True
    except Exception as e:
        print(f"   ❌ 配置文件创建失败: {e}")
        return False

def create_directories():
    """创建必要的目录"""
    print("\n📁 创建必要的目录...")
    
    desktop_path = get_desktop_path()
    directories = [
        os.path.join(desktop_path, "北京项目"),
        os.path.join(desktop_path, "北京项目", "1、送审"),
        os.path.join(desktop_path, "北京项目", "2、上线"),
        os.path.join(desktop_path, "北京项目", "3、审核结果"),
        "logs"
    ]
    
    success_count = 0
    for directory in directories:
        try:
            Path(directory).mkdir(parents=True, exist_ok=True)
            print(f"   ✅ 目录已创建: {directory}")
            success_count += 1
        except Exception as e:
            print(f"   ❌ 目录创建失败 {directory}: {e}")
    
    print(f"\n📊 目录创建结果: {success_count}/{len(directories)} 个目录创建成功")
    return success_count == len(directories)

def test_imports():
    """测试导入必需的包"""
    print("\n🧪 测试包导入...")
    
    required_modules = [
        ("pandas", "数据处理"),
        ("openpyxl", "Excel文件操作"),
        ("psutil", "进程管理"),
        ("tkinter", "图形界面")
    ]
    
    success_count = 0
    for module, description in required_modules:
        try:
            __import__(module)
            print(f"   ✅ {module} ({description}) 导入成功")
            success_count += 1
        except ImportError as e:
            print(f"   ❌ {module} ({description}) 导入失败: {e}")
    
    print(f"\n📊 导入测试结果: {success_count}/{len(required_modules)} 个模块导入成功")
    return success_count == len(required_modules)

def create_sample_files():
    """创建示例文件"""
    print("\n📄 创建示例文件...")
    
    desktop_path = get_desktop_path()
    
    # 创建示例总表
    try:
        import pandas as pd
        from openpyxl import Workbook
        
        # 创建示例总表
        total_table_path = os.path.join(desktop_path, "【北京OTT总表】提审&发上线总表.xlsx")
        if not os.path.exists(total_table_path):
            # 创建示例数据
            album_data = {
                '专辑ID': ['ALB001', 'ALB002'],
                '专辑名称': ['示例专辑1', '示例专辑2'],
                '审核状态（专辑）': ['待提审', '审核中'],
                '送审时间': ['', 'T']
            }
            
            episode_data = {
                '专辑ID': ['ALB001', 'ALB001'],
                '剧集ID': ['EP001', 'EP002'],
                '剧集名称': ['示例剧集1', '示例剧集2'],
                '审核状态（剧集）': ['待提审', '审核中'],
                '申请上线时间': ['', 'T']
            }
            
            with pd.ExcelWriter(total_table_path, engine='openpyxl') as writer:
                pd.DataFrame(album_data).to_excel(writer, sheet_name='专辑', index=False)
                pd.DataFrame(episode_data).to_excel(writer, sheet_name='剧集', index=False)
            
            print(f"   ✅ 示例总表已创建: {total_table_path}")
        else:
            print(f"   ℹ️ 总表已存在: {total_table_path}")
        
        return True
    except Exception as e:
        print(f"   ❌ 示例文件创建失败: {e}")
        return False

def print_completion_message():
    """打印完成消息"""
    print("\n" + "=" * 60)
    print("🎉 环境配置完成！")
    print("=" * 60)
    print()
    print("📋 配置摘要:")
    print("   ✅ Python环境检查通过")
    print("   ✅ 必需包安装完成")
    print("   ✅ 配置文件已创建")
    print("   ✅ 目录结构已建立")
    print("   ✅ 示例文件已创建")
    print()
    print("🚀 下一步操作:")
    print("   1. 将您的数据文件放入相应目录")
    print("   2. 运行 'python main.py' 启动程序")
    print("   3. 在设置中调整路径配置")
    print()
    print("📁 重要目录:")
    desktop_path = get_desktop_path()
    print(f"   总表位置: {desktop_path}")
    print(f"   输入表格: {desktop_path}/北京项目")
    print(f"   输出表格: {desktop_path}/北京项目/1、送审")
    print(f"   上线表格: {desktop_path}/北京项目/2、上线")
    print()
    print("💡 提示:")
    print("   - 首次使用建议先查看README.md文档")
    print("   - 如有问题请检查logs目录下的日志文件")
    print("   - 可在设置中自定义所有路径配置")
    print()

def main():
    """主函数"""
    print_banner()
    
    # 检查Python版本
    if not check_python_version():
        print("\n❌ 环境检查失败，请升级Python版本后重试")
        input("按回车键退出...")
        return
    
    # 安装必需包
    if not install_required_packages():
        print("\n❌ 包安装失败，请检查网络连接后重试")
        input("按回车键退出...")
        return
    
    # 测试导入
    if not test_imports():
        print("\n❌ 包导入测试失败，请检查安装是否正确")
        input("按回车键退出...")
        return
    
    # 创建配置文件
    create_default_config()
    
    # 创建目录
    create_directories()
    
    # 创建示例文件
    create_sample_files()
    
    # 打印完成消息
    print_completion_message()
    
    input("按回车键退出...")

if __name__ == "__main__":
    main() 