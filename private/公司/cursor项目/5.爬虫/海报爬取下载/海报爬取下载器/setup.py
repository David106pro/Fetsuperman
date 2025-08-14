#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海报爬取下载器 - 环境配置脚本
自动安装依赖、初始化配置、生成标准项目结构
"""
import sys
import subprocess
import os
import json
import shutil
import logging
from datetime import datetime

# 所有必要的依赖包 - 基于mainPro.py实际使用的依赖更新
REQUIRED_PACKAGES = [
    "customtkinter>=5.2.0",    # GUI框架
    "requests>=2.31.0",        # HTTP请求
    "pillow>=10.0.0",          # 图像处理 (包含PIL模块)
    "pandas>=2.0.0",           # 数据处理
    "openpyxl>=3.1.0",         # Excel操作
    "beautifulsoup4>=4.12.2",  # HTML解析 (bs4)
    "lxml>=4.9.0",             # XML解析器
    "urllib3>=2.0.0",          # URL处理
]

def print_status(message, status="INFO"):
    """打印带状态的消息"""
    status_icons = {
        "INFO": "ℹ️",
        "SUCCESS": "✅",
        "ERROR": "❌",
        "WARNING": "⚠️",
        "PROGRESS": "🔄"
    }
    print(f"{status_icons.get(status, 'ℹ️')} {message}")

def check_python_version():
    """检查Python版本"""
    if sys.version_info < (3, 8):
        print_status("Python版本过低，需要Python 3.8及以上版本", "ERROR")
        print_status(f"当前版本: {sys.version}", "INFO")
        return False
    print_status(f"Python版本检查通过: {sys.version.split()[0]}", "SUCCESS")
    return True

def install_package(package):
    """安装单个包"""
    try:
        print_status(f"正在安装 {package}...", "PROGRESS")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", 
            package, "--upgrade", "--quiet"
        ])
        print_status(f"{package} 安装成功", "SUCCESS")
        return True
    except subprocess.CalledProcessError as e:
        print_status(f"{package} 安装失败: {e}", "ERROR")
        return False
    except Exception as e:
        print_status(f"{package} 安装出现异常: {e}", "ERROR")
        return False

def check_and_install_dependencies():
    """检查并安装所有依赖"""
    print_status("开始检查和安装依赖包...", "INFO")
    
    failed_packages = []
    
    for package in REQUIRED_PACKAGES:
        package_name = package.split(">=")[0]  # 获取包名，去掉版本号
        try:
            # 特殊处理某些包名映射
            import_name = package_name
            if package_name == "pillow":
                import_name = "PIL"
            elif package_name == "beautifulsoup4":
                import_name = "bs4"
            
            __import__(import_name)
            print_status(f"{package_name} 已安装", "SUCCESS")
        except ImportError:
            if not install_package(package):
                failed_packages.append(package)
    
    if failed_packages:
        print_status(f"以下包安装失败: {', '.join(failed_packages)}", "ERROR")
        return False
    
    print_status("所有依赖包安装完成", "SUCCESS")
    return True

def get_desktop_path():
    """获取桌面路径"""
    return os.path.join(os.path.expanduser("~"), "Desktop")

def create_default_config():
    """创建默认配置 - 与mainPro.py中的get_default_settings()保持同步"""
    return {
        "default_platform": "爱奇艺",
        "default_precise": False,
        "default_download_type": "全部",
        "default_path": get_desktop_path(),
        "default_poster_size": "原尺寸",  # 更新为原尺寸，与UI配色优化后的默认值一致
        "default_vertical_size": "412x600",
        "default_horizontal_size": "528x296",
        "filename_format": "{标题}_{类型}_{图片尺寸}",
        "batch_search_priority": [
            "优酷视频-精确搜索", 
            "爱奇艺-精确搜索", 
            "爱奇艺-普通搜索"
        ],
        "batch_horizontal_path": os.path.join(get_desktop_path(), "横图"),
        "batch_vertical_path": os.path.join(get_desktop_path(), "竖图"),
        "delete_horizontal_path": os.path.join(get_desktop_path(), "横图"),  # 删除功能路径配置
        "delete_vertical_path": os.path.join(get_desktop_path(), "竖图"),   # 删除功能路径配置
        "batch_default_size": "原尺寸",  # 更新为原尺寸
        "batch_default_vertical_size": "412x600",
        "batch_default_horizontal_size": "528x296",
        "iqiyi_cookie": "",  # 爱奇艺Cookie设置
        "tencent_cookie": "", # 腾讯视频Cookie设置
        "youku_cookie": ""   # 优酷视频Cookie设置
    }

def create_project_structure():
    """创建标准项目结构"""
    print_status("开始创建标准项目结构...", "INFO")
    
    # 项目根目录
    project_root = "海报爬取下载器"
    
    try:
        # 如果目标目录已存在，先备份
        if os.path.exists(project_root):
            backup_name = f"{project_root}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            print_status(f"目录已存在，备份为: {backup_name}", "WARNING")
            shutil.move(project_root, backup_name)
        
        # 创建目录结构
        os.makedirs(project_root, exist_ok=True)
        os.makedirs(os.path.join(project_root, "setting"), exist_ok=True)
        
        # 复制主程序文件 - 优先使用mainPro.py
        if os.path.exists("mainPro.py"):
            shutil.copy2("mainPro.py", os.path.join(project_root, "poster_downloader.py"))
            print_status("主程序文件已复制: poster_downloader.py (来源: mainPro.py)", "SUCCESS")
        elif os.path.exists("main.py"):
            shutil.copy2("main.py", os.path.join(project_root, "poster_downloader.py"))
            print_status("主程序文件已复制: poster_downloader.py (来源: main.py)", "SUCCESS")
        else:
            print_status("未找到主程序文件 (mainPro.py 或 main.py)", "WARNING")
        
        # 创建配置文件
        config_path = os.path.join(project_root, "setting", "config.json")
        config = create_default_config()
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print_status("配置文件已创建: setting/config.json", "SUCCESS")
        
        # 创建项目文档
        doc_content = create_project_documentation()
        doc_path = os.path.join(project_root, "海报爬取下载器文档.md")
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_content)
        print_status("项目文档已创建: 海报爬取下载器文档.md", "SUCCESS")
        
        # 复制setup.py
        if os.path.exists("setup.py"):
            shutil.copy2("setup.py", os.path.join(project_root, "setup.py"))
            print_status("环境配置脚本已复制: setup.py", "SUCCESS")
        
        # 创建日志文件
        log_path = os.path.join(project_root, "project-logs.log")
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"# 海报爬取下载器 - 项目日志\n")
            f.write(f"项目创建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Python版本: {sys.version}\n")
            f.write(f"依赖包: {', '.join([pkg.split('>=')[0] for pkg in REQUIRED_PACKAGES])}\n\n")
        print_status("日志文件已创建: project-logs.log", "SUCCESS")
        
        # 创建README文件（可选）
        readme_content = create_readme_content()
        readme_path = os.path.join(project_root, "README.md")
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(readme_content)
        print_status("README文件已创建: README.md", "SUCCESS")
        
        print_status(f"项目结构创建完成: {project_root}/", "SUCCESS")
        print_status("目录结构:", "INFO")
        print(f"""
{project_root}/
├── poster_downloader.py      # 海报爬取下载器主程序
├── 海报爬取下载器文档.md        # 项目文档
├── setup.py                  # 环境配置脚本
├── README.md                 # 使用说明
├── setting/                  # 配置文件夹
│   └── config.json          # 配置文件
└── project-logs.log         # 日志文件
        """)
        
        return project_root
        
    except Exception as e:
        print_status(f"创建项目结构失败: {e}", "ERROR")
        return None

def create_project_documentation():
    """创建项目文档内容"""
    return """# 海报爬取下载器文档

## 项目概述
海报爬取下载器是一个功能强大的多平台海报图片下载工具，支持从爱奇艺、优酷视频、腾讯视频等主流视频平台爬取和下载海报图片。

## 主要功能

### 🎯 核心功能
- **多平台支持**: 爱奇艺、优酷视频、腾讯视频
- **智能搜索**: 精确搜索和模糊搜索
- **批量处理**: Excel批量导入，自动化处理
- **尺寸预设**: 基础、河南、甘肃、陕西、云南等多种尺寸
- **智能缩放**: 保持比例的智能裁剪
- **自动压缩**: 云南尺寸100KB，其他300KB自动压缩

### 📊 支持的尺寸规格
- **原尺寸**: 保持图片原始尺寸，无缩放处理
- **基础尺寸**: 竖图 412x600px, 横图 528x296px
- **河南尺寸**: 竖图 525x750px, 横图 257x145px  
- **甘肃尺寸**: 竖图 412x600px, 横图 562x375px
- **陕西尺寸**: 竖图 245x350px, 横图 384x216px
- **云南尺寸**: 竖图 262x360px, 横图 412x230px
- **自定义尺寸**: 支持用户自定义

### 🛠️ 高级功能
- **路径同步**: 修改默认路径时自动同步批量和删除路径
- **智能缩放裁剪**: 保持图片比例，居中裁剪
- **批量删除**: 支持Excel导入批量删除不匹配文件
- **Cookie管理**: 支持各平台Cookie配置
- **搜索优先级**: 可配置多平台搜索优先级

## 安装和使用

### 环境要求
- Python 3.8+
- Windows 10/11

### 安装步骤
1. 运行 `python setup.py` 安装依赖
2. 运行 `python poster_downloader.py` 启动程序

### 基本使用
1. 选择平台（爱奇艺/优酷/腾讯）
2. 输入搜索关键词
3. 选择尺寸预设
4. 设置下载路径
5. 点击搜索并下载

### 批量处理
1. 切换到"批量爬取"标签
2. 选择Excel文件（包含影片名称和CID列）
3. 设置横图和竖图保存路径
4. 配置搜索优先级
5. 开始批量处理

## 配置说明

### 基本设置
- **默认平台**: 默认搜索平台
- **默认路径**: 图片保存路径
- **文件名格式**: 支持变量替换

### 批量设置  
- **搜索优先级**: 多平台搜索顺序
- **默认尺寸**: 批量处理默认尺寸
- **路径设置**: 横图和竖图分别保存

### Cookie设置
- **爱奇艺Cookie**: 提高搜索成功率
- **优酷Cookie**: 访问更多内容
- **腾讯Cookie**: 绕过访问限制

## 技术特性

### UI界面优化 🎨
- **浅色模式设计**: 设置浅色外观模式，提供最佳视觉体验
- **高对比度配色**: 深灰色文字(#2F2F2F)搭配浅色背景，显著提高可读性
- **统一配色方案**: 
  - 主窗口：#F0F0F0 浅灰背景
  - 侧边栏：#E0E0E0 中灰背景  
  - 内容区：#FFFFFF 纯白背景
  - 激活按钮：#1F6AA5 蓝色背景 + 白色文字
  - 状态信息：#1F6AA5 蓝色高亮

### 智能缩放裁剪
- 自动选择基准边进行等比缩放
- 居中裁剪，保留图片主要内容
- 避免图片拉伸变形

### 自动压缩
- 云南尺寸: 压缩至100KB以内
- 其他尺寸: 压缩至300KB以内
- 支持JPEG/WEBP格式质量压缩

### 路径同步
- 修改默认路径时自动同步
- 批量爬取路径自动更新
- 批量删除路径自动更新

## 更新日志

### v2025.01.24
- ✅ **UI配色优化**: 全面改善界面文字可读性
  - 设置浅色模式，提高文字对比度
  - 统一配色方案，深灰文字搭配浅色背景
  - 优化按钮状态显示，激活/默认状态区分明显
  - 状态信息使用蓝色高亮，重要信息更突出
- ✅ **依赖包管理优化**: 完善requirements.txt
  - 自动安装所有必需依赖包（PIL、requests、pandas等）
  - 修复ModuleNotFoundError错误
- ✅ **默认配置更新**: 原尺寸为默认选项
- ✅ **项目结构标准化**: setup.py自动配置环境

### v2025.01.22
- ✅ 添加陕西尺寸支持 (245x350px / 384x216px)
- ✅ 实现智能缩放裁剪功能
- ✅ 添加300KB自动压缩
- ✅ 实现路径同步功能
- ✅ 修复批量下载平台匹配问题

## 技术支持
如有问题请查看日志文件 `project-logs.log` 或联系技术支持。

---
*海报爬取下载器 - 让海报下载更简单*
"""

def create_readme_content():
    """创建README内容"""
    return """# 海报爬取下载器

## 快速开始

### 1. 环境配置
```bash
python setup.py
```

### 2. 启动程序
```bash
python poster_downloader.py
```

### 3. 基本使用
1. 选择平台和尺寸
2. 输入搜索关键词
3. 设置保存路径
4. 点击搜索并下载

## 主要功能
- 🎯 多平台海报下载（爱奇艺/优酷/腾讯）
- 🎨 **新增**: 优化UI配色，文字清晰易读
- 📊 多种尺寸预设支持（包含原尺寸选项）
- 🔄 批量Excel处理
- 🛠️ 智能缩放裁剪
- 📁 路径自动同步
- 🗑️ 批量删除功能
- ⚙️ 自动环境配置（setup.py）

## 项目结构
```
海报爬取下载器/
├── poster_downloader.py      # 主程序
├── 海报爬取下载器文档.md        # 详细文档
├── setup.py                  # 环境配置
├── setting/config.json       # 配置文件
└── project-logs.log         # 日志文件
```

## 技术栈
- Python 3.8+
- CustomTkinter (GUI)
- Requests (HTTP)
- Pillow (图像处理)
- Pandas (数据处理)
- BeautifulSoup (HTML解析)

---
详细使用说明请查看 `海报爬取下载器文档.md`
"""

def main():
    """主函数"""
    print("=" * 60)
    print("🎬 海报爬取下载器 - 环境配置和项目初始化")
    print("=" * 60)
    
    # 检查Python版本
    if not check_python_version():
        sys.exit(1)
    
    # 检查并安装依赖
    if not check_and_install_dependencies():
        print_status("依赖安装失败，请手动安装失败的包", "ERROR")
        sys.exit(1)
    
    # 创建项目结构
    project_dir = create_project_structure()
    if not project_dir:
        sys.exit(1)
    
    print("=" * 60)
    print_status("🎉 环境配置和项目初始化完成！", "SUCCESS")
    print("=" * 60)
    
    print(f"""
📁 项目已创建在: {os.path.abspath(project_dir)}

🚀 接下来的步骤:
1. 进入项目目录: cd {project_dir}
2. 启动程序: python poster_downloader.py
3. 查看文档: 海报爬取下载器文档.md

🎨 最新更新 (v2025.01.24):
- ✅ UI配色全面优化，文字清晰易读
- ✅ 修复依赖包问题，程序启动更稳定
- ✅ 默认使用原尺寸，保持图片质量

💡 提示:
- 首次使用建议先查看文档了解功能
- 可在setting/config.json中修改默认配置
- 日志信息会保存在project-logs.log中
- 如遇到问题请查看UI配色优化记录.md
    """)

if __name__ == "__main__":
    main() 