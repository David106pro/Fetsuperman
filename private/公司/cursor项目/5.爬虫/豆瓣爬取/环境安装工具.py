# -*- coding: utf-8 -*-
"""
快速安装脚本
自动安装豆瓣爬虫所需的所有依赖
"""

import subprocess
import sys
import os
import platform
import requests
import zipfile
from pathlib import Path


def install_package(package_name):
    """安装Python包"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        print(f"✓ 成功安装: {package_name}")
        return True
    except subprocess.CalledProcessError:
        print(f"✗ 安装失败: {package_name}")
        return False


def check_chrome_installed():
    """检查Chrome是否已安装"""
    chrome_paths = {
        'Windows': [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ],
        'Darwin': [  # macOS
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        ],
        'Linux': [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/chromium-browser"
        ]
    }
    
    system = platform.system()
    paths = chrome_paths.get(system, [])
    
    for path in paths:
        if os.path.exists(path):
            print(f"✓ 发现Chrome浏览器: {path}")
            return True
    
    print("✗ 未发现Chrome浏览器")
    return False


def get_chrome_version():
    """获取Chrome版本"""
    try:
        if platform.system() == "Windows":
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return version
        else:
            # Linux/macOS
            result = subprocess.run(["google-chrome", "--version"], capture_output=True, text=True)
            return result.stdout.split()[-1]
    except:
        return None


def download_chromedriver():
    """下载ChromeDriver"""
    try:
        print("正在检测Chrome版本...")
        chrome_version = get_chrome_version()
        
        if not chrome_version:
            print("⚠ 无法检测Chrome版本，使用最新版本ChromeDriver")
            # 获取最新版本
            response = requests.get("https://chromedriver.chromium.org/downloads")
            # 这里可以解析页面获取最新版本，简化处理使用固定版本
            version = "126.0.6478.126"
        else:
            version = chrome_version.split('.')[0]  # 主版本号
            print(f"检测到Chrome版本: {chrome_version}")
            print(f"使用ChromeDriver版本: {version}")
        
        # 确定下载URL
        system = platform.system()
        if system == "Windows":
            if platform.architecture()[0] == "64bit":
                driver_url = f"https://chromedriver.storage.googleapis.com/{version}.0.6478.126/chromedriver_win32.zip"
                driver_name = "chromedriver.exe"
            else:
                driver_url = f"https://chromedriver.storage.googleapis.com/{version}.0.6478.126/chromedriver_win32.zip"
                driver_name = "chromedriver.exe"
        elif system == "Darwin":  # macOS
            driver_url = f"https://chromedriver.storage.googleapis.com/{version}.0.6478.126/chromedriver_mac64.zip"
            driver_name = "chromedriver"
        else:  # Linux
            driver_url = f"https://chromedriver.storage.googleapis.com/{version}.0.6478.126/chromedriver_linux64.zip"
            driver_name = "chromedriver"
        
        print(f"正在下载ChromeDriver: {driver_url}")
        
        # 下载文件
        response = requests.get(driver_url, stream=True)
        response.raise_for_status()
        
        zip_path = "chromedriver.zip"
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        # 解压文件
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall('.')
        
        # 删除zip文件
        os.remove(zip_path)
        
        # 检查文件是否存在
        if os.path.exists(driver_name):
            # 如果是Linux/macOS，添加执行权限
            if system != "Windows":
                os.chmod(driver_name, 0o755)
            
            print(f"✓ ChromeDriver下载成功: {os.path.abspath(driver_name)}")
            return True
        else:
            print("✗ ChromeDriver下载失败")
            return False
            
    except Exception as e:
        print(f"✗ 下载ChromeDriver时出错: {e}")
        return False


def main():
    """主函数"""
    print("豆瓣爬虫依赖安装程序")
    print("=" * 40)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("✗ 需要Python 3.7或更高版本")
        return
    
    print(f"✓ Python版本: {sys.version}")
    
    # 必需的包列表
    required_packages = [
        "requests",
        "beautifulsoup4",
        "openpyxl",
        "selenium"
    ]
    
    print("\n正在安装Python依赖包...")
    success_count = 0
    
    for package in required_packages:
        if install_package(package):
            success_count += 1
    
    print(f"\nPython包安装完成: {success_count}/{len(required_packages)}")
    
    if success_count < len(required_packages):
        print("⚠ 部分包安装失败，请手动安装")
        return
    
    # 检查Chrome
    print("\n检查Chrome浏览器...")
    if not check_chrome_installed():
        print("⚠ 请先安装Chrome浏览器:")
        print("  Windows: https://www.google.com/chrome/")
        print("  macOS: https://www.google.com/chrome/")
        print("  Linux: sudo apt-get install google-chrome-stable")
        
        user_input = input("\n是否已安装Chrome? (y/n): ").lower()
        if user_input != 'y':
            print("请安装Chrome后重新运行此脚本")
            return
    
    # 下载ChromeDriver
    print("\n正在下载ChromeDriver...")
    if download_chromedriver():
        print("✓ ChromeDriver安装成功")
    else:
        print("✗ ChromeDriver安装失败")
        print("请手动下载ChromeDriver:")
        print("  1. 访问: https://chromedriver.chromium.org/downloads")
        print("  2. 下载与Chrome版本匹配的ChromeDriver")
        print("  3. 将chromedriver.exe放到项目目录或PATH路径中")
    
    print("\n" + "=" * 40)
    print("安装完成！")
    print("\n使用说明:")
    print("1. 运行主程序: python test.py")
    print("2. 测试Selenium: python douban_selenium.py")
    print("3. 测试代理: python proxy_test.py")
    print("4. 查看安装说明: 安装说明.md")
    
    print("\n如果遇到问题，请检查:")
    print("- Chrome浏览器是否正确安装")
    print("- ChromeDriver版本是否与Chrome版本匹配")
    print("- 网络连接是否正常")


if __name__ == "__main__":
    main() 