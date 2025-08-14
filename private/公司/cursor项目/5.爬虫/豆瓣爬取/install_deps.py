import subprocess
import sys

def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

packages = ['requests', 'beautifulsoup4', 'openpyxl']
for package in packages:
    print(f"正在安装 {package}...")
    install_package(package)
print("所有依赖安装完成！")