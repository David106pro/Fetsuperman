#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DWG转DXF专用转换工具

提供多种转换方案：
1. AutoCAD自动转换
2. 手动转换指导
3. 在线转换服务推荐
4. 免费CAD软件推荐
"""

import os
import sys
import webbrowser
from pathlib import Path

def show_title():
    """显示标题"""
    print("=" * 60)
    print("             DWG转DXF格式转换工具")
    print("=" * 60)
    print()

def check_file_exists(file_path):
    """检查文件是否存在"""
    if not os.path.exists(file_path):
        print(f"❌ 错误：文件不存在")
        print(f"   文件路径: {file_path}")
        print("\n请检查文件路径是否正确")
        return False
    return True

def show_autocad_guide(dwg_path, dxf_path):
    """显示AutoCAD转换指导"""
    print("🔧 方案一：使用AutoCAD转换（推荐）")
    print("-" * 40)
    print("如果您安装了AutoCAD，请按以下步骤操作：")
    print()
    print("1. 启动AutoCAD")
    print("2. 打开DWG文件：")
    print(f"   文件 → 打开 → {dwg_path}")
    print("3. 另存为DXF：")
    print("   文件 → 另存为")
    print("   文件类型：AutoCAD DXF (*.dxf)")
    print(f"   保存位置：{dxf_path}")
    print("4. 点击保存")
    print()

def show_free_cad_software():
    """显示免费CAD软件推荐"""
    print("🆓 方案二：使用免费CAD软件")
    print("-" * 40)
    
    software_list = [
        {
            "name": "DraftSight",
            "url": "https://www.3ds.com/products-services/draftsight-cad-software/",
            "description": "专业2D CAD软件，支持DWG/DXF格式"
        },
        {
            "name": "LibreCAD", 
            "url": "https://librecad.org/",
            "description": "开源免费2D CAD软件"
        },
        {
            "name": "QCAD",
            "url": "https://qcad.org/",
            "description": "跨平台2D CAD软件"
        },
        {
            "name": "BricsCAD Shape",
            "url": "https://www.bricsys.com/shape/",
            "description": "免费3D建模软件，支持DWG格式"
        }
    ]
    
    for i, software in enumerate(software_list, 1):
        print(f"{i}. {software['name']}")
        print(f"   {software['description']}")
        print(f"   下载地址: {software['url']}")
        print()

def show_online_converters():
    """显示在线转换服务"""
    print("🌐 方案三：在线转换服务")
    print("-" * 40)
    
    converters = [
        {
            "name": "Zamzar",
            "url": "https://www.zamzar.com/convert/dwg-to-dxf/",
            "features": "支持多种格式，免费转换"
        },
        {
            "name": "Convertio", 
            "url": "https://convertio.co/dwg-dxf/",
            "features": "在线转换，支持批量处理"
        },
        {
            "name": "Online-Convert",
            "url": "https://www.online-convert.com/",
            "features": "专业在线转换平台"
        },
        {
            "name": "CloudConvert",
            "url": "https://cloudconvert.com/dwg-to-dxf",
            "features": "API支持，高质量转换"
        }
    ]
    
    for i, converter in enumerate(converters, 1):
        print(f"{i}. {converter['name']}")
        print(f"   特点: {converter['features']}")
        print(f"   网址: {converter['url']}")
        print()
    
    print("💡 使用在线转换的步骤：")
    print("   1. 访问上述任一网站")
    print("   2. 上传您的DWG文件")
    print("   3. 选择输出格式为DXF")
    print("   4. 下载转换后的文件")
    print()

def open_online_converter():
    """打开在线转换器"""
    try:
        url = "https://www.zamzar.com/convert/dwg-to-dxf/"
        webbrowser.open(url)
        print(f"✅ 已在浏览器中打开: {url}")
        return True
    except Exception as e:
        print(f"❌ 无法打开浏览器: {e}")
        return False

def show_conversion_tips():
    """显示转换注意事项"""
    print("⚠️  转换注意事项")
    print("-" * 40)
    print("1. 转换前请备份原始DWG文件")
    print("2. 确保DWG文件没有损坏")
    print("3. 转换后检查DXF文件是否完整")
    print("4. 某些复杂对象可能需要手动处理")
    print("5. 建议使用相同版本的CAD软件进行转换")
    print()

def main():
    """主函数"""
    show_title()
    
    # 默认文件路径
    dwg_path = r"E:\code\CursorCode\公司\cursor项目\6.老爹脚本\找柱子\S-2#地下车库结构图.dwg"
    dxf_path = dwg_path.replace('.dwg', '.dxf')
    
    print(f"📁 源文件: {dwg_path}")
    print(f"📁 目标文件: {dxf_path}")
    print()
    
    # 检查文件是否存在
    if not check_file_exists(dwg_path):
        input("按任意键退出...")
        return
    
    # 检查DXF是否已存在
    if os.path.exists(dxf_path):
        print("✅ DXF文件已存在！")
        print(f"   文件位置: {dxf_path}")
        print("\n可以直接运行柱子分析脚本了")
        input("按任意键退出...")
        return
    
    # 显示转换方案
    show_autocad_guide(dwg_path, dxf_path)
    show_free_cad_software()
    show_online_converters()
    show_conversion_tips()
    
    # 用户选择
    print("🎯 请选择转换方案：")
    print("1. 我有AutoCAD，按照方案一手动转换")
    print("2. 下载免费CAD软件进行转换")
    print("3. 使用在线转换服务")
    print("4. 直接打开在线转换网站")
    print("5. 退出")
    print()
    
    while True:
        try:
            choice = input("请输入选项 (1-5): ").strip()
            
            if choice == '1':
                print("\n✅ 请按照方案一的步骤在AutoCAD中完成转换")
                print("转换完成后，运行主脚本进行柱子分析")
                break
                
            elif choice == '2':
                print("\n✅ 请从方案二中选择合适的免费CAD软件")
                print("安装后按照类似步骤完成转换")
                break
                
            elif choice == '3':
                print("\n✅ 请从方案三中选择在线转换服务")
                print("上传DWG文件并下载转换后的DXF文件")
                break
                
            elif choice == '4':
                print("\n🌐 正在打开在线转换网站...")
                open_online_converter()
                break
                
            elif choice == '5':
                print("\n👋 退出转换工具")
                break
                
            else:
                print("❌ 无效选项，请输入1-5之间的数字")
                
        except KeyboardInterrupt:
            print("\n\n👋 用户取消操作")
            break
        except Exception as e:
            print(f"❌ 输入错误: {e}")
    
    print("\n" + "=" * 60)
    print("转换完成后，请运行 main.py 进行柱子分析")
    print("=" * 60)
    input("\n按任意键退出...")

if __name__ == "__main__":
    main()
