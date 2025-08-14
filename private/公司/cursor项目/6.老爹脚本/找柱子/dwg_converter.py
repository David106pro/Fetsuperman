#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DWG转DXF格式转换器

支持多种转换方式：
1. AutoCAD COM接口自动转换（需要安装AutoCAD）
2. 在线转换服务
3. 手动转换指导
"""

import os
import sys
from pathlib import Path
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DWGConverter:
    """DWG文件转换器"""
    
    def __init__(self):
        self.autocad_available = False
        self.check_autocad()
    
    def check_autocad(self):
        """检查AutoCAD是否可用"""
        try:
            import win32com.client
            try:
                # 尝试连接AutoCAD
                acad = win32com.client.Dispatch("AutoCAD.Application")
                self.autocad_available = True
                logger.info("✅ 检测到AutoCAD，支持自动转换")
                return True
            except:
                logger.info("❌ 未检测到AutoCAD或AutoCAD未启动")
                return False
        except ImportError:
            logger.info("❌ 缺少win32com库，无法使用AutoCAD自动转换")
            return False
    
    def convert_with_autocad(self, dwg_path: str, dxf_path: str = None) -> bool:
        """
        使用AutoCAD COM接口转换DWG为DXF
        
        Args:
            dwg_path: DWG文件路径
            dxf_path: 输出DXF文件路径（可选）
        
        Returns:
            转换是否成功
        """
        if not self.autocad_available:
            logger.error("AutoCAD不可用，无法进行自动转换")
            return False
        
        try:
            import win32com.client
            
            # 设置输出路径
            if dxf_path is None:
                dxf_path = dwg_path.replace('.dwg', '.dxf')
            
            logger.info(f"正在转换: {dwg_path} -> {dxf_path}")
            
            # 连接AutoCAD
            acad = win32com.client.Dispatch("AutoCAD.Application")
            acad.Visible = False  # 隐藏AutoCAD界面
            
            # 打开DWG文件
            doc = acad.Documents.Open(dwg_path)
            
            # 保存为DXF格式
            doc.SaveAs(dxf_path, 12)  # 12 = DXF格式
            
            # 关闭文档
            doc.Close()
            
            logger.info(f"✅ 转换完成: {dxf_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ AutoCAD转换失败: {e}")
            return False
    
    def install_pywin32(self):
        """安装pywin32库"""
        try:
            import subprocess
            logger.info("正在安装pywin32库...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            logger.info("✅ pywin32安装完成，请重新运行脚本")
            return True
        except Exception as e:
            logger.error(f"❌ pywin32安装失败: {e}")
            return False
    
    def show_manual_conversion_guide(self, dwg_path: str):
        """显示手动转换指导"""
        dxf_path = dwg_path.replace('.dwg', '.dxf')
        
        print("\n" + "="*60)
        print("           DWG转DXF手动转换指导")
        print("="*60)
        print(f"\n📁 源文件: {dwg_path}")
        print(f"📁 目标文件: {dxf_path}")
        
        print("\n🔧 转换步骤：")
        print("1. 打开AutoCAD（或其他支持的CAD软件）")
        print("2. 打开DWG文件：")
        print(f"   文件 → 打开 → 选择: {dwg_path}")
        print("3. 另存为DXF格式：")
        print("   文件 → 另存为")
        print("   文件类型选择: AutoCAD DXF (*.dxf)")
        print(f"   保存位置: {dxf_path}")
        print("4. 点击保存")
        
        print("\n💡 支持的CAD软件：")
        print("- AutoCAD (推荐)")
        print("- AutoCAD LT")
        print("- BricsCAD")
        print("- DraftSight")
        print("- LibreCAD")
        print("- QCAD")
        
        print("\n🌐 在线转换工具（备选方案）：")
        print("- https://www.zamzar.com/convert/dwg-to-dxf/")
        print("- https://convertio.co/dwg-dxf/")
        print("- https://www.online-convert.com/")
        
        print("\n" + "="*60)
    
    def convert(self, dwg_path: str, dxf_path: str = None, auto_install: bool = True) -> str:
        """
        转换DWG文件为DXF格式
        
        Args:
            dwg_path: DWG文件路径
            dxf_path: 输出DXF文件路径
            auto_install: 是否自动安装依赖
        
        Returns:
            转换后的DXF文件路径或None
        """
        # 检查DWG文件是否存在
        if not os.path.exists(dwg_path):
            logger.error(f"❌ DWG文件不存在: {dwg_path}")
            return None
        
        # 设置输出路径
        if dxf_path is None:
            dxf_path = dwg_path.replace('.dwg', '.dxf')
        
        # 检查DXF文件是否已存在
        if os.path.exists(dxf_path):
            logger.info(f"✅ DXF文件已存在: {dxf_path}")
            return dxf_path
        
        # 尝试自动转换
        if not self.autocad_available:
            # 如果没有AutoCAD，尝试安装依赖并重新检查
            if auto_install:
                print("🔧 检测到缺少AutoCAD支持，正在安装依赖...")
                if self.install_pywin32():
                    # 重新检查AutoCAD
                    self.check_autocad()
            
            # 如果还是没有AutoCAD，显示手动转换指导
            if not self.autocad_available:
                print("ℹ️  未检测到AutoCAD，将使用手动转换模式")
                self.show_manual_conversion_guide(dwg_path)
        
        # 如果可以自动转换
        if self.autocad_available:
            print("🔄 正在使用AutoCAD自动转换...")
            if self.convert_with_autocad(dwg_path, dxf_path):
                return dxf_path
            else:
                print("⚠️  自动转换失败，切换到手动转换模式")
                self.show_manual_conversion_guide(dwg_path)
        
        # 等待用户手动转换
        print(f"\n⏳ 请按照上述步骤完成转换，然后按任意键继续...")
        input("转换完成后按回车键继续...")
        
        # 检查转换结果
        if os.path.exists(dxf_path):
            logger.info(f"✅ 检测到转换后的DXF文件: {dxf_path}")
            return dxf_path
        else:
            logger.error(f"❌ 未找到转换后的DXF文件: {dxf_path}")
            return None


def main():
    """测试转换功能"""
    print("DWG转DXF格式转换器")
    print("="*40)
    
    # 测试文件路径
    dwg_file = r"E:\code\CursorCode\公司\cursor项目\6.老爹脚本\找柱子\S-2#地下车库结构图.dwg"
    
    converter = DWGConverter()
    result = converter.convert(dwg_file)
    
    if result:
        print(f"🎉 转换成功！DXF文件路径: {result}")
    else:
        print("❌ 转换失败")


if __name__ == "__main__":
    main()
