#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
建筑DWG文件柱子区域统计分析脚本

功能描述：
- 读取DWG文件并分析建筑图纸
- 根据预定义的11个区域规则统计每个区域内的柱子数量
- 自动识别轴网坐标并匹配柱子位置

作者：AI助手
创建时间：2025-01-27
版本：1.0
"""

import ezdxf
import re
import os
import sys
from typing import Dict, List, Tuple, Optional
import logging
from config import DXF_FILE_PATH, REGIONS_CONFIG, COLUMN_IDENTIFIER, LOG_FILE, LOG_LEVEL

# 配置日志
log_level = getattr(logging, LOG_LEVEL.upper(), logging.INFO)
logging.basicConfig(
    level=log_level,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class ColumnAnalyzer:
    """柱子区域统计分析器"""
    
    def __init__(self):
        """初始化分析器"""
        self.axis_coords = {}  # 轴网坐标字典
        self.columns = []      # 柱子坐标列表
        self.regions = []      # 区域定义列表
        self.region_counts = {}  # 区域柱子计数
        
        # 定义11个预设区域
        self._define_regions()
        
    def _define_regions(self):
        """定义区域的规则（从配置文件读取）"""
        logger.info("初始化区域定义...")
        
        # 从配置文件读取区域定义
        self.regions = REGIONS_CONFIG.copy()
        
        logger.info(f"已从配置文件加载 {len(self.regions)} 个区域")
        
    def parse_range(self, range_str: str) -> Tuple[str, str]:
        """
        解析范围字符串
        
        Args:
            range_str: 范围字符串，如 '17-25' 或 'AA-N'
            
        Returns:
            (start, end): 范围的起始和结束值
        """
        if '-' in range_str:
            parts = range_str.split('-')
            if len(parts) == 2:
                return parts[0].strip(), parts[1].strip()
        
        logger.warning(f"无法解析范围字符串: {range_str}")
        return '', ''
    
    def is_numeric(self, text: str) -> bool:
        """判断文本是否为纯数字"""
        return text.strip().isdigit()
    
    def is_alphabetic(self, text: str) -> bool:
        """判断文本是否为纯字母"""
        return text.strip().isalpha()
    
    def load_dwg_file(self, file_path: str) -> Optional[ezdxf.document.Drawing]:
        """
        加载DWG/DXF文件
        
        Args:
            file_path: DWG或DXF文件路径
            
        Returns:
            ezdxf文档对象或None
            
        """
        try:
            logger.info(f"正在加载文件: {file_path}")
            
            if not os.path.exists(file_path):
                logger.error(f"文件不存在: {file_path}")
                return None
            
            # 检查文件扩展名
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext == '.dwg':
                logger.error("检测到DWG文件！")
                logger.error("ezdxf库只能读取DXF格式文件，不支持DWG格式。")
                logger.error("请使用AutoCAD或其他CAD软件将DWG文件另存为DXF格式，然后重新运行脚本。")
                logger.error("转换步骤：")
                logger.error("1. 在AutoCAD中打开DWG文件")
                logger.error("2. 选择 文件 -> 另存为")
                logger.error("3. 在格式下拉菜单中选择 'AutoCAD DXF (*.dxf)'")
                logger.error("4. 保存文件并使用新的DXF文件路径运行本脚本")
                return None
            
            # 尝试读取DXF文件
            doc = ezdxf.readfile(file_path)
            logger.info(f"成功加载DXF文件，版本: {doc.dxfversion}")
            return doc
            
        except ezdxf.DXFStructureError as e:
            logger.error(f"DXF文件结构错误: {e}")
        except ezdxf.DXFValueError as e:
            logger.error(f"DXF文件值错误: {e}")
        except Exception as e:
            logger.error(f"加载文件时发生未知错误: {e}")
            if "not a DXF file" in str(e):
                logger.error("这可能是一个DWG文件或损坏的DXF文件。")
                logger.error("请确保文件是有效的DXF格式。")
        
        return None
    
    def extract_axis_coordinates(self, doc: ezdxf.document.Drawing):
        """
        提取轴网坐标
        
        Args:
            doc: ezdxf文档对象
        """
        logger.info("开始提取轴网坐标...")
        
        self.axis_coords = {}
        modelspace = doc.modelspace()
        
        # 遍历所有文本实体
        for entity in modelspace:
            if entity.dxftype() in ['TEXT', 'MTEXT']:
                try:
                    text_content = entity.dxf.text.strip()
                    insert_point = entity.dxf.insert
                    
                    # 处理纯数字轴网（竖向轴网，记录X坐标）
                    if self.is_numeric(text_content):
                        axis_num = text_content
                        if axis_num not in self.axis_coords:
                            self.axis_coords[axis_num] = {'x': insert_point.x}
                            logger.debug(f"发现数字轴网: {axis_num} at X={insert_point.x}")
                    
                    # 处理字母轴网（横向轴网，记录Y坐标）
                    elif self.is_alphabetic(text_content):
                        axis_letter = text_content.upper()
                        if axis_letter not in self.axis_coords:
                            self.axis_coords[axis_letter] = {'y': insert_point.y}
                            logger.debug(f"发现字母轴网: {axis_letter} at Y={insert_point.y}")
                        
                except AttributeError:
                    # 某些实体可能没有text属性
                    continue
                except Exception as e:
                    logger.warning(f"处理文本实体时出错: {e}")
                    continue
        
        logger.info(f"轴网坐标提取完成，共找到 {len(self.axis_coords)} 个轴网标记")
        
        # 打印找到的轴网信息
        for axis, coords in sorted(self.axis_coords.items()):
            coord_str = ""
            if 'x' in coords:
                coord_str += f"X={coords['x']:.2f}"
            if 'y' in coords:
                coord_str += f"Y={coords['y']:.2f}"
            logger.debug(f"轴网 {axis}: {coord_str}")
    
    def get_axis_coordinate(self, axis_name: str, coord_type: str) -> Optional[float]:
        """
        获取指定轴网的坐标
        
        Args:
            axis_name: 轴网名称
            coord_type: 坐标类型 ('x' 或 'y')
            
        Returns:
            坐标值或None
        """
        if axis_name in self.axis_coords:
            coords = self.axis_coords[axis_name]
            return coords.get(coord_type)
        return None
    
    def calculate_region_boundaries(self) -> List[Dict]:
        """
        计算每个区域的坐标边界
        
        Returns:
            包含边界信息的区域列表
        """
        logger.info("计算区域坐标边界...")
        
        region_boundaries = []
        
        for region in self.regions:
            # 解析X范围
            x_start, x_end = self.parse_range(region['x_range'])
            # 解析Y范围
            y_start, y_end = self.parse_range(region['y_range'])
            
            # 获取实际坐标
            min_x = self.get_axis_coordinate(x_start, 'x')
            max_x = self.get_axis_coordinate(x_end, 'x')
            min_y = self.get_axis_coordinate(y_start, 'y')
            max_y = self.get_axis_coordinate(y_end, 'y')
            
            # 检查是否所有坐标都找到了
            missing_axes = []
            if min_x is None:
                missing_axes.append(f"X轴网 {x_start}")
            if max_x is None:
                missing_axes.append(f"X轴网 {x_end}")
            if min_y is None:
                missing_axes.append(f"Y轴网 {y_start}")
            if max_y is None:
                missing_axes.append(f"Y轴网 {y_end}")
            
            if missing_axes:
                logger.warning(f"区域 {region['description']} 缺少轴网: {', '.join(missing_axes)}")
                continue
            
            # 确保坐标范围正确（min < max）
            if min_x > max_x:
                min_x, max_x = max_x, min_x
            if min_y > max_y:
                min_y, max_y = max_y, min_y
            
            boundary = {
                'name': region['name'],
                'description': region['description'],
                'min_x': min_x,
                'max_x': max_x,
                'min_y': min_y,
                'max_y': max_y,
                'x_range': region['x_range'],
                'y_range': region['y_range']
            }
            
            region_boundaries.append(boundary)
            logger.debug(f"区域边界: {region['description']} -> "
                        f"X:[{min_x:.2f}, {max_x:.2f}], Y:[{min_y:.2f}, {max_y:.2f}]")
        
        logger.info(f"成功计算 {len(region_boundaries)} 个区域的边界")
        return region_boundaries
    
    def find_columns(self, doc: ezdxf.document.Drawing):
        """
        查找所有柱子
        
        Args:
            doc: ezdxf文档对象
        """
        logger.info("开始查找柱子...")
        
        self.columns = []
        modelspace = doc.modelspace()
        
        # 查找文本实体中的柱子
        for entity in modelspace:
            try:
                if entity.dxftype() in ['TEXT', 'MTEXT']:
                    text_content = entity.dxf.text.upper()
                    if COLUMN_IDENTIFIER in text_content:
                        insert_point = entity.dxf.insert
                        self.columns.append({
                            'x': insert_point.x,
                            'y': insert_point.y,
                            'type': 'TEXT',
                            'content': text_content
                        })
                        logger.debug(f"发现文本柱子: {text_content} at ({insert_point.x:.2f}, {insert_point.y:.2f})")
                
                # 查找块引用中的柱子
                elif entity.dxftype() == 'INSERT':
                    block_name = entity.dxf.name.upper()
                    insert_point = entity.dxf.insert
                    
                    # 检查块名是否包含柱子标识符
                    if COLUMN_IDENTIFIER in block_name:
                        self.columns.append({
                            'x': insert_point.x,
                            'y': insert_point.y,
                            'type': 'BLOCK',
                            'content': block_name
                        })
                        logger.debug(f"发现块柱子: {block_name} at ({insert_point.x:.2f}, {insert_point.y:.2f})")
                    
                    # 检查块的属性
                    if hasattr(entity, 'attribs'):
                        for attrib in entity.attribs:
                            if hasattr(attrib, 'dxf') and hasattr(attrib.dxf, 'text'):
                                attrib_text = attrib.dxf.text.upper()
                                if COLUMN_IDENTIFIER in attrib_text:
                                    self.columns.append({
                                        'x': insert_point.x,
                                        'y': insert_point.y,
                                        'type': 'BLOCK_ATTRIB',
                                        'content': attrib_text
                                    })
                                    logger.debug(f"发现属性柱子: {attrib_text} at ({insert_point.x:.2f}, {insert_point.y:.2f})")
                                    break
                        
            except AttributeError:
                # 某些实体可能没有相关属性
                continue
            except Exception as e:
                logger.warning(f"处理实体时出错: {e}")
                continue
        
        logger.info(f"柱子查找完成，共找到 {len(self.columns)} 个柱子")
    
    def match_columns_to_regions(self, region_boundaries: List[Dict]):
        """
        将柱子匹配到对应区域并计数
        
        Args:
            region_boundaries: 区域边界列表
        """
        logger.info("开始匹配柱子到区域...")
        
        # 初始化计数器
        self.region_counts = {region['description']: 0 for region in region_boundaries}
        unmatched_columns = []
        
        # 遍历所有柱子
        for column in self.columns:
            x, y = column['x'], column['y']
            matched = False
            
            # 检查柱子属于哪个区域
            for region in region_boundaries:
                if (region['min_x'] <= x <= region['max_x'] and 
                    region['min_y'] <= y <= region['max_y']):
                    
                    self.region_counts[region['description']] += 1
                    logger.debug(f"柱子 ({x:.2f}, {y:.2f}) 匹配到区域: {region['description']}")
                    matched = True
                    break
            
            if not matched:
                unmatched_columns.append(column)
                logger.debug(f"柱子 ({x:.2f}, {y:.2f}) 未匹配到任何区域")
        
        logger.info(f"匹配完成，{len(unmatched_columns)} 个柱子未匹配到区域")
        
        if unmatched_columns:
            logger.warning("以下柱子未匹配到任何区域：")
            for col in unmatched_columns[:10]:  # 只显示前10个
                logger.warning(f"  - ({col['x']:.2f}, {col['y']:.2f}) {col['content']}")
            if len(unmatched_columns) > 10:
                logger.warning(f"  ... 还有 {len(unmatched_columns) - 10} 个未匹配的柱子")
    
    def print_results(self):
        """打印统计结果"""
        logger.info("=" * 60)
        logger.info("柱子区域统计结果")
        logger.info("=" * 60)
        
        total_columns = 0
        
        # 按负责人分组显示结果
        person_groups = {}
        for description, count in self.region_counts.items():
            # 从描述中提取负责人名字
            person = description.split(':')[0].strip()
            if person not in person_groups:
                person_groups[person] = []
            person_groups[person].append((description, count))
            total_columns += count
        
        # 显示每个负责人的区域统计
        for person, regions in person_groups.items():
            logger.info(f"\n【{person}】负责区域：")
            person_total = 0
            for description, count in regions:
                logger.info(f"  {description} - 柱子数量: {count}个")
                person_total += count
            logger.info(f"  {person} 负责区域总计: {person_total}个柱子")
        
        logger.info("-" * 60)
        logger.info(f"所有区域柱子总数: {total_columns}个")
        logger.info(f"图纸中发现的柱子总数: {len(self.columns)}个")
        
        if len(self.columns) != total_columns:
            unmatched = len(self.columns) - total_columns
            logger.info(f"未分配到区域的柱子: {unmatched}个")
        
        logger.info("=" * 60)
    
    def analyze_dwg(self, file_path: str) -> bool:
        """
        分析DWG文件的主要方法
        
        Args:
            file_path: DWG文件路径
            
        Returns:
            是否分析成功
        """
        try:
            # 1. 加载DWG文件
            doc = self.load_dwg_file(file_path)
            if doc is None:
                return False
            
            # 2. 提取轴网坐标
            self.extract_axis_coordinates(doc)
            if not self.axis_coords:
                logger.error("未找到任何轴网坐标，无法进行区域划分")
                return False
            
            # 3. 计算区域边界
            region_boundaries = self.calculate_region_boundaries()
            if not region_boundaries:
                logger.error("无法计算区域边界，请检查轴网定义")
                return False
            
            # 4. 查找柱子
            self.find_columns(doc)
            if not self.columns:
                logger.warning("未找到任何柱子")
            
            # 5. 匹配柱子到区域并计数
            self.match_columns_to_regions(region_boundaries)
            
            # 6. 输出结果
            self.print_results()
            
            return True
            
        except Exception as e:
            logger.error(f"分析过程中发生错误: {e}")
            return False


def main():
    """主函数"""
    print("建筑DWG/DXF文件柱子区域统计分析脚本")
    print("=" * 50)
    
    # 文件路径（从配置文件读取）
    original_file_path = DXF_FILE_PATH
    dxf_file_path = original_file_path
    
    # 检查文件是否存在
    if not os.path.exists(original_file_path):
        print(f"错误：文件不存在 - {original_file_path}")
        print("请检查文件路径是否正确")
        input("按任意键退出...")
        return
    
    # 检查文件格式并处理转换
    file_ext = os.path.splitext(original_file_path)[1].lower()
    if file_ext == '.dwg':
        # 检查是否已有对应的DXF文件
        dxf_file_path = original_file_path.replace('.dwg', '.dxf')
        
        if os.path.exists(dxf_file_path):
            print("✅ 发现已转换的DXF文件，将直接使用")
            print(f"   DXF文件: {dxf_file_path}")
            print("=" * 50)
        else:
            print("🔍 检测到DWG文件，需要转换为DXF格式才能分析")
            print("=" * 50)
            print("📋 转换选项：")
            print("1. 运行 'python DWG转换工具.py' 获取详细转换指导")
            print("2. 手动在AutoCAD中将DWG另存为DXF格式")
            print("3. 使用在线转换服务")
            print()
            print("转换步骤（AutoCAD）：")
            print("   文件 → 打开 → 选择DWG文件")
            print("   文件 → 另存为 → 选择DXF格式")
            print(f"   保存为: {dxf_file_path}")
            print()
            
            choice = input("转换完成后按回车继续，或输入 'q' 退出: ").strip().lower()
            if choice == 'q':
                print("👋 退出程序")
                return
            
            # 再次检查DXF文件是否存在
            if not os.path.exists(dxf_file_path):
                print(f"❌ 未找到转换后的DXF文件: {dxf_file_path}")
                print("请先完成DWG到DXF的转换，然后重新运行脚本")
                input("按任意键退出...")
                return
            
            print(f"✅ 找到DXF文件，开始分析: {dxf_file_path}")
            print("=" * 50)
    
    elif file_ext == '.dxf':
        dxf_file_path = original_file_path
        print("✅ 检测到DXF文件，可以直接分析")
        print("=" * 50)
    
    else:
        print(f"❌ 不支持的文件格式: {file_ext}")
        print("支持的格式: .dwg, .dxf")
        input("按任意键退出...")
        return
    
    try:
        # 创建分析器实例
        analyzer = ColumnAnalyzer()
        
        # 执行分析
        success = analyzer.analyze_dwg(dxf_file_path)
        
        if success:
            print("\n分析完成！详细日志请查看 '柱子统计分析.log' 文件")
        else:
            print("\n分析失败！请查看错误日志")
            
    except KeyboardInterrupt:
        print("\n用户中断了程序执行")
    except Exception as e:
        print(f"\n程序执行过程中发生未知错误: {e}")
        logger.error(f"主函数执行错误: {e}")
    
    input("\n按任意键退出...")


if __name__ == "__main__":
    main()
