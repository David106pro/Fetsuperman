# 海报爬取下载器

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
