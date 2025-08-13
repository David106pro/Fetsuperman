# 图片爬取步骤：
# 1、找到根url（第一页网址）并获取html
# 2、“检查”分析html，找到bs4目标“img”和“src”
# 3、获取同一网址页面下图片网址的特殊部分
# 4、排除不需要的图片（if语句判断“upload”）
# 5、拼接获得完整图片网址
# 6、新建同目录文件夹，并用os包下载一个页面内的所有图片
# 7、将两段代码进行函数封装（获取每个图片url特殊部分，拼接完成图片url并下载）
# 8、寻找不同页面网址规律，排除特殊页面，for循环存储每个页面网址进入元组
# 9、完善循环体，爬取


import requests  # 导入 requests 库，用于发送 HTTP 请求
from bs4 import BeautifulSoup  # 导入 BeautifulSoup 库，用于解析 HTML 文档
import os  # 导入 os 库，用于文件和目录操作

url = "https://pic.netbian.com/4kmeinv/"  # 定义要爬取的初始网址

# 定义一个名为 craw_html 的函数，用于获取网页的 HTML 内容
def craw_html(url):
    resp = requests.get(url)  # 发送 GET 请求获取指定 URL 的内容
    resp.encoding = 'gb18030'  # 设置响应的编码为 'gb18030'
    print(resp.status_code)  # 打印响应的状态码
    html = resp.text  # 获取响应的文本内容（即 HTML 代码）
    # print(html)  # 注释掉的打印 HTML 代码的部分
    return html  # 返回获取到的 HTML 内容

# 定义一个名为 prase_and_dowmload 的函数，用于解析 HTML 并下载图片
def prase_and_dowmload(html):
    soup = BeautifulSoup(html, "html.parser")  # 使用 BeautifulSoup 解析 HTML 内容
    imgs = soup.find_all("img")  # 查找所有的 <img> 标签
    for img in imgs:  # 遍历找到的所有图片
        src = img["src"]  # 获取图片的 src 属性值
        if "/uploads/" not in src:  # 如果 src 属性值中不包含 "/uploads/"，则跳过
            continue
        src = f"https://pic.netbian.com{src}"  # 为图片链接补充完整的网址前缀
        print(src)  # 打印图片链接
        filename = os.path.basename(src)  # 获取图片链接中的文件名
        with open(f"美女图片/{filename}","wb") as f:  # 以二进制写入模式打开或创建一个文件
            resp_img = requests.get(src)  # 发送 GET 请求获取图片数据
            f.write(resp_img.content)  # 将获取到的图片数据写入文件

urls = ["https://pic.netbian.com/4kmeinv/"] + [  # 定义一个包含多个网址的列表
    f"https://pic.netbian.com/4kmeinv/index_{i}.html"
    for i in range(2, 11)
]
for url in urls:  # 遍历 urls 列表中的每个网址
    print("####正在爬取：", url)  # 打印当前正在爬取的网址
    html = craw_html(url)  # 调用 craw_html 函数获取该网址的 HTML 内容
    prase_and_dowmload(html)  # 调用 prase_and_dowmload 函数解析 HTML 并下载图片