from utils import url_manager  # 导入自定义的 URL 管理器模块
import requests  # 导入 requests 库，用于发送 HTTP 请求
from bs4 import BeautifulSoup  # 导入 BeautifulSoup 库，用于解析 HTML
import re  # 导入正则表达式模块，用于匹配 URL 模式

root_url = "http://www.crazyant.net"  # 定义爬虫的根 URL
urls = url_manager.url_manager()  # 创建 URL 管理器实例
urls.add_new_url(root_url)  # 将根 URL 添加到 URL 管理器中

fout = open("craw_all_pages.txt", "w")  # 打开一个文件，用于保存抓取到的 URL 和标题
while urls.has_new_url():  # 当 URL 管理器中还有待爬取的 URL 时，继续循环
    curr_url = urls.get_url()  # 获取一个新的 URL
    r = requests.get(curr_url, timeout=3)  # 发送 HTTP GET 请求，超时时间为 3 秒
    if r.status_code != 200:  # 检查 HTTP 响应状态码是否为 200
        print("error,return status_code is not 200", curr_url)  # 如果不是 200，打印错误信息
        continue  # 跳过当前循环，继续处理下一个 URL
    soup = BeautifulSoup(r.text, "html.parser")  # 使用 BeautifulSoup 解析 HTML 内容
    title = soup.title.string  # 提取页面的标题

    fout.write("%s\t%s\n" % (curr_url, title))  # 将 URL 和标题写入文件
    fout.flush()  # 刷新文件缓冲区
    print("success:%s, %s,%d" % (curr_url, title, len(urls.new_urls)))  # 打印成功信息

    links = soup.find_all("a")  # 查找页面中的所有链接
    for link in links:  # 遍历所有链接
        href = link.get("href")  # 获取链接的 href 属性
        if href is None:  # 如果 href 为空，跳过
            continue
        pattern = r'^http://www.crazyant.net/\d+.html$'  # 定义 URL 模式
        if re.match(pattern, href):  # 检查链接是否匹配模式
            urls.add_new_url(href)  # 将匹配的链接添加到 URL 管理器中

fout.close()  # 关闭文件
