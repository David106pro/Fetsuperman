import logging  # 导入日志模块
import time  # 导入时间模块
import urllib.parse  # 导入 URL 解析模块
import requests  # 导入请求模块
import random  # 导入随机模块
from bs4 import BeautifulSoup  # 从 BeautifulSoup 库导入 BeautifulSoup 类
from playwright.sync_api import sync_playwright, Playwright  # 从 playwright 库导入相关模块
from selenium.webdriver import Chrome  # 从 selenium 库导入 Chrome 类
from utils_log import init_log, init_log_name  # 从自定义模块导入函数

init_log(init_log_name(f'D:/logs', __file__))  # 初始化日志
path_to_extension = "./my-extension"  # 定义扩展路径
user_data_dir = "/tmp/test-user-data-dir"  # 定义用户数据目录

headers = {  # 定义请求头
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0",
    "cookie": "__mta=143505930.1721729537864.1722391669891.1722398395719.56; uuid_n_v=v1; uuid=09EDF75048DC11EFAD2C9D84A9502795C74339EF0D0246B69AA625DDDBABD5AB; _lxsdk_cuid=190906c187fc8-02326d3fde828b-4c657b58-151800-190906c187fc8; WEBDFPID=1uxz77v921v05uu7yyy3xw2852014z7x809v6x7954z97958v86v0749-2037168061264-1721808060272KMUCGQQfd79fef3d01d5e9aadc18ccd4d0c95071445; _lx_utm=utm_source%3DBaidu%26utm_medium%3Dorganic; _lxsdk=09EDF75048DC11EFAD2C9D84A9502795C74339EF0D0246B69AA625DDDBABD5AB; _csrf=e621735b58dac37e972a4bb926f6a547551dab7f73b922a119caf2bd348f5558; Hm_lvt_703e94591e87be68cc8da0da7cbd0be2=1722330825,1722332208,1722391670,1722398388; HMACCOUNT=020341BA99F5C478; __mta=143505930.1721729537864.1722332208176.1722398387968.60; Hm_lpvt_703e94591e87be68cc8da0da7cbd0be2=1722398396; _lxsdk_s=19106ed243d-c3c-34f-1af%7C%7C4",
    "referer": "https://www.maoyan.com/query?kw=%E5%9B%9E%E4%B9%A1%E9%80%97%E5%84%BF"
}
devices = ['iPhone X', 'iPhone 11', 'iPhone 12', 'iPhone 13', 'iPhone 14']  # 定义设备列表

# 模拟人工访问界面
def run(name, playwright: Playwright):  # 定义一个名为 run 的函数，接收 Playwright 对象作为参数
    browser = playwright.chromium.launch(channel= "msedge", headless=False,  # 启动 Edge 浏览器实例，有界面，启用沙箱
                                         chromium_sandbox=True)
    device = playwright.devices[random.choice(devices)]  # 从 devices 中随机选择一个设备配置
    context = browser.new_context(  # 根据所选设备配置创建新的浏览器上下文
        **device
    )
    page = context.new_page()  # 在上下文里创建新页面

    url = "https://i.maoyan.com/apollo/search?searchtype=movie&$from=canary"  # 定义要访问的网址
    # 使用 setExtraHTTPHeaders 方法设置请求头
    page.set_extra_http_headers(headers=headers)  # 为当前页面设置额外的 HTTP 请求头
    page.goto(url)  # 让页面跳转到指定网址
    time.sleep(0.5)
    # 模拟输入操作
    page.type('//*[@id="app"]/div/div[4]/div/div[1]/div[1]/input', name)
    # 模拟点击操作
    page.click('//*[@id="app"]/div/div[4]/div/div[2]/div/div/div/div[1]/div/div[1]')
    time.sleep(15)  # 随机等待 3 到 8 秒

    html = page.content()  # 获取当前页面的 HTML 内容
    # print(html)  # 打印 HTML 内容
    soup = BeautifulSoup(html, 'html.parser')
    save_file(f'D:/logs/{name}.html', soup)  # 使用 BeautifulSoup 解析 HTML 内容
    # context.close()
    return soup  # 返回解析后的页面内容

# 爬取票房
def parse_box(soup):
    data_box = soup.find('div', class_='data-box')  # 查找具有特定类名的 div 元素
    if data_box is not None:  # 如果找到
        items = data_box.find_all('div', class_='item')  # 在找到的元素中查找所有具有特定类名的 div 元素
        last_value = items[-1].find('div', class_='value').text  # 获取最后一个元素中的特定类名的 div 元素的文本内容
        return last_value  # 返回文本内容
    else:
        return "no found1"  # 如果未找到，返回特定字符串
# 爬取海报 src
def parse_post(soup):
    imgs = soup.find_all("img")  # 查找所有的 <img> 标签
    for img in imgs:  # 遍历找到的所有图片
        src = img["src"]  # 获取图片的 src 属性值
        if "https://p0.pipi.cn/mmdb/" not in src:  # 如果 src 属性值中不包含特定字符串，跳过
            continue
        src = src.replace("w/300/h/414", "w/464/h/644")  # 将字符串替换
        return src  # 返回替换后的 src 值
# 存储海报到本地 D 盘
def save_post(src, picture_name):
    # 将文件地址改为 D 盘，并新建一个文件夹：“猫眼海报”
    with open(f"D:/猫眼海报/{picture_name}.jpg", "wb") as f:  # 以二进制写入模式打开或创建一个文件
        resp_img = requests.get(src)  # 发送 GET 请求获取图片数据
        if resp_img:  # 如果获取成功
            logging.info(f'size: {len(resp_img.content)}')
            f.write(resp_img.content)  # 将获取到的图片数据写入文件
            return "存储完毕"  # 返回存储成功的消息
        else:
            return "存储失败"  # 如果获取失败，返回存储失败的消息

def save_file(content, file_path):
    logging.info(f'file_path: {file_path} size: {len(content)}')  # 记录日志

    with open(file_path, "wb") as f:  # 以二进制写入模式打开或创建一个文件
        f.write(content)  # 将获取到的图片数据写入文件
    return "存储完毕"  # 返回存储成功的消息

# 爬取评分+评论人数
def get_movie_info(soup):
    left_div = soup.find('div', class_='left-section')  # 查找具有特定类名的 div 元素
    spans = left_div.find_all('span')  # 在找到的元素中查找所有的 span 元素
    span_infos = [span.get_text() for span in spans]  # 获取所有 span 元素的文本内容
    if '人想看' in ''.join(span_infos):  # 根据文本内容判断
        score = 'no found'  # 设置特定值
        rating_count = 'no found'  # 设置特定值
    elif '人评' in ''.join(span_infos):  # 根据文本内容判断
        values = [info.split()[0] for info in span_infos if info.split()]  # 处理文本内容
        score = values[0] if len(values) > 0 else 'no found'  # 设置特定值
        rating_count = values[1] if len(values) > 1 else 'no found'  # 设置特定值
    else:
        values = [info.split()[0] for info in span_infos if info.split()]  # 处理文本内容
        score = values[0] if len(values) > 0 else 'no found'  # 设置特定值
        rating_count = 'no found'  # 设置特定值
    return score, rating_count  # 返回处理后的结果

with sync_playwright() as playwright:  # 同步使用 playwright
    picture_name = name = "我们的未来"
    soup = run(name, playwright)  # 调用函数获取页面内容
    box = parse_box(soup)  # 解析票房
    src = parse_post(soup)  # 解析海报 src
    save_state = save_post(src, picture_name)  # 存储海报
    src_content = ""
    save_state = save_file(f'D:/logs/{name}.jpg', src_content)  # 使用 BeautifulSoup 解析 HTML 内容

    score, rating_count = get_movie_info(soup)  # 获取评分和评论人数
    print("正在爬取——电影名:", name, "票房数", box, "海报存取状态:", save_state, "猫眼评分:", score, "评分人数:", rating_count)  # 打印相关信息