import requests
from bs4 import BeautifulSoup
import pprint
import pandas as pd #生成excel表格的
import time

page_index = range(0, 250, 25)  # 定义要抓取的页面索引范围，从0开始，每次增加25，直到250

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}


def download_all_htmls():
    htmls = []  # 创建一个空列表，用于存储每个页面的HTML内容
    for idx in page_index:  # 遍历每个页面索引
        url = f"http://movie.douban.com/top250?start={idx}&filter="  # 构建URL
        print("crawling html:", url)  # 打印当前正在抓取的URL
        for attempt in range(3):  # 尝试重试3次
            r = requests.get(url, headers=headers)  # 发送HTTP GET请求，添加headers参数
            if r.status_code == 200:  # 如果响应状态码是200，表示成功
                htmls.append(r.text)  # 将页面的HTML内容添加到列表中
                break  # 成功后跳出重试循环
            else:
                print(f"Attempt {attempt + 1} failed with status code {r.status_code}. Retrying...")  # 打印失败信息
                time.sleep(5)  # 等待5秒后重试
        else:
            raise Exception(f"Failed to fetch the page after 3 attempts. URL: {url}")  # 如果重试3次后仍然失败，抛出异常
    return htmls  # 返回所有页面的HTML内容列表


def parse_single_html(html):
    soup = BeautifulSoup(html, 'html.parser')  # 使用BeautifulSoup解析HTML文档
    article_items = (
        soup.find("div", class_="article")  # 找到包含电影信息的主要div元素
        .find("ol", class_="grid_view")  # 找到包含电影条目的ol元素
        .find_all("div", class_="item")  # 找到所有包含单个电影信息的div元素
    )
    datas = []  # 创建一个空列表，用于存储每个电影的信息
    for article_item in article_items:  # 遍历每个电影条目
        rank = article_item.find("div", class_="pic").find("em").get_text()  # 提取排名
        info = article_item.find("div", class_="info")  # 找到包含电影详细信息的div元素
        title = info.find("div", class_="hd").find("span", class_="title").get_text()  # 提取电影标题
        stars = (
            info.find("div", class_="bd")
            .find("div", class_="star")
            .find_all("span")  # 找到包含评分信息的所有span元素
        )
        rating_star = stars[0]["class"][0]  # 提取评分星级
        rating_num = stars[1].get_text()  # 提取评分数
        comments = stars[3].get_text()  # 提取评论数

        datas.append({  # 将提取的信息添加到列表中
            "rank": rank,
            "title": title,
            "rating_star": rating_star.replace("rating", "").replace("-t", ""),  # 处理星级信息
            "rating_num": rating_num,
            "comments": comments.replace("人评价", "")  # 处理评论数信息
        })
    return datas  # 返回当前页面的所有电影信息


def main():
    htmls = download_all_htmls()  # 下载所有页面的HTML内容

    all_datas = []  # 创建一个空列表，用于存储所有页面的电影信息
    for html in htmls:  # 遍历每个页面的HTML内容
        all_datas.extend(parse_single_html(html))  # 解析页面并将信息添加到列表中

    df = pd.DataFrame(all_datas)  # 将所有数据转换为DataFrame格式
    df.to_excel("豆瓣电影TOP250.xlsx", index=False)  # 将数据保存到Excel文件中，不包含索引
    print("Data saved to 豆瓣电影TOP250.xlsx")  # 打印保存成功的信息


if __name__ == "__main__":
    main()  # 判断是否在主程序中运行，如果是，则调用主函数
