# 博客园网页爬取（后台数据异步加载：用Jason）
# 1、请求访问，筛选有效headers
# 2、请求参数：字典key，字符串key：启动器中对应链接的jason中的大括号包含url和data
# 3、用bs4解析，封装函数，整页循环
# 4、pandas存储到表格中

import json
import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://www.cnblogs.com/AggSite/AggSitePostList"

headers = {
    "Content-Type": "application/json; charset=UTF-8",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
}

def craw_page(page_index):
    data = {
        "CategoryType": "SiteHome",
        "ParentCategoryId": 0,
        "CategoryId": 808,
        "PageIndex": page_index,
        "TotalPostCount": 2000,
        "ItemListActionName": "AggSitePostList"
    }
    # 将数据转换为 JSON 字符串
    data_json = json.dumps(data)
    resp = requests.post(url, data=data_json, headers=headers)
    # print(resp.status_code)
    return resp.text

def parse_data(html):
    soup = BeautifulSoup(html, "html.parser")
    articles = soup.find_all("article", class_="post-item")
    datas = []
    for article in articles:
        link = article.find("a", class_="post-item-title")
        title = link.get_text().strip()  # 修改：获取文本并去除前后空格
        href = link["href"]

        author = article.find("a", class_="post-item-author").get_text().strip()  # 修改：获取作者文本并去除前后空格

        icon_digg = 0
        icon_comment = 0
        icon_views = 0
        for a in article.find_all("a"):
            if "icon_digg" in str(a):
                icon_digg = a.find("span").get_text()
            if "icon_comment" in str(a):
                icon_comment = a.find("span").get_text()
            if "icon_views" in str(a):
                icon_views = a.find("span").get_text()

        datas.append([title, href, author, icon_digg, icon_comment, icon_views])
    return datas

all_datas=[]
for page in range(200):
    print("正在爬取：",page)
    html = craw_page(0)
    datas = parse_data(html)
    all_datas.extend(datas)

df = pd.DataFrame(all_datas, columns=["title", "href", "author", "icon_digg", "icon_comment", "icon_views"])
df.to_excel("博客园200页文章信息.xlsx", index = False)