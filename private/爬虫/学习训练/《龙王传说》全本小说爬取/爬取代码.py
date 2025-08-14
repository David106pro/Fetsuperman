import requests
from bs4 import BeautifulSoup

def get_novel_chapters():
    root_url = "http://www.89996.net/longwangchuanshuo/"
    r = requests.get(root_url)
    r.encoding = "gbk"
    soup = BeautifulSoup(r.text, "html.parser")

    data = []
    for dd in soup.find_all("dd"):
        link = dd.find("a")
        if not link:
            continue
        # 拼接完整的 URL
        full_url = "http://www.89996.net" + link['href']
        data.append((full_url, link.get_text()))
    return data

def get_chapter_content(url):
    r = requests.get(url)
    r.encoding = 'gbk'
    soup = BeautifulSoup(r.text, "html.parser")
    return soup.find("div", id='content').get_text()

# 以下是调用函数并执行后续操作的部分
chapters = get_novel_chapters()
total_cnt = len(chapters)
idx = 0
for chapter in chapters:
    idx +=1
    print(idx, total_cnt)
    url, title = chapter
    # 修正文件名的生成方式，使用更规范的字符串格式化
    content = get_chapter_content(url)
    content = content.replace('\xa0', '')  # 替换为空格
    with open(f"{title}.txt", "w", encoding="gbk") as fout:
        fout.write(content)