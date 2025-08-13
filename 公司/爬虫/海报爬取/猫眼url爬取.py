import re
from bs4 import BeautifulSoup

html = """
<div class="actors">
<a href="/asgard/celebrity/12561">皮埃尔·罗谢夫&nbsp;/&nbsp;</a><a href="/asgard/celebrity/391321">皮奥·马麦&nbsp;/&nbsp;</a><a href="/asgard/celebrity/155404">罗克珊·梅斯基达</a>

"""

def get_actors(soup):
    div = soup.find('div', class_="actors")
    if div:
        actors = [re.sub(r'\s*[/]\s*', ',', a.text.strip()) for a in div.find_all('a')]
        actors = [re.sub(r'\s*&\nbsp;\s*', ',', actor) for actor in actors]
        actors = [re.sub(r'\s*&nbsp;\s*', '', actor) for actor in actors]
        actors = [re.sub(r'\s*,\s*$', '', actor) for actor in actors]
        result = '/'.join(actors)  # 将数组中的所有值用“/”链接并存入 result 变量
        return result
    else:
        return "no found"

soup = BeautifulSoup(html, 'html.parser')
result = get_actors(soup)
print(result)