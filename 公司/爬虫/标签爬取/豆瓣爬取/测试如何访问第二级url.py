import requests
from bs4 import BeautifulSoup

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Cookie': 'bid=_PFUYl8B7qY; ll="108288"; _pk_id.100001.4cf6=0d53d319030ff1a7.1720419526.; _vwo_uuid_v2=D7538082B8EDB72774B12DFB37F81BD1B|294f9e79497f5751092275e93394e8e5; __yadk_uid=L8nVjMFI5LDAlMd02u6LYgUpC0mJ6hCQ; dbcl2="281958561:gfTbwIjn9IQ"; push_noty_num=0; push_doumail_num=0; __utmv=30149280.28195; ck=GEkJ; __utmc=30149280; __utmc=223695111; frodotk_db="6af6b99d54891ac040e42cd7fd9b5040"; __utmz=30149280.1721603612.22.9.utmcsr=search.douban.com|utmccn=(referral)|utmcmd=referral|utmcct=/movie/subject_search; ap_v=0,6.0; __utma=30149280.718252523.1720419527.1721606853.1721612671.24; __utma=223695111.311806802.1720419527.1721606856.1721612676.23; __utmz=223695111.1721612676.23.10.utmcsr=search.douban.com|utmccn=(referral)|utmcmd=referral|utmcct=/movie/subject_search; _pk_ref.100001.4cf6=%5B%22%22%2C%22%22%2C1721615526%2C%22https%3A%2F%2Fsearch.douban.com%2Fmovie%2Fsubject_search%3Fsearch_text%3D1%2F2%E6%AE%B5%E6%83%85%26cat%3D1002%22%5D; _pk_ses.100001.4cf6=1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0'
}

url = 'https://movie.douban.com/subject/2332443/'

# 发送请求并获取页面内容
response = requests.get(url, headers = headers)
print(response.status_code)
html = response.text
# print(html)

# 使用 BeautifulSoup 解析页面
soup = BeautifulSoup(html, 'html.parser')

# 定位并获取指定文本
info_div = soup.find('div', id='info')
genre_span = info_div.find('span', property='v:genre')
genre_text = genre_span.text if genre_span else '未找到类型信息'

# 打印获取到的文本
print(genre_text)