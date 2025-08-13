import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
    "Referer": "http://tianqi.2345.com/wea_history/54511.htm",
    "Cookie": "lastCountyId=54511; lastCountyTime=1720839003; lastCountyPinyin=beijing; lastProvinceId=12; lastCityId=54511; Hm_lvt_a3f2879f6b3620a363bec646b7a8bcdd=1720839007; Hm_lpvt_a3f2879f6b3620a363bec646b7a8bcdd=1720839007; HMACCOUNT=4FA0EBF46E976E3E"
}

url = f"https://movie.douban.com/subject/35863336/"
response = requests.get(url, headers=headers)

response.encoding = response.apparent_encoding

if response.status_code != 200:
    print(f"Failed to retrieve search results")

soup = BeautifulSoup(response.text, 'html.parser')

soup.find("div", id_="info")
genres = [span.get_text() for span in soup.find_all('span', property='v:genre')]

if genres:
    print("Genres:", " / ".join(genres))
else:
    print("No genres found")


