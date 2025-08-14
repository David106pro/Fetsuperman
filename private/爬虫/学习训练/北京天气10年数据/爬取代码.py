import requests
import pandas as pd
from io import StringIO  # 导入StringIO

# 设置目标URL
url = "http://tianqi.2345.com/Pc/GetHistory"
# 请求头设置
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
    "Referer": "http://tianqi.2345.com/wea_history/54511.htm",
    "Cookie": "lastCountyId=54511; lastCountyTime=1720839003; lastCountyPinyin=beijing; lastProvinceId=12; lastCityId=54511; Hm_lvt_a3f2879f6b3620a363bec646b7a8bcdd=1720839007; Hm_lpvt_a3f2879f6b3620a363bec646b7a8bcdd=1720839007; HMACCOUNT=4FA0EBF46E976E3E"
}
def craw_table(year, month):
    # 参数设置
    params = {
        "areaInfo[areaId]": 54511,  # 地区ID
        "areaInfo[areaType]": 2,  # 地区类型
        "date[year]": year,
        "date[month]": month
    }
    # 发送GET请求
    resp = requests.get(url, headers=headers, params=params)
    # 打印响应状态码
    # print(resp.status_code)
    # 检查响应状态码是否为200
    if resp.status_code == 200:
        # 解析响应中的JSON数据，并提取出"data"字段
        data = resp.json()["data"]
        # 将HTML字符串包裹在StringIO对象中
        data_io = StringIO(data)
        # 使用pandas库读取HTML数据，并将其转换为DataFrame对象
        df = pd.read_html(data_io)[0]
        # 打印0ataFrame的前5行，检查数据是否正确
        return df
    else:
        print("Failed to retrieve data")

df_list = []
for year in range(2013, 2024):
    for month in range(1, 13):
        print("爬取：",year, month)
        df = craw_table(year, month)
        df_list.append(df)

pd.concat(df_list).to_excel("北京10年天气数据.xlsx", index=False)






