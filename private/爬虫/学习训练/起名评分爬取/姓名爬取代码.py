from urllib.parse import urlencode
import requests

url = "https://life.httpcn.com/xingming.asp"
data = {
    "isbz": 1,
    "xing": "裴".encode("gb2312"),
    "ming": "大旺".encode("gb2312"),
    "sex": 1,
    "data_type": 0,
    "year": 1980,
    "month": 7,
    "day": 22,
    "hour": 11,
    "minute": 10,
    "pid": "北京".encode("gb2312"),
    "cid": "北京".encode("gb2312"),
    "wxxy": 0,
    "xishen": "金".encode("gb2312"),
    "yongshen": "金".encode("gb2312"),
    "check_agree": "agree",
    "act": "submit"
}

headers = {
    "Content-Type": "application/x-www-form-urlencoded"
}

r = requests.get(url, data=urlencode(data), headers=headers)

print(r.status_code)
# r.encoding = "gb2312"
print(r.text)

