# 1、人工登录脉脉网页版，“检查”职言网页内容规律（“负载/page”跟随滚动页面变化，url地址：部分表头请求地址+荷载信息）
# 2、request申请html失败，添加“user-agent”和“cookie”
# 3、申请依旧失败，继续添加headers：“Referer”，"X-Csrf-Token"（教程到此为止），
# 4、加入很多条都没有效果，最终我讲全部headers都加到代码里，成功爬取html
# 5、用json.cn进行代码解析，在解析好的文本中寻找合适bs4的字段（需要import json）、
# 6、找到准确字段“list”和“text”并用for循环输出查看是否正确，确认一页爬取完毕
# 7、准备封装，在params中确认起始页（第一页可以运行，第0页运行失败）
# 8、封装函数（页面爬取）
# 9、内容存储到新建txt文件，循环爬取10页内容，文件编码选用utf-8
import requests
import json

url = "https://maimai.cn/sdk/web/content/get_list"

headers = {
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5",
    "Cookie": "seid=s1721114432503; guid=GxseBBgeHAQYGRwEGxgaVhsYGhMcEh0aVhoEGgQaBBwYGwVNTm8KHBkEHRkfBUNYS0xLeQoaBBoEGgQcGBsFT0dFWEJpCgNFQUlPbQpPQUNGCgZmZ35iYQIKHBkEHRkfBV5DYUhPfU9GWlprCgMeHFIKER4cREN9ChEaBBobCn5kClldRU5EQ30CChoEHwVLRkZDUEVn; AGL_USER_ID=66149673-e9e3-44b7-bf15-05fa10399527; _buuid=4053684c-4f32-4caf-bcbc-c1286df8870a; browser_fingerprint=5F497CF8; gdxidpyhxdE=zp%2FZDQDeKukkxj3udInBMCcLlxwI83s9hO3f8NBEaJxZWjRQ%2BnfN%2FkPDwj6synIuh%5CfA1sMk0zDXIJpnp5pUnJrtCzgvD1O7nL%2BiX%2BzzCBVuKxhJnDqGmTWbNsQuLQx5mjCq0Q7hriZAZNowapy6dJ6uUyaZmB6BEmK1lIVi%5CzOHKodD%3A1721115357267; u=242111299; u.sig=8CKKYV5e78pV7C3rOeDj5nwxlF0; access_token=1.a903ec5a725ad673156483184fe8cfbd; access_token.sig=psa_FllQHvn3otIAVCCJ4G8-5vs; u=242111299; u.sig=8CKKYV5e78pV7C3rOeDj5nwxlF0; access_token=1.a903ec5a725ad673156483184fe8cfbd; access_token.sig=psa_FllQHvn3otIAVCCJ4G8-5vs; channel=www; channel.sig=tNJvAmArXf-qy3NgrB7afdGlanM; maimai_version=4.0.0; maimai_version.sig=kbniK4IntVXmJq6Vmvk3iHsSv-Y; session=eyJzZWNyZXQiOiI4QkJiWG10MFFMQXZUT00tWE0yaWRfOTIiLCJ1IjoiMjQyMTExMjk5IiwiX2V4cGlyZSI6MTcyMTIwMDk5NDkyMywiX21heEFnZSI6ODY0MDAwMDB9; session.sig=9Rj9bX9P4eE1kda8sxb0tRG9dzg; HWWAFSESID=eb93aaca7b6d85c357; HWWAFSESTIME=1721180852365; csrftoken=PegyrLCM-V72rWa4bGZsbx_KveXeHX3QK_q0",
    "Priority": "u=1, i",
    "Referer": "https://maimai.cn/gossip_list",
    "Sec-Ch-Ua": "\"Not/A)Brand\";v=\"8\", \"Chromium\";v=\"126\", \"Microsoft Edge\";v=\"126\"",
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": "Windows",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
    "X-Csrf-Token": "PegyrLCM-V72rWa4bGZsbx_KveXeHX3QK_q0"
}

def craw_pages(page_number):
    params = {
        'api': 'gossip/v3/square',
        'u': '242111299',
        'page': 4,
        'before_id': 0
    }

    resp = requests.get(url, params=params, headers=headers)
    # print(resp.status_code)
    data = json.loads(resp.text)
    datas = []
    for text in data["list"]:
        try:
            datas.append(text["text"])
        except UnicodeEncodeError:
            pass
    return datas

with open("脉脉结果.txt","w", encoding="utf-8") as f:
    for page in range(1, 10 + 1):
        print("craw page",page)
        datas = craw_pages(page)
        f.write("\n".join(datas) + "\n")