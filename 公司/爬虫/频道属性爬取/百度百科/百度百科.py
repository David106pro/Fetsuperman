import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

# 你提供的名称列表
name_list = {
    "悬崖之上_极",
    "我在中国拍电影",
    "韩国史五千年_生存之路",
    "92班：联赛梦未成 第二季",
    "92班：联赛梦未成 第一季",
    "加拉帕戈斯群岛：世界的边缘_原音版",
    "绿色动物",
    "野生历史",
    "人类：生命的分支",
    "畜牧",
    "数字",
    "不朽的大帝秦王：秦始皇",
    "生存战略",
    "雄鸟",
    "五行_文明的起源",
    "松广寺",
    "哎呀妈呀",
    "哎呦你真美",
    "爱的味道1",
    "爱的味道2",
    "爱的味道3",
    "宝宝学说话：让宝宝爱表达 会沟通",
    "北京评书大会",
    "北京喜剧幽默大赛",
    "冰雪公主：拯救恐龙世界",
    "常用英语口语800句",
    "超级镜头：白垩纪恐龙时代",
    "超级镜头：昆虫的秘密",
    "超级镜头：深海奇观",
    "超级镜头：雨林里的小动物",
    "超级镜头：远古地球霸主排行榜",
    "超级汽车：动物汽车大碰撞",
    "出发吧去露营！",
    "初级商务英语听说",
    "初级商务英语写作",
    "初级商务英语阅读",
    "春妮的周末时光",
    "云享旅拍",
    "从零开始学旅游英语",
    "从零开始学通用英语语法",
    "大道-高速公路行车安全",
    "大棚蔬菜巧用肥",
    "大秦腔",
    "但愿人长久",
    "当他们渐渐老去",
    "档案",
    "道家的智慧",
    "德云社相声动漫版",
    "低盐饮食保健康",
    "翟叔撩口语系列课程",
    "夺宝秦兵",
    "儿童免疫规划",
    "防治不孕不育症",
    "钢铁记忆",
    "高级商务英语听说",
    "高级商务英语写作",
    "高级商务英语阅读",
    "戈峻夜话",
    "共和国档案",
    "逛遍全中国",
    "韩非子与法家智慧",
    "好奇动物：成语故事新说",
    "好奇动物：奇奇怪怪动物园",
    "好奇动物：英语启蒙乐园",
    "好奇恐龙乐园：奥特战士大战恐龙人",
    "好奇恐龙乐园：绿巨人拯救恐龙世界",
    "好奇恐龙乐园：探索缤纷的恐龙世界",
    "合理膳食与健康",
    "红军不怕远征难",
    "红色摇篮",
    "花生玉米间作套种技术",
    "华山论鉴",
    "华豫之门",
    "华豫之门2021",
    "华豫之门2022",
    "华豫之门2023",
    "欢天戏地",
    "加油吧孩子",
    "加油毛孩子",
    "家庭科学实验：游戏时间",
    "家宴1",
    "家宴2",
    "家宴3",
    "家宴4",
    "健康广场快乐你我",
    "健康乡村之高血压的持久战",
    "健康乡村之农村抗生素的滥用与危害",
    "匠心守艺第二季",
    "匠心守艺第三季",
    "匠心守艺第一季",
    "酱油与健康",
    "解放人民的选择",
    "看电影学英语",
    "科学种田巧施肥",
    "老板悄悄话",
    "老家的味道",
    "老家的味道2021",
    "老家的味道2022",
    "老师请回答",
    "梨园春",
    "梨园春2015",
    "梨园春2016",
    "梨园春2017",
    "梨园春2018",
    "梨园春2019",
    "梨园春2020",
    "梨园春2021",
    "梨园春2022",
    "梨园春2023",
    "李焕之与国歌",
    "厉害了我的歌",
    "辽美旅拍",
    "零基础搞定常用英语口语",
    "论语的智慧",
    "妈妈育上娃",
    "美餐",
    "美语截拳道——打造完美美语发音",
    "孟子的教诲",
    "秘食中国",
    "你从井冈山走来",
    "念城味",
    "念念不忘",
    "农村电力设施和电能保护常识",
    "农村垃圾污水简易处理",
    "暖暖的新家",
    "奇妙的大自然：小动物 大知识",
    "奇妙的动物：动物是如何隐身的",
    "奇妙的仿生学",
    "奇妙的身体：揭秘人体的奥秘",
    "儒商大学",
    "黔行旅拍",
    "秦之声",
    "去野吧餐桌",
    "趣味儿童口才诗：练口才 学表达",
    "日常之味",
    "山海筑梦",
    "陕西人物故事",
    "上菜第二季",
    "生于1978",
    "时代先锋",
    "手足口病防治常识",
    "蔬菜王国历险记：给孩子的植物启蒙课",
    "双语小童话：儿童成长必看的经典童话",
    "天天养生",
    "天下收藏",
    "听语音 学发声：看图认物小百科",
    "捅破数学level1",
    "捅破数学level2",
    "捅破数学level3",
    "捅破数学level4",
    "捅破数学level5",
    "捅破数学level6",
    "脱口而出",
    "玩转魔方：二阶魔方CFOP还原法入门(O,P入门学习)",
    "玩转魔方：基础异型魔方教学",
    "玩转魔方：进阶魔方学习(四阶魔方 五阶魔方)",
    "玩转魔方：三阶魔方CFOP还原法入门(C讲解)",
    "玩转魔方：三阶魔方CFOP还原法入门(F2L详解)",
    "玩转魔方：三阶魔方层先法教学",
    "皖美旅拍1",
    "皖美旅拍2",
    "伟大的贡献",
    "为你喝彩",
    "我爱书画",
    "我爱我家",
    "我家有明星",
    "我要动起来",
    "我要上幼儿园啦！ 幼儿入园起点绘本",
    "武林风",
    "武林风2020",
    "武林风2021",
    "武林风2022",
    "武林风2023",
    "武林笼中对",
    "武林笼中对2019",
    "武林笼中对2020",
    "武林笼中对2022",
    "西藏",
    "吸烟有害健康",
    "嘻哈包袱铺2020年至2022年相声合集-片库",
    "喜剧合伙人",
    "先锋",
    "现代农村生活环境治理",
    "乡村记忆",
    "乡村文明你我他",
    "向前一步",
    "消失的亚洲-澳门",
    "消失的亚洲-吉隆坡",
    "消失的亚洲-曼谷",
    "消失的亚洲-首尔",
    "消失的亚洲-台湾",
    "消失的亚洲-香港上",
    "消失的亚洲-香港下",
    "消失的亚洲-新加坡",
    "硝烟战火一百年",
    "小领袖",
    "新概念英语1通关班",
    "新概念英语2通关班",
    "新概念英语3通关班",
    "新视觉",
    "星夜故事",
    "行进中国",
    "行进中国-黄河篇",
    "养生堂",
    "养生堂",
    "医者",
    "英语能力读写基础Level1",
    "英语能力读写基础Level2",
    "英语能力读写基础Level3",
    "英语能力听说基础Level1",
    "英语能力听说基础Level2",
    "英语能力听说基础Level3",
    "影视风云",
    "影响第二季"
}


# 结果存储
results = []

# 设置请求头，模拟浏览器访问
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# 遍历名称列表，逐个搜索
for name in name_list:
    search_url = f"https://www.baidu.com/s?wd={name}"
    response = requests.get(search_url, headers=headers)

    # 设置正确的编码格式
    response.encoding = response.apparent_encoding

    if response.status_code != 200:
        print(f"Failed to retrieve search results for {name}")
        results.append((name, "Failed to retrieve"))
        continue

    soup = BeautifulSoup(response.text, 'html.parser')
    # 获取百度百科的链接
    first_result = soup.find("h3", class_="t")
    if first_result and "百度百科" in first_result.get_text():
        link = first_result.find("a", href=True)
        if link:
            baike_url = link['href']  # 获取百度百科页面的URL
            baike_response = requests.get(baike_url, headers=headers)
            baike_response.encoding = baike_response.apparent_encoding

            if baike_response.status_code == 200:
                baike_soup = BeautifulSoup(baike_response.text, 'html.parser')
                basic_info = baike_soup.find("div", class_="basicInfo_XhoZ7 J-basic-info")
                if basic_info:
                    item_wrapper = basic_info.find_all("div", class_="itemWrapper_QDDui")
                    genres = []
                    for item in item_wrapper:
                        dt = item.find("dt", class_="basicInfoItem_SJQkr itemName_GjVei")
                        if dt and "类型" in dt.get_text():
                            dd = item.find("dd", class_="basicInfoItem_SJQkr itemValue_qYGbJ")
                            if dd:
                                genre = dd.find("span", class_="text_dt0NV")
                                if genre:
                                    genres.append(genre.get_text())
                    genres = " / ".join(genres) if genres else "No category found"
                else:
                    genres = "No category found"
            else:
                print(f"Failed to retrieve details for {name}")
                genres = "Failed to retrieve details"
        else:
            genres = "No result found"
    else:
        genres = "No result found"

    results.append((name, genres))
    print(f"Name: {name}, Category: {genres}")

    # 适当延时以避免触发反爬机制
    time.sleep(1)

# 将结果保存到 DataFrame 并导出为 Excel 文件
df = pd.DataFrame(results, columns=["Name", "Category"])
df.to_excel("search_results_baidu.xlsx", index=False)
print("Results saved to search_results_baidu.xlsx")
