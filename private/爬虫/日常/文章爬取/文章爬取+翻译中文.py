import requests
from bs4 import BeautifulSoup
from googletrans import Translator

def scrape_paragraphs(output_file):
    url = input("请输入URL地址: ")
    translator = Translator()  # 创建翻译对象

    try:
        # 发送HTTP请求到指定的URL
        response = requests.get(url)
        response.raise_for_status()  # 如果请求失败，抛出HTTPError

        # 解析HTML内容
        soup = BeautifulSoup(response.content, 'html.parser')

        # 查找所有<p>标签
        paragraphs = soup.find_all('p')

        # 打开文件并写入所有<p>标签中的文本
        with open(output_file, 'w', encoding='utf-8') as file:
            for p in paragraphs:
                text = p.get_text()
                # 检查文本是否为中文
                if not any('\u4e00' <= char <= '\u9fff' for char in text):
                    # 翻译文本为中文
                    translated = translator.translate(text, src='en', dest='zh-cn')
                    # 检查翻译结果是否为 None
                    if translated and translated.text:
                        translated_text = translated.text
                    else:
                        translated_text = text  # 如果翻译失败，保留原文
                else:
                    translated_text = text
                file.write(translated_text + '\n')

        print(f"内容已成功存储到 {output_file}")

    except requests.exceptions.RequestException as e:
        print(f"请求过程中出现错误: {e}")
    except Exception as e:
        print(f"处理过程中出现错误: {e}")

# 示例使用
output_file = "output.txt"  # 输出文件的路径
scrape_paragraphs(output_file)
