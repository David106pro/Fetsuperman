import pandas as pd
from imap_tools import MailBox
from bs4 import BeautifulSoup
import ssl # 导入ssl模块

# --- 请在这里修改您的邮箱信息 ---
EMAIL_ADDRESS = "fuchaoyi@enjoy-tv.cn"
# !!! 重要：这里应该填写您在邮箱网页设置中生成的"授权码"，而不是登录密码
PASSWORD = "Yy10060106" 
# 腾讯企业邮的服务器地址，如果不是腾讯，请修改为正确的地址
IMAP_SERVER = "imap.exmail.qq.com" 
# 您存放这些邮件的文件夹/分组名称
MAILBOX_FOLDER = "INBOX" # "INBOX" 是收件箱，请修改为您实际的文件夹名

def extract_info_from_email(email_body):
    """从邮件正文中提取信息"""
    soup = BeautifulSoup(email_body, 'lxml')
    
    data = {}
    copyright_info = ""
    
    # 提取表格数据
    table = soup.find('table')
    if table:
        rows = table.find_all('tr')
        if len(rows) > 1:
            headers = [th.text.strip() for th in rows[0].find_all('td')] # 获取表头
            cells = rows[1].find_all('td')
            for i, header in enumerate(headers):
                if i < len(cells):
                    data[header] = cells[i].text.strip()
    
    # 提取红色字体的版权信息
    # 查找 style 属性中包含 color:red 的标签
    red_text_tag = soup.find(lambda tag: tag.name and 'color:red' in tag.get('style', '').lower())
    if red_text_tag:
        copyright_info = red_text_tag.text.strip()
    else:
        # 备用方案，查找 <font color="red"> 标签
        font_tag = soup.find('font', attrs={'color': 'red'})
        if font_tag:
            copyright_info = font_tag.text.strip()
            
    data['版权信息'] = copyright_info
    return data

# 用于存放所有提取结果的列表
all_data = []

try:
    print("正在尝试连接邮箱服务器...")
    print(f"服务器地址: {IMAP_SERVER}")
    print(f"邮箱账号: {EMAIL_ADDRESS}")
    print("注意: 如果连接失败，请确保:")
    print("1. 已在邮箱网页端开启IMAP服务")
    print("2. 使用的是授权码而非普通密码")
    print("3. 网络防火墙允许IMAP连接")
    
    # 尝试连接，设置超时时间
    mailbox = MailBox(IMAP_SERVER)
    mailbox.login(EMAIL_ADDRESS, PASSWORD, initial_folder=MAILBOX_FOLDER)
    
    print(f"成功登录邮箱 {EMAIL_ADDRESS}，开始处理文件夹 '{MAILBOX_FOLDER}' 中的邮件...")
    
    # 获取邮件数量
    mail_count = len(list(mailbox.fetch(limit=None, mark_seen=False)))
    print(f"在文件夹 '{MAILBOX_FOLDER}' 中找到 {mail_count} 封邮件")
    
    if mail_count == 0:
        print("未找到邮件，请检查文件夹名称是否正确")
        print("常见文件夹名称：INBOX(收件箱)、Sent(已发送)、Draft(草稿箱)")
        mailbox.logout()
        exit()
    
    # 遍历文件夹中的每一封邮件
    for i, msg in enumerate(mailbox.fetch(), 1):
        print(f"正在处理第 {i}/{mail_count} 封邮件: {msg.subject}")
        if msg.html:
            extracted_data = extract_info_from_email(msg.html)
            if extracted_data:
                all_data.append(extracted_data)
        else:
            print(f"邮件 {msg.subject} 不包含HTML内容，已跳过。")
    
    mailbox.logout()

    print(f"\n邮件处理完成，共提取到 {len(all_data)} 条数据。")

    # 将数据保存到Excel文件
    if all_data:
        df = pd.DataFrame(all_data)
        output_filename = '影片信息汇总.xlsx'
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"数据已成功保存到文件: {output_filename}")
    else:
        print("没有提取到任何有效数据。")

except Exception as e:
    print(f"\n处理过程中发生错误: {e}")
    print("请检查以下几点：")
    print("1. IMAP服务器地址是否正确？（例如腾讯企业邮是 imap.exmail.qq.com）")
    print("2. 密码是否为在邮箱网页端生成的【授权码】？")
    print("3. 邮箱是否已开启IMAP服务？")
    print("4. 邮件文件夹名称是否正确？")