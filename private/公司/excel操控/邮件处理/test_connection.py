import pandas as pd
from imap_tools import MailBox
from bs4 import BeautifulSoup
import ssl

# --- 请在这里修改您的邮箱信息 ---
EMAIL_ADDRESS = "fuchaoyi@enjoy-tv.cn"
PASSWORD = "Yy10060106"  # 如果是普通密码不行，需要生成授权码
IMAP_SERVER = "imap.exmail.qq.com"
MAILBOX_FOLDER = "INBOX"  # 收件箱

def test_connection():
    """测试邮箱连接"""
    try:
        print("正在测试邮箱连接...")
        print(f"服务器: {IMAP_SERVER}")
        print(f"账号: {EMAIL_ADDRESS}")
        print("=" * 50)
        
        # 尝试不同的连接方式
        connection_methods = [
            {"name": "默认SSL连接", "use_ssl": True},
            {"name": "非SSL连接", "use_ssl": False}
        ]
        
        for method in connection_methods:
            try:
                print(f"\n尝试方式: {method['name']}")
                
                if method['use_ssl']:
                    mailbox = MailBox(IMAP_SERVER)
                else:
                    # 尝试非SSL连接（端口143）
                    mailbox = MailBox(IMAP_SERVER, port=143)
                
                print("正在登录...")
                mailbox.login(EMAIL_ADDRESS, PASSWORD)
                
                print("✓ 登录成功！")
                
                # 获取文件夹列表
                print("获取邮箱文件夹列表...")
                folders = mailbox.folder.list()
                print("可用文件夹:")
                for folder in folders:
                    print(f"  - {folder}")
                
                # 选择收件箱并获取邮件数量
                mailbox.folder.set(MAILBOX_FOLDER)
                mail_count = len(list(mailbox.fetch(limit=10, mark_seen=False)))
                print(f"在 '{MAILBOX_FOLDER}' 中找到 {mail_count} 封邮件（仅显示前10封）")
                
                # 获取最新的一封邮件作为示例
                if mail_count > 0:
                    print("\n最新邮件信息:")
                    for i, msg in enumerate(mailbox.fetch(limit=1)):
                        print(f"主题: {msg.subject}")
                        print(f"发件人: {msg.from_}")
                        print(f"日期: {msg.date}")
                        print(f"是否有HTML内容: {'是' if msg.html else '否'}")
                        if msg.html:
                            print(f"HTML内容长度: {len(msg.html)} 字符")
                
                mailbox.logout()
                print("✓ 连接测试成功！")
                return True
                
            except Exception as e:
                print(f"✗ {method['name']} 失败: {e}")
                continue
        
        print("\n所有连接方式都失败了")
        return False
        
    except Exception as e:
        print(f"连接测试失败: {e}")
        return False

if __name__ == "__main__":
    success = test_connection()
    
    if not success:
        print("\n" + "=" * 50)
        print("连接失败可能的原因:")
        print("1. 密码错误 - 需要使用邮箱授权码而不是登录密码")
        print("2. IMAP服务未开启 - 需要在邮箱设置中开启IMAP服务")
        print("3. 网络防火墙阻止 - 检查防火墙设置")
        print("4. 服务器地址错误 - 确认IMAP服务器地址")
        print("\n腾讯企业邮箱设置步骤:")
        print("1. 登录网页版邮箱")
        print("2. 进入设置 -> 客户端设置")
        print("3. 开启IMAP服务")
        print("4. 生成授权码并使用授权码作为密码")