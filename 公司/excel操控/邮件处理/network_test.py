import socket
import ssl
from imap_tools import MailBox

def test_network_connectivity():
    """测试网络连通性"""
    print("=" * 60)
    print("网络连通性测试")
    print("=" * 60)
    
    servers_to_test = [
        ("imap.exmail.qq.com", 993, "腾讯企业邮箱 IMAP SSL"),
        ("imap.exmail.qq.com", 143, "腾讯企业邮箱 IMAP"),
        ("smtp.exmail.qq.com", 587, "腾讯企业邮箱 SMTP"),
        ("www.baidu.com", 80, "百度网站（测试基础网络）")
    ]
    
    for server, port, description in servers_to_test:
        print(f"\n测试 {description}")
        print(f"服务器: {server}:{port}")
        
        try:
            # 创建socket连接测试
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(10)  # 10秒超时
            result = sock.connect_ex((server, port))
            sock.close()
            
            if result == 0:
                print("✓ 网络连通性正常")
                
                # 如果是SSL端口，测试SSL握手
                if port == 993:
                    try:
                        context = ssl.create_default_context()
                        with socket.create_connection((server, port), timeout=10) as sock:
                            with context.wrap_socket(sock, server_hostname=server) as ssock:
                                print("✓ SSL握手成功")
                    except Exception as ssl_e:
                        print(f"✗ SSL握手失败: {ssl_e}")
                        
            else:
                print(f"✗ 网络连接失败，错误代码: {result}")
                
        except Exception as e:
            print(f"✗ 连接测试失败: {e}")

def test_imap_with_different_settings():
    """使用不同设置测试IMAP连接"""
    print("\n" + "=" * 60)
    print("IMAP连接测试（不同配置）")
    print("=" * 60)
    
    EMAIL_ADDRESS = "fuchaoyi@enjoy-tv.cn"
    PASSWORD = "Yy10060106"
    
    # 不同的服务器配置
    configs = [
        {
            "name": "腾讯企业邮箱 - SSL端口993",
            "server": "imap.exmail.qq.com",
            "port": 993,
            "ssl": True
        },
        {
            "name": "腾讯企业邮箱 - 普通端口143",
            "server": "imap.exmail.qq.com", 
            "port": 143,
            "ssl": False
        },
        {
            "name": "尝试QQ邮箱服务器",
            "server": "imap.qq.com",
            "port": 993,
            "ssl": True
        }
    ]
    
    for config in configs:
        print(f"\n测试配置: {config['name']}")
        print(f"服务器: {config['server']}:{config['port']}")
        
        try:
            if config['ssl']:
                mailbox = MailBox(config['server'], port=config['port'])
            else:
                # 对于非SSL连接，需要特殊处理
                mailbox = MailBox(config['server'], port=config['port'])
            
            print("正在尝试登录...")
            mailbox.login(EMAIL_ADDRESS, PASSWORD)
            print("✓ 登录成功！")
            
            # 获取文件夹信息
            folders = mailbox.folder.list()
            print(f"✓ 成功获取到 {len(folders)} 个文件夹")
            
            mailbox.logout()
            print("✓ 测试完成")
            return True
            
        except Exception as e:
            print(f"✗ 失败: {e}")
    
    return False

def main():
    print("邮箱连接诊断工具")
    print("开始全面诊断...")
    
    # 1. 网络连通性测试
    test_network_connectivity()
    
    # 2. IMAP连接测试
    success = test_imap_with_different_settings()
    
    print("\n" + "=" * 60)
    print("诊断总结")
    print("=" * 60)
    
    if success:
        print("✓ 邮箱连接成功！可以继续使用邮件处理脚本")
    else:
        print("✗ 所有连接尝试都失败了")
        print("\n可能的解决方案:")
        print("1. 检查网络连接和防火墙设置")
        print("2. 确认邮箱账号密码正确")
        print("3. 确认已开启IMAP服务")
        print("4. 尝试使用授权码替代普通密码")
        print("5. 联系网络管理员检查企业防火墙设置")

if __name__ == "__main__":
    main()
