class url_manager():
    # 初始化方法，创建两个集合用于存储新的和已处理的 URL
    def __init__(self):
        self.new_urls = set()  # 存储待爬取的 URL
        self.old_urls = set()  # 存储已爬取的 URL

    # 添加一个新的 URL
    def add_new_url(self, url):
        if url is None or len(url) == 0:  # 如果 URL 为空，直接返回
            return
        if url in self.new_urls or url in self.old_urls:  # 如果 URL 已经在 new_urls 或 old_urls 中，直接返回
            return
        self.new_urls.add(url)  # 将 URL 添加到 new_urls 集合中

    # 批量添加新的 URL
    def add_new_urls(self, urls):
        if urls is None or len(urls) == 0:  # 如果 URL 列表为空，直接返回
            return
        for url in urls:  # 遍历 URL 列表
            self.add_new_url(url)  # 调用 add_new_url 方法将每个 URL 添加到 new_urls 集合中

    # 获取一个新的 URL
    def get_url(self):
        if self.has_new_url():  # 如果存在新的 URL
            url = self.new_urls.pop()  # 从 new_urls 集合中移除一个 URL
            self.old_urls.add(url)  # 将这个 URL 添加到 old_urls 集合中
            return url  # 返回这个 URL
        else:
            return None  # 如果没有新的 URL，返回 None

    # 检查是否存在新的 URL
    def has_new_url(self):
        return len(self.new_urls) > 0  # 如果 new_urls 集合中有 URL，返回 True，否则返回 False


# 测试 UrlManager 类
if __name__ == "__main__":
    url_manager = url_manager()

    url_manager.add_new_url("url1")  # 添加单个 URL
    url_manager.add_new_urls(["url1", "url2"])  # 批量添加 URL
    print(url_manager.new_urls, url_manager.old_urls)  # 打印 new_urls 和 old_urls 集合的内容

    print("#" * 30)
    new_url = url_manager.get_url()  # 获取一个新的 URL
    print(url_manager.new_urls, url_manager.old_urls)  # 打印 new_urls 和 old_urls 集合的内容

    print("#" * 30)
    new_url = url_manager.get_url()  # 再次获取一个新的 URL
    print(url_manager.new_urls, url_manager.old_urls)  # 打印 new_urls 和 old_urls 集合的内容

    print("#" * 30)
    print(url_manager.has_new_url())  # 检查是否存在新的 URL
