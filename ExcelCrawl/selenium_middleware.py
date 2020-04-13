# -*- coding: utf-8 -*-
import os
import time

from selenium import webdriver

from scrapy.http import HtmlResponse


class SeleniumMiddleware(object):

    def __del__(self):
        self.close_driver()

    # 关闭浏览器
    def close_driver(self):
        if (self.driver is not None):
            self.driver.quit()
            self.driver = None

    def process_request(self, request, spider):
        if spider.name == "news_spider":
            chrome_options = webdriver.ChromeOptions()
            # 启用headless模式
            chrome_options.add_argument('--headless')
            # 关闭gpu
            chrome_options.add_argument('--disable-gpu')
            # 关闭图像显示
            chrome_options.add_argument('--blink-settings=imagesEnabled=false')

            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'chromedriver.exe')

            self.driver = webdriver.Chrome(executable_path=path, chrome_options=chrome_options)
            self.driver.get(request.url)
            download_time = request.meta.get('download_time', "")
            # 设置加载时间，加载js渲染的网站，并且防止请求频繁被禁IP
            if (download_time):
                time.sleep(download_time)
            else:
                time.sleep(0.5)
            body = self.driver.page_source
            print("加载完成："+str(request.url))
            self.close_driver()
            return HtmlResponse(url=request.url, body=body, encoding='utf-8', request=request)
        return None
