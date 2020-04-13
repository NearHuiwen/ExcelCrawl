# -*- coding: utf-8 -*-

import random

import requests
import scrapy
import xlrd
from PyQt5 import QtWidgets
from fake_useragent import UserAgent
from lxml import etree
from scrapy import Request
import xlwings as xw

from ExcelCrawl.items import ExcelcrawlItem
import re

from ExcelCrawl.utils.common import str_to_bool


class NewsSpider(scrapy.Spider):
    name = 'news_spider'
    write_dict = {}

    def __init__(self, filename, sheet_index, urlcol, savecol, hastitile, has_proxy, download_time, *args, **kwargs):
        super(NewsSpider, self).__init__(*args, **kwargs)

        self.filename = filename
        self.sheet_index = int(sheet_index)
        self.urlcol = int(urlcol)
        self.savecol = int(savecol)
        self.hastitile = str_to_bool(hastitile)
        self.has_proxy = str_to_bool(has_proxy)
        self.download_time = float(download_time)

        worksheet = xlrd.open_workbook(self.filename)  # 打开Excel文件
        sheet = worksheet.sheet_by_index(self.sheet_index)  # 打开Excel文件第几个表

        self.urls_list = sheet.col_values(self.urlcol)

        if (self.hastitile):
            del (self.urls_list[0])

    def start_requests(self):
        if (self.has_proxy):
            url = "https://www.xicidaili.com/nt"
            ua = UserAgent()
            header = {'User-Agent': ua.random}
            response = requests.get(url=url, headers=header)
            etree_obj = etree.HTML(response.text)
            tr_list = etree_obj.xpath('//table[@id="ip_list"]/tr')
            proxy_list = []
            for tr_ip in tr_list[1:51]:
                pattern = r"\d+\.?\d*"
                # 速度
                speed = tr_ip.xpath('./td[7]/div/div/@style')[0].strip()
                speed = float(re.findall(pattern, speed)[0])
                if (95 <= speed):
                    # 连接时间
                    connection_time = tr_ip.xpath('./td[8]/div/div/@style')[0].strip()
                    connection_time = float(re.findall(pattern, connection_time)[0])
                    if (95 <= connection_time):
                        # 存活时间
                        survival_time = tr_ip.xpath('./td[9]/text()')[0].strip()
                        if ("天" in survival_time):
                            ip = tr_ip.xpath('./td[2]/text()')[0].strip()
                            scheme = tr_ip.xpath('./td[6]/text()')[0].strip().lower()
                            port = tr_ip.xpath('./td[3]/text()')[0].strip()
                            proxy = '%s://%s:%s' % (scheme, ip, port)
                            proxy_list.append(proxy)

        # 判断是否网站
        pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
        for row_index, crawl_url in enumerate(self.urls_list):
            crawl_url = str(crawl_url)
            if (self.hastitile):
                row_index = row_index + 1
            if ("" == crawl_url):
                self.write_dict.setdefault(row_index, "网站为空")
            elif (re.match(pattern, crawl_url)):
                if (self.has_proxy):
                    proxy_ip = random.choice(proxy_list)
                    meta = {'proxy_ip': proxy_ip, 'row_index': row_index, 'download_time': self.download_time}
                else:
                    meta = {'row_index': row_index, 'download_time': self.download_time}
                yield Request(url=crawl_url, meta=meta, callback=self.processing_parse, dont_filter=True)
            else:
                self.write_dict.setdefault(row_index, "该字段不是网站")

    def processing_parse(self, response):
        row_index = response.meta.get('row_index', "0")
        if (200 == response.status):
            try:
                #在这里编写需爬取的东西，这里举个例子，爬取新闻标题
                # 搜狐新闻
                if ("www.sohu.com" in response.request.url):
                        processing_state = response.xpath('//div[@class="text-title"]/h1/text()').extract_first().strip()
                # 腾讯新闻
                elif ("new.qq.com" in response.request.url):
                        processing_state = response.xpath("//div[@class='LEFT']/h1/text()").extract_first().strip()
                else:
                    processing_state = "该网站为非可爬取网站"
            except Exception as e:
                exception_str = "错误点:" + "response.xpath" + '\n' + '错误的原因是:' + str(
                    e)
                print(exception_str)
                processing_state = "该网站爬取失败"
        elif (403 == response.status):
            processing_state = "没有权限访问该网站"
        elif (404 == response.status):
            processing_state = "该网站不存在"
        elif (302 == response.status):
            processing_state = "该网站被重定向无法访问"
        else:
            processing_state = "错误代码：" + response.status

        item = ExcelcrawlItem()
        item["row_index"] = row_index
        item["crawl_url"] = response.request.url
        item["processing_state"] = processing_state
        yield item

    def close(spider, reason):
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(spider.filename)  # 打开Excel文件
        sheet = wb.sheets[spider.sheet_index]  # 打开Excel文件第几个表
        try:
            for key, value in spider.write_dict.items():
                print('写入第' + str(key) + '行:' + value)
                sheet.range(key + 1, spider.savecol + 1).value = value
        except Exception as e:
            exception_str = "错误点:" + "sheet.range(key + 1, spider.savecol + 1).value = value" + '\n' + '错误的原因是:' + str(
                e)
            print(exception_str)
            raise e
        finally:
            wb.save()  # 保存文档
            wb.close()  # 关闭文档
            app.quit()  # 停止excel程序
            QtWidgets.QMessageBox.about(None, "提示", "爬取完成")