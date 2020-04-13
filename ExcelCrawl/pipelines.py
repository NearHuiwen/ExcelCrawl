# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


class ExcelcrawlPipeline(object):
    def process_item(self, item, spider):
        spider.write_dict.setdefault(item['row_index'], item['processing_state'])
        return item
