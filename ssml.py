# This Python file uses the following encoding: utf-8
# 作者：black_lang
# 创建时间：2023/9/5 19:15
# 文件名：test.py

import pprint
import re
import time
import openpyxl
import requests
from faker import Faker
from lxml import etree


class SSML():
    def __init__(self):
        self.has_max_pagenp = False
        self.max_pageno = 2
        self.data = None
        # 信息存储表
        self.universitys = []

    def config(self, data):
        '''
        自己设置config参数，即post提交参数
        :param data:
        :return:
        '''

        self.data = data
        self.data['pageno'] = 1  # 页面数

    def headers(self):
        fa = Faker()
        headers = {
            'user-agent': fa.user_agent()
        }
        return headers

    def school_parse(self, html):
        '''
        解析网页获取院校信息列表
        '''
        tree = etree.HTML(html)
        trs = tree.xpath('//div[@class="zsml-list-box"]/table[@class="ch-table"]/tbody/tr')
        university_li = []
        for tr in trs:
            university = {}
            university['name'] = tr.xpath('./td[1]/form/a/text()')[0]
            university['href'] = 'https://yz.chsi.com.cn' + tr.xpath('./td[1]/form/a/@href')[0]
            university['area'] = tr.xpath('./td[2]/text()')[0]
            university['yjsy'] = 1 if tr.xpath('./td[3]/i/text()') else 0
            university['zzhx'] = 1 if tr.xpath('./td[4]/i/text()') else 0
            university['bsd'] = 1 if tr.xpath('./td[5]/i/text()') else 0
            university_li.append(university)

        if self.has_max_pagenp:
            self.max_pageno = int(tree.xpath('//ul[@class="ch-page"]/li')[-2].xpath('./a/text()')[0])

        return university_li

    def get_school_li(self):
        url_school = 'https://yz.chsi.com.cn/zsml/queryAction.do'
        university_li = []
        while True:
            if self.data['pageno'] < self.max_pageno:
                self.data['pageno'] += 1
            else:
                break
            response = requests.post(url=url_school, headers=self.headers(), data=self.data)
            if response.status_code == 200:
                response.encoding = 'utf-8'
                university_li.extend(self.school_parse(response.text))
            else:
                print('获取异常：', response.headers)
                input('任意输入继续：')
                break
        pprint.pp(university_li)
        print("学校列表信息获取成功！")
        time.sleep(0.5)
        self.universitys = university_li.copy()
        return university_li

    def get_zhuanye(self):
        # 根据院校主页url获取其专业信息，并加入到self.universitys中
        for i in range(len(self.universitys)):
            url = self.universitys[i]['href']
            print(self.universitys[i]['name'] + '\n' + url + '\n')
            response = requests.get(url, headers=self.headers())
            if response.status_code == 200:
                response.encoding = 'utf-8'
                zhuanye_li = self.zhuanye_parse(response.text)
                self.universitys[i]['zhuangye'] = zhuanye_li
            else:
                print('获取学校专业目录出错：', self.universitys[i])
                input('随意输入继续：')

        return self.universitys

    def zhuanye_parse(self, html):
        '''
        解析院校主页
        :param html:
        :return:
        '''
        tree = etree.HTML(html)
        zhuanye_li = []
        trs = tree.xpath('//div[@class="zsml-list-box"]/table/tbody/tr')
        for tr in trs:
            zhuanye = {}
            zhuanye['kaoshi_type'] = tr.xpath('./td[1]/text()')[0]  # 考试方式
            zhuanye['yxs'] = tr.xpath('./td[2]/text()')[0]  # 院系所
            zhuanye['zy'] = tr.xpath('./td[3]/text()')[0]  # 专业
            zhuanye['yyfx'] = tr.xpath('./td[4]/text()')[0]  # 研究方向
            zhuanye['xxfs'] = tr.xpath('./td[5]/text()')[0]  # 学习方式
            zhuanye['zdls'] = ''.join(tr.xpath('./td[6]/div/span/text()'))  # 指导导师
            zhuanye['nzrs'] = tr.xpath('./td[7]/script/text()')[0].split("'")[-2]  # 拟招人数
            zhuanye['ksfw'] = self._get_fw('https://yz.chsi.com.cn' + tr.xpath('./td[8]/a/@href')[0])  # 考试范围链接

            zhuanye['bz'] = tr.xpath('./td[9]/script/text()')[0].split("'")[-2]  # 备注
            zhuanye_li.append(zhuanye)
        return zhuanye_li

    def _get_fw(self, url):
        # 获取考试范围
        response = requests.get(url=url, headers=self.headers())
        response.encoding = 'utf-8'
        tree = etree.HTML(response.text)
        tbodys = tree.xpath('//div[@class="zsml-wrapper"]/div[@class="zsml-result"]/table/tbody')
        fws = ''
        for tbody in tbodys:
            fw_li = [re.sub("[\n \r]", '', fw) for fw in tbody.xpath('./tr/td/text()')]
            fw_li.remove('')
            fws += ' | '.join(fw_li) + '\n'
        return fws

    def save(self, filename='data.xlsx'):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "School Data"  # 设定工作表的名称

        # 写入表头
        header = ['学校', '链接', '地区', '研究生院', '自划线学校', '博士点', '考试方式', '院系所', '专业', '研究方向',
                  '学习方式', '指导老师', '拟招生人数', '考试范围', '备注']
        sheet.append(header)

        for item in self.universitys:
            for i in range(len(item['zhuangye'])):
                fw_str = ''
                for ffw in item['zhuangye'][i]['ksfw']:
                    fw_str += ' '.join(ffw) + ' || '
                if i > 0:
                    row = [
                        '',
                        '',
                        '',
                        '',
                        '',
                        '',
                        item['zhuangye'][i].get('kaoshi_type', ''),
                        item['zhuangye'][i].get('yxs', ''),
                        item['zhuangye'][i].get('zy', ''),
                        item['zhuangye'][i].get('yyfx', ''),
                        item['zhuangye'][i].get('xxfs', ''),
                        item['zhuangye'][i].get('zdls', ''),
                        item['zhuangye'][i].get('nzrs', ''),
                        fw_str,
                        item['zhuangye'][i].get('bz', ''),
                    ]
                else:
                    row = [
                        item.get('name', ''),
                        item.get('href', ''),
                        item.get('area', ''),
                        item.get('yjsy', ''),
                        item.get('zzhx', ''),
                        item.get('bsd', ''),
                        item['zhuangye'][i].get('kaoshi_type', ''),
                        item['zhuangye'][i].get('yxs', ''),
                        item['zhuangye'][i].get('zy', ''),
                        item['zhuangye'][i].get('yyfx', ''),
                        item['zhuangye'][i].get('xxfs', ''),
                        item['zhuangye'][i].get('zdls', ''),
                        item['zhuangye'][i].get('nzrs', ''),
                        fw_str,
                        item['zhuangye'][i].get('bz', ''),
                    ]

                sheet.append(row)

        # 保存工作簿为Excel文件
        print(f"数据已存储到{filename}")
        workbook.save(filename)
