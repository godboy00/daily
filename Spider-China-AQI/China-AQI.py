# -*- coding: utf-8 -*-
from selenium import webdriver
from BeautifulSoup import BeautifulSoup
import time, re
import xlrd
import xlwt
from xlutils.copy import copy
import random


class Web:
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.base_url = "http://datacenter.mep.gov.cn:8099/ths-report/report!list.action?xmlname=1463473852790"
        self.driver.get(self.base_url)
        self.init_exc()

    def init_exc(self):
        wb = xlwt.Workbook()
        sheet = wb.add_sheet('sheet1', cell_overwrite_ok=True)
        title_list = [(0, 'city'), (1, 'AQI'), (2, 'Prime'), (3, 'Level'), (4, 'date'), (5, 'Sever')]
        for (a, b) in title_list:
            sheet.write(0, a, b)
        wb.save('data_20050101-20131231-1414.xls')

    def do(self):
        self.get_page()
        time.sleep(random.randint(3,6))
        try:
            self.driver.find_element_by_link_text(u"下一页").click()
            self.do()
        except:
            page1 = self.driver.find_element_by_id("inPageNo").getText()
            page = int(page1)+1
            return 0

    def get_page(self):
        html_source = self.driver.page_source
        soup = str(BeautifulSoup(html_source))
        # print soup

        for i in range(30):
            n = i + 1
            city = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="0" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)
            date = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="1" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)
            value = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="2" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)
            prime = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="3" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)
            level = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="4" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)
            sever = re.findall(
                '<td rowid="%s" mergecol="-1" mergerow="-1" colid="5" rowspan="1" colspan="1" style="text-align:center; overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">(.+)</td>'%n,
                soup)

            if len(city) == 0:
                return 0
            if len(value) == 0:
                value = [0]
            if len(prime) == 0:
                prime = [0]
            if len(level) == 0:
                level = [0]
            if len(sever) == 0:
                sever = [0]

            title_list = [(0, str(city[0]).decode('utf-8')), (1, int(value[0])), (2, str(prime[0]).decode('utf-8')), (3, str(level[0]).decode('utf-8')), (4, str(date[0]).decode('utf-8')), (5, str(sever[0]).decode('utf-8'))]
            rb = xlrd.open_workbook("data_20050101-20131231-1414.xls")
            wb = copy(rb)
            ws = wb.get_sheet(0)
            sh = rb.sheet_by_index(0)
            count = sh.nrows
            for (a, b) in title_list:
                ws.write(count, a, b)
            wb.save("data_20050101-20131231-1414.xls")
        return 0

if __name__ == "__main__":
    a = Web()
    time.sleep(20)
    if a.do() == 0:
        print "Done"

