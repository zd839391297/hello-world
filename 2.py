import requests
from bs4 import BeautifulSoup
import xlrd
import xlwt
from xlutils.copy import copy

oldWb = xlrd.open_workbook('test_xlwt.xls')
newWb = copy(oldWb)
booksheet = newWb.get_sheet(0)


#workbook = xlwt.Workbook(encoding='utf-8')
'''booksheet = workbook.add_sheet('test case', cell_overwrite_ok=True)  
booksheet.write(0,0,'期刊名字')
booksheet.write(0,1,'期刊ISSN')
booksheet.write(0,2,'2017-2018最新影响因子')
booksheet.write(0,3,'2017-2018自引率')
booksheet.write(0,4,'五年影响因子')
booksheet.write(0,5, '期刊官方网站')
booksheet.write(0,6,'期刊投稿网址')
booksheet.write(0,7, '是否OA开放访问')
booksheet.write(0,8,'通讯方式')
booksheet.write(0,9,'涉及的研究方向')
booksheet.write(0,10,'出版国家或地区')
booksheet.write(0,11,'出版周期')
booksheet.write(0,12, '出版年份')
booksheet.write(0,13, '年文章数')'''
for i in range(80, 100):
    for n in range(10*i, 10*i+10):
        url = "http://www.letpub.com.cn/index.php?journalid=" + \
            str(n)+"&page=journalapp&view=detail"
        # code="utf-8"
        r = requests.get(url)
        # r.raise_for_status()
        r.encoding = "UTF-8"
        soup = BeautifulSoup(r.text, 'lxml')
        try:
            soup_1_find = soup.body.div.next_sibling
        except:
            n = n+1
        for i in range(10):
            soup_1_find = soup_1_find.next_sibling
        soup_2_find = soup_1_find.div
        for i in range(12):
            soup_2_find = soup_2_find.next_sibling
        soup_3_find = soup_2_find.div.h2
        for i in range(19):
            soup_3_find = soup_3_find.next_sibling
        soup_4_find = soup_3_find.tbody.tr.next_sibling
        for i in range(14):
            booksheet.write(n, i, soup_4_find.td.next_sibling.string)
            soup_4_find = soup_4_find.next_sibling
    newWb.save("test_xlwt.xls")
