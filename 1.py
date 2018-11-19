import requests
import numpy as np
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
n=33
while n<40:
    url="http://www.letpub.com.cn/index.php?journalid="+str(n)+"&page=journalapp&view=detail"
    code="UTF-8"

    r=requests.get(url)
    r.raise_for_status()
    r.encoding=code
    soup = BeautifulSoup(r.text,'lxml')
    soup_body=soup.body
    soup_1_find=soup_body.div.next_sibling
    i=0
    while i <10:
        soup_1_find=soup_1_find.next_sibling
        i=i+1
    soup_2_find=soup_1_find.div
    i=0
    while i<12:
        soup_2_find=soup_2_find.next_sibling
        i=i+1
    soup_3_find=soup_2_find.div.h2
    i=0
    while i<19:
        soup_3_find=soup_3_find.next_sibling
        i=i+1
    soup_4_find=soup_3_find.tbody.tr.next_sibling
    i=0
    while i<14:
        booksheet.write(n,i,soup_4_find.td.next_sibling.string)
        soup_4_find=soup_4_find.next_sibling
        i=i+1
    n=n+1
newWb.save("test_xlwt.xls")
