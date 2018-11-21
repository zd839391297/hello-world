# 本程序适用于Letpub网站的数据，数据采用一次保存

# 导入库requests、beautifulsoup、xlrd（用于读execl）、xlutils(用于写execl)
import requests
from bs4 import BeautifulSoup
import xlrd
from xlutils.copy import copy

# 打开已经存在的execl表格文件
# 如果还没有创建execl表，可以参考以下代码创建
# 也可以手工创建，但是要注意程序已经默认空出了第一行的表头
# import xlwt
# workbook = xlwt.Workbook(encoding='utf-8')
# booksheet = workbook.add_sheet('test case', cell_overwrite_ok=True)
# booksheet.write(0,0,'期刊名字')
oldWb = xlrd.open_workbook('1.xls')

# 在内存区复制一份原表格文件
newWb = copy(oldWb)

# 打开表格的工作表
# 如果不是第一张表格，应当更改sheet_number
sheet_number = 0
booksheet = newWb.get_sheet(sheet_number)

# 定义requests制作的html页面字符编码，如无特殊要求，可不修改设置
code = "utf-8"

# 输入要查询的数据范围，该网站的数据范围为0-10567
start = 1
end = 3

# 开始循环爬取数据
for n in range(start, end):

    # 获取网站的html码
    url = "http://www.letpub.com.cn/index.php?journalid=" + \
        str(n)+"&page=journalapp&view=detail"
    r = requests.get(url)
    # r.raise_for_status()
    r.encoding = code

    # 转换代码，格式为lxml,格式可更改
    soup = BeautifulSoup(r.text, 'lxml')

    # 信息定位
    soup_1_find = soup.body.div.next_sibling
    for i in range(10):
        soup_1_find = soup_1_find.next_sibling
    soup_2_find = soup_1_find.div
    for i in range(12):
        soup_2_find = soup_2_find.next_sibling
    soup_3_find = soup_2_find.div.h2

    # 检查是否存在该条数据，因网站偶尔会存在已删除的记录
    try:
        for i in range(19):
            soup_3_find = soup_3_find.next_sibling
    except AttributeError:
        continue

    # 有用的信息域
    soup_4_find = soup_3_find.tbody.tr.next_sibling

    # 写入名称到excel相应列
    soup_5_find = soup_4_find.span
    booksheet.write(n, 0, soup_5_find.a.string)
    booksheet.write(n, 1, soup_5_find.font.string)

    # 写入除名称之外的信息到excel相应列
    
    for i in range(1, 15):
        booksheet.write(n, i, soup_4_find.td.next_sibling.string)
        soup_4_find = soup_4_find.next_sibling

    # 写入中科院SCI期刊分区(最新版本)到execl相应列
    soup_6_find = soup_4_find.next_sibling.td.next_sibling.table.tr.next_sibling.td
    soup_7_find = soup_6_find.next_sibling
    soup_8_find = soup_7_find.next_sibling
    soup_9_find = soup_8_find.next_sibling

    # 大类学科
    booksheet.write(n, 15, soup_6_find.next_element)

    # 大类学科分区
    booksheet.write(n, 16, soup_6_find.span.string)

    # 小类学科
    booksheet.write(n, 17, soup_7_find.table.tr.td.get_text(strip=True))

    # 小类学科分区
    booksheet.write(n, 18, soup_7_find.table.tr.td.next_sibling.string)

    # 是否为top期刊
    booksheet.write(n, 19, soup_8_find.string)

    # 是否为综述期刊
    booksheet.write(n, 20, soup_9_find.string)

    # 写入SCI期刊coverage
    soup_10_find = soup_4_find.next_sibling.next_sibling
    booksheet.write(n, 21, soup_10_find.td.next_sibling.a.string)

    # 写入PubMed Central (PMC)链接
    soup_11_find = soup_10_find.next_sibling
    booksheet.write(n, 22, soup_11_find.td.next_sibling.a.string)

# 保存此次爬取的数据
newWb.save("1.xls")
