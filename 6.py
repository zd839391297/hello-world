# 本程序适用于查询名称，数据采用一次保留

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
oldWb = xlrd.open_workbook('test_xlwt.xls')

# 在内存区复制一份原表格文件
newWb = copy(oldWb)

# 打开表格的工作表
# 如果不是第一张表格，应当更改sheet_number
sheet_number = 0
booksheet = newWb.get_sheet(sheet_number)

# 定义requests制作的html页面字符编码，如无特殊要求，可不修改设置
code = "utf-8"

# 输入要查询的数据范围，该网站的数据范围为0-10567
start = 0
end = 10567

# 开始循环爬取数据
for n in range(start, end):

    # 获取网站的代码
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

    # 将数据写入excel表中
    soup_4_find = soup_3_find.tbody.tr.next_sibling.span
    booksheet.write(n, 0, soup_4_find.a.string)
    booksheet.write(n, 1, soup_4_find.font.string)

# 保存此次爬取的数据
newWb.save("test_xlwt.xls")
