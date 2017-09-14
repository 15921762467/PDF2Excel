# -*- coding:utf-8 -*-
import xlsxwriter

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator

# 获取文档对象
fp = open('test.pdf', 'rb')
# 创建一个与文档相关联的解释器
parser = PDFParser(fp)
# PDF文档对象
doc = PDFDocument(parser)
# 链接解释器和文档对象
parser.set_document(doc)
# doc.set_parser(parser)
# 初始化文档
# doc.initialize("")
# 创建PDF资源管理器
resource = PDFResourceManager()
# 参数分析器
laparam = LAParams()
# 创建一个聚合器
device = PDFPageAggregator(resource, laparams=laparam)
# 创建PDF页面解释器
interperter = PDFPageInterpreter(resource, device)
# 使用文档对象得到页面集合

list_of_col = []

for page in PDFPage.create_pages(doc):
    # 使用页面解释器来读取
    interperter.process_page(page)
    # 使用聚合器来获取内容
    layout = device.get_result()
    print type(layout)
    print layout
    L = []
    for out in layout:
        if hasattr(out, "get_text"):
            # print out.get_text()
            L.append(out.get_text())
    # print L


    for items in L:
        col = []
        for item in items.split('\n'):
            print item
            if item:
                col.append(item)
        list_of_col.append(col)
    print list_of_col

# 建立文件以及sheet
workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
# 设置粗体, 默认是False
# bold = workbook.add_format({'bold': True})

# Add a number format for cells with money. 定义数字格式
# money = workbook.add_format({'num_format':'$#,##0'})

# Write some data headers. 带自定义粗体bold格式写表头
# worksheet.write('A1', 'Item', bold)
# worksheet.write('B1', 'Item', bold)

# Start from the fist cell below the headers
# row = 1
# col = 1
row = 0
col = 0

for col_items in list_of_col:
    row = 0
    for item in col_items:
        worksheet.write(row, col, item)  # 带默认格式写入
        # worksheet.write(row, col+1, cost, money) #带自定义格式写入
        row += 1
    col += 1

# Write a total using a formula
# worksheet.write(row, 0, 'Total', bold)
# worksheet.write(row, 1, '=SUM(B2:B5)', money)

workbook.close()