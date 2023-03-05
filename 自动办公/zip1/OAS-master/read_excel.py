import xlrd
wb =  xlrd.open_workbook('虚假数据.xls')
sheet = wb.sheet_by_index(0)
"""
也可用以下方式获取sheet对象
# sheet =  wb.sheet[0]
# sheet = wb.sheet_by_name('第一个sheet')

"""
rows = sheet.nrows #获取总行数
cols = sheet.ncols #获取总列数

for row in range(rows):
    for col in range(cols):
        print(sheet.cell(row,col).value,end=',')
    print('\n')

