import xlrd
from datetime import date,datetime

data =  xlrd.open_workbook('联系人.xls')  # 打开指定的excel文件

sheet_name1  =  data.sheet_names()    # 获取所有sheet名
sheet_name2  =  data.sheet_names()[1] # 根据下标获取sheet名
print(sheet_name1)
print(sheet_name2)

sheet2 = data.sheet_by_index(1)   # 通过索引获取sheet2的名称，同时获取列数、行数
print("sheet2名称：{}\nsheet2列数:{}\nsheet2行数:{}".format(sheet2.name,sheet2.ncols,sheet2.nrows))


sheet1 = data.sheet_by_name('银行2') # 通过sheet名，获取行和列
print(sheet1.row_values(3))
print(sheet1.col_values(3))

print(sheet1.cell(1,0).value)  # 获取指定单元格的值（第2行第1列）
print(sheet1.cell(1,0).ctype)  # 获取指定单元格的数据格式.ctype:0 empty,1 string,2 number,3 date,4 boolean,5 error


if sheet1.cell(3,6).ctype == 3:  #  获取单元格内容为日期类型的方式
    print(sheet1.cell(3,6).value)
    date_value = xlrd.xldate_as_tuple(sheet1.cell(3,6).value,data.datemode)
    print(date_value)
    print(date(*date_value[:3]))
    print(date(*date_value[:3]).strftime("%Y/%m/%d"))
    print(date(*date_value[:3]).strftime("%y/%m/%d"))

# 获取单元格为number的方式（转为整型）
if sheet1.cell(3,5).ctype == 2:
    print(sheet1.cell(3,5).value)
    num_value = int(sheet1.cell(3,5).value)
    print(num_value)

# 获取合并单元格的属性
