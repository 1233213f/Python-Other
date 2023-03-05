import xlrd
import faker
from xlutils.copy import copy

wb = xlrd.open_workbook('虚假数据.xls',formatting_info=True)   # wb只读，不可操作，对象是:xlrd.sheet.Sheet
xwb = copy(wb)   # xwb可写，对象是:xlwt.Worksheet.Worksheet。
sheet = xwb.get_sheet('第一个sheet')   # xwb无sheet_by_name() 或 sheet_by_index()属性，不能用此来获取指定sheet
rows = sheet.get_rows()
fake = faker.Faker()  # 生成虚假数据

for i in range(len(rows),150): # 从第len(rows)行开始，追加50行数据
    sheet.write(i,0,fake.first_name() + ' ' + fake.last_name())
    sheet.write(i,1,fake.address())
    sheet.write(i,2,fake.phone_number())
    sheet.write(i,3,fake.city())

xwb.save('虚假数据.xls')

