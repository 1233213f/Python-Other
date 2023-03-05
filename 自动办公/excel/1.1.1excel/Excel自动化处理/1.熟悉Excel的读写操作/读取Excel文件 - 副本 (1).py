import xlwt # pip install xlwt
wb = xlwt.Workbook('虚假用户数据.xls') # Excel的文件对象

sheet = wb.add_sheet('第一个sheet')


head_data = ['姓名','地址','手机号','城市']
for head in head_data:
    sheet.write(0,head_data.index(head),head)


import faker
fake = faker.Faker()
for i in range(1,100):
    sheet.write(i,0,fake.first_name() + ' ' + fake.last_name())
    sheet.write(i,1,fake.address())
    sheet.write(i,2,fake.phone_number())
    sheet.write(i,3,fake.city())


wb.save('虚假用户数据-two.xls')
#%%

import xlrd # pip install xlrd
wb = xlrd.open_workbook('虚假用户数据-two.xls')

sheets = wb.sheets()                 #获取工作表list。
sheet = sheets[0]                  #通过索引顺序获取。

sheet = wb.sheet_by_index(0)         #通过索引顺序获取。

sheet = wb.sheet_by_name('第一个sheet')  #通过名称获取。


rows = sheet.nrows
cols = sheet.ncols
print(rows, cols)


for row in range(rows):
    for col in range(cols):
        print(sheet.cell(row,col).value,end=' , ')
    print('\n')






