#%%

import xlrd # pip install xlrd xlutils xlwt
from xlutils.copy import copy

#%%

wb = xlrd.open_workbook('new-虚假用户数据.xls',formatting_info=True)

#%%

xwb = copy(wb)

#%%

sheet = xwb.get_sheet('第一个sheet')

#%%

rows = sheet.get_rows()

#%%

len(rows)

#%%

import faker
fake = faker.Faker()
for i in range(len(rows),len(rows)+200):
    sheet.write(i,0,fake.first_name() + ' ' + fake.last_name())
    sheet.write(i,1,fake.address())
    sheet.write(i,2,fake.phone_number())
    sheet.write(i,3,fake.city())

#%%

xwb.save('new-虚假用户数据.xls')

#%%


