#%%

import xlwt
import faker
import random
import datetime


#%%

def create_excel_file(filename,department):
    wb = xlwt.Workbook(filename)
    sheet = wb.add_sheet('sheet1')
    fake = faker.Faker("zh_CN")
    head_data = ['部门','姓名','工号','薪资（元）','迟到次数（次）','奖金（元）','实发工资']
    for head in head_data:
        sheet.write(0,head_data.index(head),head)
    for i in range(1, random.randint(5,100)):
        sheet.write(i,0, department)
        sheet.write(i,1, fake.last_name()+fake.first_name())
        sheet.write(i,2, "G{}".format(random.randint(1,1000)))
        sheet.write(i,3, random.randint(4000,16000))
        sheet.write(i,4, random.choice([0,0,0,1,2,3,4]))
        sheet.write(i,5, random.choice([200,300,400,500,600,700,800,900]))
    wb.save(filename)

#%%

department_name = ['技术部','推广部','客服部','行政部','财务部']
for dep in department_name:
    xls_name = "{}-{}.xls".format(datetime.datetime.now().strftime("%Y-%m"), dep)
    create_excel_file(xls_name,dep)
    print(xls_name," 新建完成")

#%%


