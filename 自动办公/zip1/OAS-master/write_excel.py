import xlwt
import faker #faker是一个虚假信息生成库，用于测试
wb = xlwt.Workbook() # 新建一个Workbook对象
sheet  = wb.add_sheet('第一个sheet')
head_data = ['姓名','地址','手机号','城市']
for head in head_data:
    sheet.write(0,head_data.index(head),head) #头部数据从第0行开始写入
fake = faker.Faker()
for i in range(1,100):
    sheet.write(i,0,fake.first_name()+' '+ fake.last_name())
    sheet.write(i,1,fake.address())
    sheet.write(i,2,fake.phone_number())
    sheet.write(i,3,fake.city())

wb.save('虚假数据.xls')