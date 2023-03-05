# coding = utf-8
import xlrd
import xlwt

# 读取excel文件
# 打开excel文件，打开工作表
xlsx = xlrd.open_workbook('路径')

# 读取工作簿(数字从0开始，打开第几个sheet)
table = xlsx.sheet_by_index(0)

# 读取某个单元格的数据
# 第一种方式
a = table.cell_value  #参数填写坐标，也是从0开始
# 第二种方式
b = table.cell().cell
print(a, b)

#写入excel文件
#创建文件
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('name')
worksheet.write(1, 1, '内容')
new_workbook.save('路径')