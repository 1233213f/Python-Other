from xlutils.copy import copy
import xlrd
import xlwt

# 第一步：打开文件或者创建文件
workbook1 = xlrd.open_workbook('D:/0304.xls', formatting_info=True)
work_sheet = workbook1.sheet_by_index(0)

workbook2 = copy(workbook1)
work_sheet2 = workbook2.get_sheet(0)
# 设置单元格字体格式
style = xlwt.XFStyle()
font = xlwt.Font()
font.name = '微软雅黑'
font.bold = True
font.height = 220
style.font =font
# 设置单元格格式
borders = xlwt.Borders()
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN
borders.left = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
style.borders = borders

# 设置对齐
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
style.alignment = alignment

# 填充单元格数据
work_sheet2.write(1, 1, 10, style)
work_sheet2.write(2, 1, 10, style)
work_sheet2.write(3, 1, 10, style)
work_sheet2.write(4, 1, 10, style)
work_sheet2.write(5, 1, 10, style)
workbook2.save('D:/填写.xls')