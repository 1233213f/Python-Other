#%% md

# Python自动化工资结算和统计报表

'''每个月的2号，各个部门的领导都会将上个月的员工迟到次数和奖金，以及员工的其他信息统计成Excel信息，统一发给财务部的你整。你有一天的时间去完成这些报表的整理和统计，然后交给领导检查和发放工资。

为了解放双手，有更多的自由时间去看书和充实自己，现在你要做一个Python自动化工资结算和统计报表的程序，让程序快速的计算出每个人的工资，并将统计信息结合模板，生成“xxxx年xx月各部门员工数据总览”。

规定：
- 迟到一次扣20，一个月最多扣200

#%%'''

import datetime
import xlrd,xlwt
from xlutils.copy import copy


department = ['技术部','推广部','客服部','行政部','财务部']
template_name = "月结统计模板.xls"
today_datetime = datetime.datetime.now()
need_process_xls = []
for dep in department:
    xls_name = "{}-{}.xls".format(datetime.datetime.now().strftime("%Y-%m"), dep)
    need_process_xls.append(xls_name)
print(need_process_xls)

#%%

def process_xls_return_data(xls_name):
    wb = xlrd.open_workbook(xls_name)
    wb_sheet = wb.sheet_by_index(0)
    xwb = copy(wb)
    xwb_sheet = xwb.get_sheet('sheet1')
    rows = wb_sheet.nrows
    for row in range(1,rows):
        bm = wb_sheet.cell(row,0).value
        xm = wb_sheet.cell(row,1).value
        gh = wb_sheet.cell(row,2).value
        gz = wb_sheet.cell(row,3).value
        cdcs = wb_sheet.cell(row,4).value
        jj = wb_sheet.cell(row,5).value
        sfgz = gz-(cdcs*20)+jj # 实发工资 = 工资  -  （迟到次数*20）  +  奖金
        print(bm,xm,gh, gz,cdcs,jj,sfgz)
        xwb_sheet.write(row,6, sfgz)
        xwb.save(xls_name)

for cls in need_process_xls:
    process_xls_return_data(cls)

#%%

def process_xls_return_data(xls_name):
    staff_number = 0 # 总人数字段
    cd_number = 0 # 迟到人数字段
    jj_number = 0 # 拿奖金人数字段
    total_pay = 0 # 总应发工资字段
    wb = xlrd.open_workbook(xls_name,formatting_info=True)
    wb_sheet = wb.sheet_by_index(0)
    xwb = copy(wb)
    xwb_sheet = xwb.get_sheet('sheet1')
    rows = wb_sheet.nrows
    for row in range(1,rows):
        staff_number = staff_number + 1 # 每个数据都是一个员工，直接+1
        bm = wb_sheet.cell(row,0).value # 部门名称
        xm = wb_sheet.cell(row,1).value
        gh = wb_sheet.cell(row,2).value
        gz = wb_sheet.cell(row,3).value
        cdcs = wb_sheet.cell(row,4).value
        if cdcs>0:  # 如果迟到次数大于0，则是迟到过的人，迟到人数+1
            cd_number = cd_number + 1
        jj = wb_sheet.cell(row,5).value
        if jj>0:  # 如果奖金大于0，则是获得了奖金的人，拿奖金人数+1
            jj_number = jj_number + 1
        sfgz = gz-(cdcs*20)+jj # 实发工资 = 工资  -  （迟到次数*20）  +  奖金
        total_pay = total_pay + sfgz # 将所有的实发工资加到一起，就是总的实发工资
        print(bm,xm,gh, gz,cdcs,jj,sfgz)
        xwb_sheet.write(row,6, sfgz)
    xwb.save(xls_name)
    print([bm, staff_number, cd_number, jj_number, total_pay])
    return [bm, staff_number, cd_number, jj_number, total_pay] # 最后将部门 总人数  总迟到人数  总拿奖金人数  总实发工资做成列表，一并返回


all_info = []
for cls in need_process_xls:
    one_partment = process_xls_return_data(cls)
    all_info.append(one_partment) # 将函数的返回值，放到列表中，就得到了所有部门的统计信息

wb = xlrd.open_workbook(template_name,formatting_info=True)
wb_sheet = wb.sheet_by_index(0)
xwb = copy(wb)
xwb_sheet = xwb.get_sheet('Sheet1')
current_row = wb_sheet.nrows
year_month = datetime.datetime.now().strftime("%Y-%m")
title = "{}-各部门员工数据总览"

def create_style(): # 定义字体格式，返回一个字体大小24，垂直居中 水平居中 宋体格式 的样式
    style = xlwt.XFStyle()
    fnt = xlwt.Font()                        # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = u'宋体'
    fnt.height = 20*24
    alignment = xlwt.Alignment()
    alignment.horz = 0x02          # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.vert = 0x01          # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    style.font = fnt
    style.alignment = alignment
    return style

xwb_sheet.write(0,0, title.format(year_month) , create_style()) # 写入头部标题，内容是“xxxx-xx-各部门员工数据总览”，样式是 宋体 大小24 垂直水平居中

for info in all_info: # 循环所有的部门信息，全部写入到文件中
    xwb_sheet.write(current_row,0, info[0])
    xwb_sheet.write(current_row,1, info[1])
    xwb_sheet.write(current_row,2, info[2])
    xwb_sheet.write(current_row,3, info[3])
    xwb_sheet.write(current_row,4, info[4])
    current_row = current_row + 1
xwb.save(title.format(year_month)+'.xls') # 最后保存，文件名是 xxxx-xx-各部门员工数据总览.xls

#%%


