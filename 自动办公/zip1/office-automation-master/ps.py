# Author：刘清九
# Data  ：2020/7/27
"""一个简单的口令保管箱"""
import sys
import pyperclip

passwords = {"luggage": '12345', 'email': 'faofofno', "blog": 'fnqp;nbopqeoitt092'}

# 确认是否输入账号参数
if len(sys.argv) < 2:
    print('请先输入文件名，再输入账号！')
    sys.exit()

account = sys.argv[1]

# 查询账号密码，并返回到剪贴板
if account in passwords:
    pyperclip.copy(passwords[account])
    print('口令已经复制到剪贴板。')
else:
    print('你没有存储这个账号口令！')


