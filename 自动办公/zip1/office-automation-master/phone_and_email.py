# Author：刘清九
# Data  ：2020/7/24

import re
import pyperclip

content = str(pyperclip.paste())

# 建立号码正则表达式
phone_regex = re.compile(r'''
            (\d{3}|\(\d{3}\))?      # 区号
            (-|\s|\.)?              # 分隔符
            (\d{3})
            (-|\s|\.)?              # 分隔符
            (\d{4})
            ''', re.VERBOSE)

# 建立邮箱正则表达式
email_regex = re.compile(r'''(
            [a-zA-Z0-9]+            # 用户名
            @
            [a-z0-9]+      
            (\.com)
            )''', re.VERBOSE)

# 进行电话号码和邮箱的匹配
result = phone_regex.findall(content)
for i, groups in enumerate(result):
    result[i] = '-'.join([groups[0], groups[2], groups[4]])

email_result = email_regex.findall(content)
email_matches = []
for groups in email_result:
    email_matches.append(groups[0])

# 把结果粘贴到剪贴板
phone = '\t'.join(result)
email = ' '.join(email_matches)

pyperclip.copy(phone + email)
