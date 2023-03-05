# Author：刘清九
# Data  ：2020/7/18

import pyperclip

content = pyperclip.paste()
content_list = content.split("\n")
for i in range(len(content_list)):
    content_list[i] = content_list[i].strip(".*\r ").split(' ')
# print(content_list)
zipped = zip(*content_list)

for item in zipped:
    for m, n in enumerate(item):
        if m != 5:
            print(n.rjust(10), end="\t")
        else:
            print(n.rjust(10), end='\n')

# 恭喜任务圆满完成
