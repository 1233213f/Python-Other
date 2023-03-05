# Author：刘清九
# Data  ：2020/8/1

"""
多重剪贴板,可以根据关键字调出相应的剪贴板内容

save <keyword> 将剪贴板内容关联到关键字
<keyword>       将对应内容加载到剪贴板
list            加载所有内容到剪贴板
delete <keyword> 删除关键字对应的内容
"""

import sys
import pyperclip
import shelve

mcb_shelf = shelve.open('mcb')

# 保存剪贴板内容
if len(sys.argv) == 3:
    if sys.argv[1].lower() == 'save':
        mcb_shelf[sys.argv[2]] = pyperclip.paste()
    elif sys.argv[1].lower() == 'delete':
        del mcb_shelf[sys.argv[2]]


# 列出关键字或者加载内容
elif len(sys.argv) == 2:
    if sys.argv[1] == 'list':
        pyperclip.paste(list(mcb_shelf.keys()))
    elif sys.argv[1] in mcb_shelf:
        pyperclip.paste(mcb_shelf[sys.argv[1]])

mcb_shelf.close()
