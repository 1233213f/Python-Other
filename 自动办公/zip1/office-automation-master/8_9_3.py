"""打开文件夹所有.txt文件，查找匹配用户提供正则表达式的所有行，结果打印到屏幕上"""

import os
import re


def function1():
    # 查找所有.txt的文件
    files = os.listdir(os.getcwd())
    text_files = []
    for i in files:
        if i.endswith('.txt'):
            text_files.append(i)
    print(text_files)

    # 打开文件，并将行读入列表
    lines = []
    for text_file in text_files:
        with open(text_file) as f:
            lines.extend(f.readlines())

    regex = re.compile(input('请输入要匹配的内容：'))
    for line in lines:
        mo = regex.search(line)
        if mo is not None:
            print(line)


function1()
