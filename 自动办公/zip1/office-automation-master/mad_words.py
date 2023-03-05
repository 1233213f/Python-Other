# Author：刘清九
# Data  ：2020/8/2
import re

with open('tianci.txt') as f:
    content = f.read()
print(content)

# 找到并替换单词
while True:
    old_word_regex = re.compile(r'noun|verb|adjective')
    mo = old_word_regex.search(content)
    if mo is None:
        break
    else:
        old_word = mo.group()
        new_word = input('Enter a %s:' % old_word)
        content = content.replace(old_word, new_word, 1)

print(content)

# 写入新文件
with open('new_tianci', 'w') as f1:
    f1.write(content)
