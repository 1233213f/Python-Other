# Author：刘清九
# Data  ：2020/7/25

import os, shelve

my_files = ['file1.txt', "file2.docx"]
for file in my_files:
    print(os.path.join('C:\\User\\ninel', file))

os.getcwd()
# os.makedirs(r'C:\Users\ninel\OneDrive\桌面\new_file\new_test\readmes')
os.path.abspath('.idea')

os.path.relpath(r'C:\\Users\\ninel\\OneDrive\\桌面\\office automation\\idea',
                r'C:\Users\ninel\OneDrive\桌面\new_file\new_test\readmes')
path1 = r'C:\Users\ninel\OneDrive\桌面\office automation'
a = os.path.dirname(r'C:\Users\ninel\OneDrive\桌面\office automation')
b = os.path.basename(r'C:\Users\ninel\OneDrive\桌面\office automation')
print((a, b))
c = path1.split(os.path.sep)
print(c)
d = os.path.split(path1)
print(d)
e = os.path.getsize(path1)
print(e)
f = os.listdir(path1)
print(f)

shelf_file = shelve.open('mydate')
cats = ['marry', 'karry']
shelf_file['cats'] = cats
# shelf_file.close()
print(list(shelf_file.keys()))
print(shelf_file['cats'])