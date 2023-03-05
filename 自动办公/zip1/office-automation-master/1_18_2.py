# Author：刘清九
# Data  ：2020/7/29

"""strip()的正则表达式版本"""
string1 = input('请输入字符串：')


def strip1(string, characters='\t\n '):
    """与strip()功能相同"""
    # 从左检查字符串，并strip
    for i, value in enumerate(string):
        a = string[0]
        if a in characters:
            string = string[1:]
        else:
            break

    # 从右检查字符串，并strip
    for i, value in enumerate(string):
        b = string[-1]
        if b in characters:
            # 切片包括左边，但是不包括右边，一定要记住啊！
            string = string[:-1]
        else:
            break

    print(string)


strip1(string1)
