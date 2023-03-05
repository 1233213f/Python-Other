# Author：刘清九
# Data  ：2020/7/29

'''确定一个字符串是否是强口令，可能需要一个或者多个正则表达式'''

import re


def case1(pass_ward):
    '''判断是否包含小写字母'''
    characters1 = []
    for i in range(26):
        characters1.append(chr(ord('a') + i))

    for i in pass_ward:
        if i in characters1:
            return True


def case2(pass_ward):
    '''判断是否包含大写字母'''
    characters2 = []
    for i in range(26):
        characters2.append(chr(ord('A') + i))

    for i in pass_ward:
        if i in characters2:
            return True


def case3(pass_word):
    '''判断是否包含数字'''
    for i in range(10):
        if str(i) in pass_word:
            return True


def strong_pw():
    '''判断密码是否为强口令'''
    # 确认密码长度不少于8位
    password = input('请输入密码：')
    if len(password) < 8:
        print('password is wrong.')
    else:
        if case1(password) and case3(password) and case2(password):
            print('passward is right.')
        else:
            print('password is wrong.')

strong_pw()
