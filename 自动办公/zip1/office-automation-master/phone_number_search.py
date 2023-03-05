# Author：刘清九
# Data  ：2020/7/23

import re

phone_number = re.compile(r'(\d\d\d)-(\d\d\d-\d\d\d\d)')
mo = phone_number.search('my number is 415-444-4242')
print("phone number found:" + mo.group())
print(mo.group(1))
print(mo.groups())

hero = re.compile(r'batman|spiderman')
result = hero.search('batman and spiderman')
print(result.group())

another_result = hero.search('spiderman and batman')
print(another_result.group())

bat_regex = re.compile(r'bat(man|bat|coter|moblie)')
bat_mo = bat_regex.search("batmoblie lost a wheel")
print(bat_mo.group())
print(bat_mo.group(1))

bat_regex2 = re.compile(r'bat(wo)?man')
mo2 = bat_regex2.search('batman is beautiful')
print(mo2.group())

phone_regex = re.compile(r'(\d\d\d-)?(\d\d\d-\d\d\d\d)')
mo3 = phone_regex.search('my phone number is 553-3382')
print(mo3.group())
print(mo3.group(2))

bat_regex3 = re.compile(r'bat(wo)*man')
mo4 = bat_regex3.search('batwowowoman')
print(mo4.group())

bat_regex6 = re.compile(r'bat(wo)+man')
mo6 = bat_regex6.search('batman')
if mo6 is None:
    print('True')


phone_number = re.compile(r'(\d{3})-(\d{3})-(\d{4})')
mo = phone_number.search('work: 444-532-5433, home: 335-353-7835')
# print(mo.group())
mo8 = phone_number.findall('work: 444-532-5433, home: 335-353-7835')
print(mo8)