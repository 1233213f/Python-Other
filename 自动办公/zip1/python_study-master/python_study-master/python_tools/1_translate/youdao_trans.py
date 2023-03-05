# -*- coding: utf-8 -*-
# @Time : 2020/3/9 9:02
# @Author : kanmendashu2020
# @File : youdao_trans.py
# @Software: PyCharm
# 百度翻译 http://api.fanyi.baidu.com/api/trans/product/apidoc

import json
import urllib.request
import urllib.parse


def translate(content):
    '''实现有道翻译的接口'''
    youdao_url = 'http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule'
    data = {}

    data['i'] = content
    data['from'] = 'zh-CHS'
    data['to'] = 'ja'
    data['smartresult'] = 'dict'
    data['client'] = 'fanyideskweb'
    data['salt'] = '15837161155145'  # '1525141473246'
    data['sign'] = 'a201e60cce224708e06de7eba99ff308'  # '47ee728a4465ef98ac06510bf67f3023'
    # data['ts'] = '1583716239164'
    # data['bv'] = '7e3150ecbdf9de52dc355751b074cf60'
    data['doctype'] = 'json'
    data['version'] = '2.1'
    data['keyfrom'] = 'fanyi.web'
    data['action'] = 'FY_BY_CLICKBUTTION'
    data['typoResult'] = 'false'
    data = urllib.parse.urlencode(data).encode('utf-8')

    youdao_response = urllib.request.urlopen(youdao_url, data)
    youdao_html = youdao_response.read().decode('utf-8')
    target = json.loads(youdao_html)

    trans = target['translateResult']
    ret = ''
    for i in range(len(trans)):
        line = ''
        for j in range(len(trans[i])):
            line = trans[i][j]['tgt']
        ret += line + '\n'
    return ret


print(translate('天气'))
