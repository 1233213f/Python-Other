# -*- coding: UTF-8 -*-
# 公众号  ：  技术TA说
# 代码作者  ：  看门大叔
# 创建时间  ：  2020/3/11  21:48
# 文件名称  :  9x9_live.py
# 开发工具  :  PyCharm

import requests, re, os


def crawl_love_image():
    print("正在抓取我爱你图片...")
    image_n = 1
    for i in range(1, 22):
        url = "http://tieba.baidu.com/p/3108805355?pn={}".format(i)
        response = requests.get(url)
        html = response.text
        pattern = re.compile(
            r'<div.*?class="d_post_content j_d_post_content.*?">.*?<img class="BDE_Image" src="(.*?)".*?>.*?</div>',
            re.S)
        image_url = re.findall(pattern, html)

        for j, data in enumerate(image_url):
            pics = requests.get(data)
            pic_path = 'image'
            mkdir(pic_path)
            fq = open(pic_path + '\\' + str(i) + "_" + str(j) + '.jpg', 'wb')  # 下载图片，并保存和命名
            fq.write(pics.content)
            fq.close()
            print('这是第{}张照片'.format(image_n))
            image_n += 1
    print("图片抓取完成")


def mkdir(path):
    folder = os.path.exists(path)

    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs创建文件时如果路径不存在会创建这个路径
        print("---new folder...---")
        print("---OK---")
    else:
        print("正在保存图片中...")


crawl_love_image()
