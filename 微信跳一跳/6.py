import requests

session = requests.session()

import logging

import requests

logging.basicConfig(level=logging.DEBUG,

format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',

datefmt='%a, %d %b %Y %H:%M:%S',

filename='myapp.log',

filemode='w')

logging.debug(session.get('http://www.qq.com'))

logging.debug(session.get('http://www.qq.com'))