
# -*- coding: utf-8 -*-
# 1、创建日期为名的日志
# 2、输出自定义的info和系统的err
# 3、同时显示程序执行的时间和行号

import logging
import logging.handlers
import datetime
import os, sys
import traceback
# 在当前文件夹下，创建日志目录
path = 'workLog/'
# 时间
today = datetime.date.today()
formatted_today = today.strftime('%y%m%d')
glueStr10 = "-" * 10


# 自定义信息输出到日志
def info(info):
    work_log = logging.getLogger("log")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False

    isExists = os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
    today = datetime.date.today()
    formatted_today = today.strftime('%y%m%d')
    # formatted_today = formatted_today.getDate()

    handlers = logging.handlers.RotatingFileHandler(path + "work" + formatted_today + ".log", 'a', maxBytes=0,
                                                    backupCount=1, encoding='utf-8')
    # 获取被调用函数的位置信息
    funcPath = sys._getframe().f_back.f_code.co_filename  # 当前调用位置所在的文件名
    if '/' in funcPath:
        funcPath = funcPath.split('/')[-1]
    if '\\' in funcPath:
        funcPath = funcPath.split('\\')[-1]
    funcName = sys._getframe().f_back.f_code.co_name  # 获取调用函数名
    lineNumber = str(sys._getframe().f_back.f_lineno)  # 获取调用函数的行号
    # 获取被调用函数的位置信息

    handlers.setFormatter(
        logging.Formatter(
            '%(asctime)s - ' + funcPath + '.' + funcName + '[line:' + lineNumber + '] - %(levelname)s: ----------%(message)s '))
    # handlers.setFormatter(logging.Formatter(funcName + lineNumber))
    work_log.addHandler(handlers)
    work_log.setLevel(logging.INFO)
    work_log.info(glueStr10 + info)
    work_log.removeHandler(handlers)
    return work_log


# 错误信息输出到日志
def err():
    work_log = logging.getLogger("err")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False

    isExists = os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)
    # formatted_today = formatted_today.getDate()
    handlers = logging.handlers.RotatingFileHandler(path + "work" + formatted_today + ".log", 'a', maxBytes=0,
                                                    backupCount=1, encoding='utf-8')
    # 获取被调用函数的位置信息
    funcPath = sys._getframe().f_back.f_code.co_filename  # 当前调用位置所在的文件名
    if '/' in funcPath:
        funcPath = funcPath.split('/')[-1]
    if '\\' in funcPath:
        funcPath = funcPath.split('\\')[-1]
    funcName = sys._getframe().f_back.f_code.co_name  # 获取调用函数名
    lineNumber = str(sys._getframe().f_back.f_lineno)  # 获取调用函数的行号
    # 获取被调用函数的位置信息

    handlers.setFormatter(
        logging.Formatter(
            '%(asctime)s - ' + funcPath + '.' + funcName + '[line:' + lineNumber + '] - %(levelname)s: %(message)s'))

    # handlers.setFormatter(
    #     logging.Formatter(
    #         '%(asctime)s - ' + funcPath + '.' + funcName + '[line:' + lineNumber + '] - %(levelname)s: %(message)s'))
    work_log.addHandler(handlers)
    work_log.exception(glueStr10 + "Exception Logged")
    work_log.removeHandler(handlers)
