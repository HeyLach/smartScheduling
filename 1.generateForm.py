# -*- coding: UTF-8 -*-

##################################
# AIGO Smart Scheduling
# team CKC
##################################

import logging
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from logging.handlers import TimedRotatingFileHandler
from libs.common import 讀取Yaml檔, weekSlicing
from libs.excelRelative import genWorkingHourForm

test = False
# 產生待填表單
def proc(start, fin, workHour):
    logger.info(f'產生自 {start} 至 {fin} 的待填工時表單')
    weeks, hours, dates, comments = weekSlicing(start, fin, workHour)
    Path('output').mkdir(exist_ok=True)
    genWorkingHourForm(weeks, hours, dates, comments, '待填工時表單.xlsx')
    logger.info(f'產生待填工時表單完成')

if __name__ == '__main__':
    conf = 讀取Yaml檔('src/config.yml')

    # logger handle
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # set-up console logger
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s - %(message)s'))
    logger.addHandler(console_handler)
    # set-up file logger
    Path(conf['日誌路徑']).parent.mkdir(parents=True, exist_ok=True)
    file_handler = TimedRotatingFileHandler(conf['日誌路徑'], when="midnight", interval=1, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s - %(message)s'))
    logger.addHandler(file_handler)

    startDate = datetime.now() + timedelta(days=(7 - datetime.now().weekday()))
    if test:
        date = '20210401'
    else:
        date = input('請輸入日期, 如: 20211010\n')
    startDate = datetime.strptime(date, '%Y%m%d')
    startDate += timedelta(days=(7 - startDate.weekday()))

    proc(start=startDate.strftime('%Y%m%d'), fin=(startDate+timedelta(days=34)).strftime('%Y%m%d'), workHour=conf['每日正常工時'])

