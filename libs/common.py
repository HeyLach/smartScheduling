#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
import codecs
import json
import yaml
import requests as rq
from pathlib import Path
from datetime import datetime, timedelta

debug = False

def 讀取Yaml檔(ymlfile):
    try:
        with open(ymlfile, 'r') as ymlf:
            return yaml.safe_load(ymlf)
    except Exception as e:
        ans = input('The config file is not there, do you want me to create a new one? (Y/N)\n').lower()
        if ans == 'y' or ans == 'yes':
            content = {}
            with open(ymlfile, 'w') as ymlf:
                yaml.dump(content, ymlf, encoding='utf-8', allow_unicode=True)
            return content
        else:
            raise BaseException('[Error] Uable to read the configuration file.')


def weekSlicing(start, end, work_hour):
    getHolidays()
    check = dict()
    (weeks, hours, dates, comments) = list(), list(), list(), list()
    week, hour, date, comment = list(), list(), list(), list()
    current = datetime.strptime(start, '%Y%m%d')

    # 分週與每週新增工時
    while current != datetime.strptime(end, '%Y%m%d') + timedelta(days=1):
        year = current.strftime('%Y')
        if year not in check.keys():
            with open(f'src/holiday_{year}.json', 'r') as f:
                check[year] = json.loads(f.read())
        if check[year][current.strftime('%Y%m%d')]['星期'] == '一' and week:
            weeks.append(week)
            hours.append(hour)
            dates.append(date)
            comments.append(comment)
            week, hour, date, comment = list(), list(), list(), list()
        if check[year][current.strftime('%Y%m%d')]['星期'] == '日':
            current += timedelta(days=1)
            continue
        week.append(current)
        hour.append(work_hour * 60 * (not check[year][current.strftime('%Y%m%d')]['是否放假']))
        date.append(check[year][current.strftime('%Y%m%d')]['星期'])
        comment.append(check[year][current.strftime('%Y%m%d')]['備註'])
        current += timedelta(days=1)
    if week:
        weeks.append(week)
        hours.append(hour)
        dates.append(date)
        comments.append(comment)
    return weeks, hours, dates, comments


def getHolidays():
    data_url = 'https://data.gov.tw/dataset/14718'
    json_files = re.findall(r'json:"([^"]+md5_url=[^"]+)"', rq.get(data_url).text)
    Path('src').mkdir(exist_ok=True)
    for url in json_files:
        try:
            url = codecs.decode(url, 'unicode-escape')
            content = json.loads(codecs.decode(rq.get(url).text, 'unicode-escape'))
            data = dict()
            for date in content:
                data[date['西元日期']] = dict()
                data[date['西元日期']]['星期'] = date['星期']
                data[date['西元日期']]['備註'] = date['備註']
                data[date['西元日期']]['是否放假'] = True
                if date['是否放假'] == '0':
                    data[date['西元日期']]['是否放假'] = False
            year = content[0]['西元日期'][:4]
            with open(f'src/holiday_{year}.json', 'w') as jsonfile:
                json.dump(data, jsonfile, ensure_ascii=False)
        except:
            pass

# reorders 將訂單依其出貨日期及生產時間排序, 影響排程可行性
# 主key: 出貨時間
# 副key: 生產時間
def reorders(orders):
    # 生產時間排序: 由長至短
    orderList = sorted(orders, key=lambda order: order[-2], reverse=True)
    # 出貨時間排序: 由近到遠
    orderList = sorted(orderList, key=lambda order: order[-4])
    if debug:
        print(f"原訂單排序: {orders}")
        print(f"新訂單排序: {orderList}")
    return orderList

# optimize 最佳化排程
# seq 工班列表
# orders 訂單
# shifts 工班資訊
def optimize(seq, orders, shifts):
    # 讀取換線時間列表, 提供換線時間查詢
    conf = 讀取Yaml檔(Path(os.path.dirname(__file__)).parent / 'src' / '換線時間.yml')
    # 依序讀取工班
    for shift in seq:

        # 預排清單 該類別之次類別:
        # e.g., 類別:小型訂單
        orderTypes = list(orders.keys())

        # solution 次類別排程
        solution = [shifts[shift]['前筆訂單類型']]

        # 將次類別訂單依出貨時間與生產時間排序
        for type in orderTypes:
            orders[type] = reorders(orders[type])

        # 依據換線時間做訂單排程
        while orderTypes:
            # 計算切換各類型所需換線時間
            setupTimes = dict()
            for nextType in orderTypes:
                previousType = solution[-1]
                setupTimes[nextType] = conf[previousType][nextType]
            if debug:
                print(f'orderTypes: {orderTypes}')
                print(f'setuptimes: {setupTimes}')

            # 選擇換線時間最短次類別
            minType = min(setupTimes, key=setupTimes.get)
            solution.append(minType)
            if debug:
                print(f'solution: {solution}')
            # 從預排清單中移除該次類別
            orderTypes.remove(minType)

        # 排入製造工班
        # 判斷能否排入訂單依據, 總工時
        manufacturingHours = shifts[shift]['計算用工時']
        # 判斷是否超過產能
        flag = False

        # 依據目前類別序
        for index in range(1,len(solution)):
            # 換線時間總計
            shifts[shift]['換線時間'].append(conf[solution[index-1]][solution[index]])
            # 產能 -= 換線
            manufacturingHours -= conf[solution[index-1]][solution[index]]
            # 該類別訂單, 逐筆填入
            while orders[solution[index]]:
                # 取得訂單
                order = orders[solution[index]].pop(0)
                if debug:
                    print(f'order: {order}')
                    print(f'可用工時(扣除前): {manufacturingHours}')
                # 可用產能扣掉該筆訂單生產時間
                manufacturingHours -= order[-2]
                # 新增工時, 所有訂單時間總和
                shifts[shift]['新增工時'] += order[-2]
                # 將訂單加入清單
                shifts[shift]['訂單'].append(order)
                if debug:
                    print(f'可用工時(扣除後): {manufacturingHours}')
                if manufacturingHours < 0:
                    flag = True
                    if debug:
                        print(f'訂單: {shifts[shift]["訂單"]}')
                        print(f'總換線時間: {shifts[shift]["換線時間"]}')
                    break
            if flag:
                break
            del orders[solution[index]]

        if debug:
            print(f'shift: {shifts[shift]}')
        shifts[shift]['計算用工時'] = manufacturingHours

    return shifts, orders

def shiftScheduler(shifts, orders, specialShift):
    # 設定可排工班
    small = ['製造1班', f'製造{specialShift^1}班', f'製造{specialShift}班']
    manufacturing = [f'製造{specialShift}班', f'製造{specialShift^1}班', '製造1班']
    special= [f'製造{specialShift ^ 1}班', f'製造{specialShift}班', '製造1班']


    # optimize 對各類型進行最佳化排程
    # arg1 工班列表
    # arg2 訂單
    # arg3 工班資訊
    shifts, orders['小'] = optimize(small, orders['小'], shifts)
    shifts, orders['製'] = optimize(manufacturing, orders['製'], shifts)
    shifts, orders['特殊'] = optimize(manufacturing, orders['特殊'], shifts)
    for types in ['小', '製', '特殊']:
        orders['其他'].update(orders[types])
    shifts, orders['其他'] = optimize(special, orders['其他'], shifts)
    for shift in small:
        shifts[shift]['本週剩餘工時'] = max(-1 * (shifts[shift]['總工時'] - shifts[shift]['上週剩餘工時'] - shifts[shift]['新增工時'] - \
                                  sum(shifts[shift]['換線時間']) - shifts[shift]['補正工時']), 0)
    total = 0
    count = 0
    for types in orders['其他'].keys():
        count += len(orders['其他'][types])
        for order in orders['其他'][types]:
            total += order[-2]
    shifts['製造1班']['額外加班'] = total
    shifts['製造1班']['備註'] = '' if count == 0 else count
    return shifts

if __name__ == "__main__":
    pass
