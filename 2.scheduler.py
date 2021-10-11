# -*- coding: UTF-8 -*-

##################################
# AIGO Smart Scheduling
# team CKC
##################################

import logging
from pathlib import Path
from datetime import datetime
from libs.common import 讀取Yaml檔, shiftScheduler
from libs.excelRelative import addColumns, readWorkingHourForm, scheduleOutput
from logging.handlers import TimedRotatingFileHandler

debug = False

class scheduler():
    def __init__(self, src, output):
        self.src = Path(src)
        self.output = Path(output)
        self.orders = None
        self.conf = 讀取Yaml檔('src/config.yml')

    # 附加額外資訊
    def additionalInfo(self, generateFile=False):
        # 製令單加上產能、生產時間、類型
        self.orders = addColumns(self.src, generateFile, self.output)
        logger.info(f'總製令單筆數: {self.orders.shape[0]}, 欄位數: {self.orders.shape[1]}')
        logger.info(f'資料概觀:\n{self.orders.head(5)}')

    # 訂單依週別分群
    def orderSplitByWeek(self):
        logger.info(f'讀取待填工時表單')
        # 取得週列表
        self.info = readWorkingHourForm()
        logger.info(f'讀取完畢')
        # 針對可排訂單進行週別分類
        for i in range(5):
            self.info['週別'][str(i+1)]['訂單'] = list()
        for order in self.orders.values:
            shippingDate = datetime.strptime(order[-4], '%Y/%m/%d')
            # 生產時間：
            if shippingDate < self.info['週別']['1']['開始'] or shippingDate > self.info['週別']['5']['結束'] or order[-3] == None or order[-1] == None:
                continue

            for i in range(5):
                if shippingDate >= self.info['週別'][str(i+1)]['開始'] and shippingDate <= self.info['週別'][str(i+1)]['結束']:
                    self.info['週別'][str(i+1)]['訂單'].append(order)
                    break
    # 訂單排程
    def weeklyClass(self):

        # self.final 記錄排程最終資訊
        self.final = dict()
        for i in range(1, 6):
            self.final[str(i)] = {f'製造{j}班': 0 for j in range(1, self.conf['製造工班數']+1)}

        # 考量限制，依週別排程
        for week in self.info['週別'].keys():
            workingHour = dict()
            # workingHour 
            for shift in range(1, self.conf['製造工班數']+1):

                # workingHour為當週該製造工班資訊
                # 新增工時: 所有訂單生產時間總和 ok!
                # 每日工時: 讀取待填表單, 每天上班工時, 已加上加班工時並扣除補正工時與清潔時間 ok!
                # flag: 工班屬性
                # 前筆訂單類型: 第一週-讀取待填表單
                #             第二至五週-前一週最後一筆訂單類型
                # 訂單: 排程後訂單 (用以排入訂單)
                # 總工時: 稼動總工時 = 正常工時+上週剩餘工時 = sum(每日工時) + 補正工時 ok!
                # 換線時間: 換線所需總時間 ok!
                # 上週剩餘工時: 第一週-0, 如有前一週剩餘工時, 請填表人將該值加至補正工時
                #             第二至五週-前一週之本週剩餘工時
                # 本週剩餘工時: 總工時 - 上週剩餘工時 - 新增工時 - 換線時間 - 補正工時
                # 補正工時: 讀取待填表單
                # 加班工時: 讀取待填表單, 輸出表單中紀錄用
                total = sum(self.info[f'製造{shift}班']['工時'][week].values())
                if self.info[f'製造{shift}班']['補正工時'][week]:
                    total += self.info[f'製造{shift}班']['補正工時'][week]
                workingHour[f'製造{shift}班'] = {'新增工時': 0,
                                              '每日工時': self.info[f'製造{shift}班']['工時'][week],
                                              'flag': self.conf['flag'][shift-1],
                                              '前筆訂單類型': self.info[f'製造{shift}班']['前筆訂單類型'],
                                              '訂單': list(),
                                              '總工時': total,
                                              '計算用工時': total,
                                              '換線時間': list(),
                                              '上週剩餘工時': 0,
                                              '本週剩餘工時': 0,
                                              '補正工時': self.info[f'製造{shift}班']['補正工時'][week] if self.info[f'製造{shift}班']['補正工時'][week]  else 0,
                                              '加班工時': self.info[f'製造{shift}班']['加班工時'][week],
                                              '額外加班': 0,
                                              '備註': ''
                                              }
            # 非第一週的設定:
            # 上週剩餘工時需讀取前一週之本週剩餘工時
            # 前筆訂單類型需讀取前一週之最後一筆訂單類型
            if week != '1':
                workingHour[f'製造{shift}班']['上週剩餘工時'] = self.final[str(int(week) - 1)][f'製造{shift}班']['本週剩餘工時']
                workingHour[f'製造{shift}班']['前筆訂單類型'] = self.final[str(int(week) - 1)][f'製造{shift}班']['訂單'][-1][11]
            if debug:
                print(f"上週剩餘工時: {workingHour[f'製造{shift}班']['上週剩餘工時']}")
                print(f"前筆訂單類型: {workingHour[f'製造{shift}班']['前筆訂單類型']}")

            logger.info(
                f'''開始進行 {self.info["週別"][week]["開始"].strftime('%Y/%m/%d')} 至 {self.info["週別"][week]["結束"].strftime('%Y/%m/%d')}''')
            # self.classification: 訂單依類別分類，根據廠商說明分成
            # 小型: 小型類別主要由製造一班生產
            # 製: 製1~8由二、三班輪流，以本排程將預設以週次決定當週生產之班別，第一週由製造二班、第二週由製造三班
            # 特殊: 廠商說明某些類型將安排於製1~8前後生產
            # 其他: 其他類型工單

            self.info['週別'][week]['訂單'] = self.classification(self.info['週別'][week]['訂單'])

            overall = 0
            for i in self.info['週別'][week]['訂單'].keys():
                total = 0
                for type in self.info['週別'][week]['訂單'][i]:
                    total += len(self.info['週別'][week]['訂單'][i][type])
                logger.info(f"類別: {i} - 共 {len(self.info['週別'][week]['訂單'][i])} 種, {total} 筆")
                overall += total
            logger.info(f"該週總訂單數: {overall}")
            logger.info(f'''完成 {self.info["週別"][week]["開始"].strftime('%Y/%m/%d')} 至 {self.info["週別"][week]["結束"].strftime('%Y/%m/%d')} 訂單分類''')

            # 生產製1~8工班
            specialShift = self.info['週別'][week]['製造']
            logger.info(
                f'''本週生產製 1~8之製造工班為: {self.info['週別'][week]['製造']} 班''')

            # shiftScheduler: 進行排程
            # arg1: workingHour 為當週工班資訊
            # arg2: 分群後訂單
            # arg3: 生產製1~8工班
            workingHour = shiftScheduler(workingHour, self.info['週別'][week]['訂單'], specialShift)
            # 排程資訊更新
            for shift in range(1, self.conf['製造工班數']+1):
                self.final[week][f'製造{shift}班'] = workingHour[f'製造{shift}班']
            # print(self.final['1'][f'製造1班'])
            if debug:
                print(f"本週剩餘工時: {workingHour[f'製造{shift}班']['本週剩餘工時']}")
        # 輸出排程表單
        scheduleOutput(self.final)

    # 逐週將訂單依類別分群
    def classification(self, orders):
        classifiedOrders = dict()
        for order in orders:
            if order[-1] not in classifiedOrders.keys():
                classifiedOrders[order[-1]] = [order]
            else:
                classifiedOrders[order[-1]].append(order)
        return self.conditionalClassification(classifiedOrders)

    # 將分群後訂單進一步依排程方式分群
    def conditionalClassification(self, ClassOfOrders):
        newCategories = {'小': dict(), '製': dict(), '特殊': dict(), '其他': dict()}

        # 特殊類型
        condition = []
        for key, value in ClassOfOrders.items():
            # 條件一、特殊類型
            if key in condition:
                newCategories['特殊'][key] = value
            # 條件二、小型
            elif '小' in key: 
                newCategories['小'][key] = value
            # 條件三、製1~8
            elif '製' in key:
                newCategories['製'][key] = value
            else:
                newCategories['其他'][key] = value
        return newCategories

if __name__ == "__main__":
    conf = 讀取Yaml檔('src/config.yml')

    # logger handle
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s - %(message)s'))
    logger.addHandler(console_handler)

    Path(conf['日誌路徑']).parent.mkdir(parents=True, exist_ok=True)
    file_handler = TimedRotatingFileHandler(conf['日誌路徑'], when="midnight", interval=1, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s - %(message)s'))
    logger.addHandler(file_handler)

    # 設定讀取製令單
    play = scheduler(src='src/1-8月製令單_new.xlsx', output='output/')
    play.additionalInfo(generateFile=True)
    play.orderSplitByWeek()
    play.weeklyClass()
