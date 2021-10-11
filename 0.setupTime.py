# -*- coding: UTF-8 -*-

##################################
# AIGO Smart Scheduling
# team CKC
##################################
import yaml
import openpyxl
import pandas as pd

def main():
    wb = openpyxl.load_workbook('src/換線時間對照表.xlsx')
    types = []
    df = pd.DataFrame(wb['對照表'].values).iloc[2:, 2:]
    df.columns = types
    setupTime = dict()
    for col in types:
        setupTime[col] = dict()
        for col2 in types:
            setupTime[col][col2] = 0

    for i in range(df.shape[0]):
        for index, value in enumerate(df[types[i]].values):
            if value and i != index:
                setupTime[types[i]][types[index]] = value
                setupTime[types[index]][types[i]] = value
    with open('src/換線時間.yml', 'w') as outfile:
        yaml.dump(setupTime, outfile, allow_unicode=True, default_flow_style=False)

if __name__=="__main__":
    main()