# -*- coding: UTF-8 -*-
"""
@Project :TFInternShare
@File    :指数截面波动率.py
@IDE     :Pycharm
@Author  :tutu
@Date    :2024/1/3 10:24
为月报准备，以月为单位储存数据。
"""
# 载入包
from iFinDPy import *  # 同花顺API接口
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import numpy as np
import warnings
import configparser
warnings.filterwarnings(action='ignore')
warnings.filterwarnings(action='ignore')  # 导入warnings模块，并指定忽略代码运行中的警告信息
plt.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文显示乱码的问题
plt.rcParams['axes.unicode_minus'] = False
config = configparser.ConfigParser()
config.read("config.ini", encoding="utf-8")
# 连接API接口
apikey = [config.get("apikey", "ID3"), config.get("apikey", "password3")]
thsLogin = THS_iFinDLogin(apikey[0], apikey[1])

# ---------------------------------------------更改两个变量----------------------------------------------
# 更改时间区间(周报为一个月，月报为一年)
monthdate = ["20231129", "20231229"]

# ---------------------------------------------下面为代码-----------------------------------------------
# 获取股票代码
allStock = THS_DR('p03291', 'date=' + monthdate[1] + ';blockname=001005010;iv_type=allcontract', 'p03291_f002:Y',
                  'format:dataframe').data.iloc[:, [0]].values.tolist()
allStock1 = [item for sublist in allStock for item in sublist]
dateperiod = [year + month + '31' for year in [monthdate[0][0:4], monthdate[1][0:4]] for month in
              ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']]
monthdate1 = [monthdate[0][0:6]+'31',monthdate[1][0:6]+'31']
dateperiod1 = []
for i in dateperiod:
    if monthdate1[1] >= i >= monthdate1[0]:
        dateperiod1.append(i)
    else:
        pass
dateperiod1 = list(set(dateperiod1))  # 利用集合去重
print(dateperiod1)

# 获取股票涨跌幅数据，以月为一个区间获取
j = 0
for enddate in dateperiod1:
    try:
        allRatioDF1M = pd.read_csv("input/allratioDF" + enddate[0:6] + ".csv")
    except FileNotFoundError:
        print("本地文件不存在，尝试从接口获取数据...")
        num = 1000
        n = int(len(allStock1) / num)
        # 日频数据
        allRatioDF1M = THS_HQ(','.join(allStock1[n * num::]), 'changeRatio', '', enddate[0:6] + '01', enddate[0:6]).data
        for i in range(n):
            df1 = THS_HQ(','.join(allStock1[i * num:(i + 1) * num]), 'changeRatio', '', enddate[0:6]+'01', enddate[0:6]).data
            allRatioDF1M = pd.concat([allRatioDF1M, df1])
        allRatioDF1M.to_csv("input/allratioDF" + enddate[0:6] + ".csv")
    j = j+1

    if j == 0:
        allRatioDF = allRatioDF1M
    else:
        allRatioDF = pd.concat([allRatioDF, allRatioDF1M])

ThreeIndex = THS_HQ('000300.SH,000852.SH,000905.SH', 'changeRatio', '', momthdate[0], momthdate[1]).data
