# -*- coding: UTF-8 -*-
"""
@Project :TFInternShare
@File    :交易集中度.py
@IDE     :Pycharm
@Author  :tutu
@Date    :2023/12/26 11:16
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

# 获取股票交易额数据，以月为一个区间获取
j = 0
for enddate in dateperiod1:
    try:
        allRatioDF1M = pd.read_csv("input/allamountDF" + enddate[0:6] + ".csv")
    except FileNotFoundError:
        print("本地文件不存在，尝试从接口获取数据...")
        num = 1000
        n = int(len(allStock1) / num)
        # 日频数据
        allRatioDF1M = THS_HQ(','.join(allStock1[n * num::]), 'amount', '', enddate[0:6] + '01', enddate[0:6]).data
        for i in range(n):
            df1 = THS_HQ(','.join(allStock1[i * num:(i + 1) * num]), 'amount', '', enddate[0:6]+'01', enddate[0:6]).data
            allRatioDF1M = pd.concat([allRatioDF1M, df1])
        allRatioDF1M.to_csv("input/allamountDF" + enddate[0:6] + ".csv")
    j = j+1

    if j == 0:
        allRatioDF = allRatioDF1M
    else:
        allRatioDF = pd.concat([allRatioDF, allRatioDF1M])

allRatioDF = allRatioDF.drop_duplicates(subset=['thscode', 'time'], keep='first')
allRatioDF = allRatioDF.pivot(index='thscode', columns='time', values='amount')
date_columns = allRatioDF.columns[1:]
selected_columns = [col for col in date_columns if monthdate[0] <= col <= monthdate[1]]
allRatioDF = allRatioDF[['thscode'] + selected_columns]
column_names = allRatioDF.columns

con104list = []
con204list = []

num_date = len(column_names) - 1
for i in range(num_date):
    column_name = column_names[i + 1]
    # 按照某一列的值由大到小进行排序
    df_sorted = allRatioDF.sort_values(by=column_name, ascending=False)
    # 选择前 10% 的数据
    top_10_percent = df_sorted.head(int(0.1 * len(df_sorted)))
    # 计算前 10% 的数据某一列的加总
    sum_top_10_percent = top_10_percent[column_name].sum()
    # 计算整体某一列的加总
    sum_total = allRatioDF[column_name].sum()
    # 计算结果
    result1 = sum_top_10_percent / sum_total
    con104list.append(result1 * 100)

    top_20_percent = df_sorted.head(int(0.2 * len(df_sorted)))
    # 计算前 10% 的数据某一列的加总
    sum_top_20_percent = top_20_percent[column_name].sum()
    # 计算结果
    result2 = sum_top_20_percent / sum_total
    con204list.append(result2 * 100)

plt.figure(figsize=(20, 6))
date = [item[0:4] + item[5:7] + item[8:10] for item in column_names[1::]]
x = date
names = ["前10%成交额占比", "前20%成交额占比"]
color=['#dc965a','#efbf96']
ally = [con104list, con204list]
date_list = []
for item in date:
    date_list.append(item)
    date_list = date_list + [" " for i in range(9)]
date_list = date_list[0:-9]
i = 0
for y in ally:
    f = interp1d(np.arange(len(x)), y, kind='cubic')
    x_smooth = np.linspace(0, len(x) - 1, len(date_list))
    y_smooth = f(x_smooth)
    plt.plot(x_smooth, y_smooth, label=names[i], linewidth=2, color='red')
    i = i + 1
plt.legend(loc='upper center', bbox_to_anchor=(0.5, 1), ncol=4)
num_labels_to_display = 10
step = len(x_smooth) // (num_labels_to_display - 1)
x_ticks_to_display = x_smooth[::step]
date_labels_to_display = date_list[::step]
plt.xticks(x_ticks_to_display, date_labels_to_display, rotation=90)
plt.ylabel('%')
plt.grid(axis='x',linestyle='-')  # 只显示横向网格线
plt.grid(axis='y',linestyle='')  # 不显示竖向网格线
plt.tight_layout()
plt.savefig('output/交易集中度.png')  # 保存为png图片
