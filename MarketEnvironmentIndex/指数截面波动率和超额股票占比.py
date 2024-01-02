# -*- coding: UTF-8 -*-
"""
@Project :TFInternShare
@File    :指数截面波动率和超额股票占比.py
@IDE     :Pycharm
@Author  :tutu
@Date    :2023/12/25 15:16
不排除首日上市的股票
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
apikey = [config.get("apikey", "ID2"),config.get("apikey", "password2")]
thsLogin = THS_iFinDLogin(apikey[0], apikey[1])

# ---------------------------------------------更改两个变量----------------------------------------------
# 更改时间区间(周报为一个月，月报为一年)
momthdate = ["20221229", "20231229"]

# 选择频率，True为日频，False为周频
Flag = False

# ---------------------------------------------下面为代码----------------------------------------------
# 获取股票代码
allStock = THS_DR('p03291', 'date=' + momthdate[1] + ';blockname=001005010;iv_type=allcontract', 'p03291_f002:Y',
                  'format:dataframe').data.iloc[:, [0]].values.tolist()
allStock1 = [item for sublist in allStock for item in sublist]

# 获取股票收益率数据
try:
    allRatioDF1 = pd.read_csv("input/allRatioDF" + momthdate[0] + '-' + momthdate[1] + ".csv")
except FileNotFoundError:
    print("本地文件不存在，尝试从接口获取数据...")
    n = int(len(allStock1) / 1000)
    if Flag:
        # 日频数据
        allRatioDF = THS_HQ(allStock1[n * 1000::], 'changeRatio', '', momthdate[0], momthdate[1]).data
        for i in range(n):
            df1 = THS_HQ(allStock1[i * 1000:(i + 1) * 1000], 'changeRatio', '', momthdate[0], momthdate[1]).data
            allRatioDF = pd.concat([allRatioDF, df1])
    else:
        # 周频数据
        allRatioDF = THS_HQ(allStock1[n * 1000::], 'changeRatio','Interval:W', momthdate[0], momthdate[1]).data
        for i in range(n):
            df1 = THS_HQ(allStock1[i * 1000:(i + 1) * 1000], 'changeRatio','Interval:W', momthdate[0], momthdate[1]).data
            allRatioDF = pd.concat([allRatioDF, df1])
    allRatioDF1 = allRatioDF.pivot(index='thscode', columns='time', values='changeRatio')
    allRatioDF1.to_csv("input/allRatioDF" + momthdate[0] + '-' + momthdate[1] + ".csv")
# 读取数据
allRatioDF1 = pd.read_csv("input/allRatioDF" + momthdate[0] + '-' + momthdate[1] + ".csv")
column_names = allRatioDF1.columns

# 指数的涨跌幅
if Flag:
    ThreeIndex = THS_HQ('000300.SH,000852.SH,000905.SH', 'changeRatio', '', momthdate[0], momthdate[1]).data
else:
    ThreeIndex = THS_HQ('000300.SH,000852.SH,000905.SH', 'changeRatio', 'Interval:W', momthdate[0], momthdate[1]).data
ThreeIndex1 = ThreeIndex.pivot(index='thscode', columns='time', values='changeRatio')

# 时间序列
date = [item[0:4] + item[5:7] + item[8:10] for item in column_names[1::]]
date_len = len(date)

# 计算波动率
std5000 = []
std300 = []
std500 = []
std1000 = []

# 计算超额股票占比
ratio300 = []
ratio500 = []
ratio1000 = []

# 计算中位数超越幅度
media300 = []
media500 = []
media1000 = []

for i in range(date_len):
    # 全市场
    DF5000 = allRatioDF1[~allRatioDF1['thscode'].isin(column_names)].dropna(subset=[column_names[i + 1]])[
        [column_names[i + 1]]]  # 剔除掉新股
    std5000 = std5000 + DF5000.std().tolist()

    # 沪深300
    index300 = THS_DR('p03473', 'iv_date=' + date[i] + ';iv_zsdm=000300.SH', 'p03473_f002:Y',
                      'format:dataframe').data.iloc[:, [0]].values.tolist()
    index300 = [item for sublist in index300 for item in sublist]
    DF300 = allRatioDF1[allRatioDF1['thscode'].isin(index300)][[column_names[i + 1]]]
    std300 = std300 + DF300.std().tolist()
    count_above_threshold = len(DF300[DF300[column_names[i + 1]] > ThreeIndex1.loc['000300.SH', column_names[i + 1]]])
    ratio300.append(count_above_threshold / 300 * 100)
    median_value = DF300[column_names[i + 1]].median()
    media300.append(median_value - ThreeIndex1.loc['000300.SH', column_names[i + 1]])

    # 中证500
    index500 = THS_DR('p03473', 'iv_date=' + date[i] + ';iv_zsdm=000905.SH', 'p03473_f002:Y',
                      'format:dataframe').data.iloc[:, [0]].values.tolist()
    index500 = [item for sublist in index500 for item in sublist]
    DF500 = allRatioDF1[allRatioDF1['thscode'].isin(index500)][[column_names[i + 1]]]
    std500 = std500 + DF500.std().tolist()
    count_above_threshold = len(DF500[DF500[column_names[i + 1]] > ThreeIndex1.loc['000852.SH', column_names[i + 1]]])
    ratio500.append(count_above_threshold / 500 * 100)
    median_value = DF500[column_names[i + 1]].median()
    media500.append(median_value - ThreeIndex1.loc['000852.SH', column_names[i + 1]])

    # 中证1000
    index1000 = THS_DR('p03473', 'iv_date=' + date[i] + ';iv_zsdm=000852.SH', 'p03473_f002:Y',
                       'format:dataframe').data.iloc[:, [0]].values.tolist()
    index1000 = [item for sublist in index1000 for item in sublist]
    DF1000 = allRatioDF1[allRatioDF1['thscode'].isin(index1000)][[column_names[i + 1]]]
    std1000 = std1000 + DF1000.std().tolist()
    count_above_threshold = len(DF1000[DF1000[column_names[i + 1]] > ThreeIndex1.loc['000905.SH', column_names[i + 1]]])
    ratio1000.append(count_above_threshold / 1000 * 100)
    median_value = DF1000[column_names[i + 1]].median()
    media1000.append(median_value - ThreeIndex1.loc['000905.SH', column_names[i + 1]])

# ########################################## 绘图指数截面波动率图 ##########################################
plt.figure(figsize=(20, 6))
date = [item[0:4] + item[5:7] + item[8:10] for item in column_names[1::]]
x = date
names = ["全市场", "沪深300", "中证500", "中证1000"]
ally = [std5000, std300, std500, std1000]
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
    plt.plot(x_smooth, y_smooth, label=names[i], linewidth=2)
    i = i + 1
plt.legend(loc='upper center', bbox_to_anchor=(0.5, 1), ncol=4)
plt.title('指数截面波动率')

num_labels_to_display = len(date)
step = len(x_smooth) // (num_labels_to_display - 1)
x_ticks_to_display = x_smooth[::step]
date_labels_to_display = date_list[::step]
plt.xticks(x_ticks_to_display, date_labels_to_display, rotation=0)
plt.ylabel('%')
plt.grid(True)
plt.tight_layout()
plt.savefig('output/指数截面波动率.png')  # 保存为png图片

# ########################################## 超额股票占比 ##########################################
types = ["沪深300","中证500","中证1000"]
y_ratio = [ratio300,ratio500,ratio1000]
y_median = [media300,media500,media1000]

date = [item[0:4] + item[5:7] + item[8:10] for item in column_names[1::]]
date_list = []
for item in date:
    date_list.append(item)
    date_list = date_list + [" " for i in range(9)]
date_list = date_list[0:-9]
for i in range(3):
    x = date
    y = y_ratio[i]
    f = interp1d(np.arange(len(x)), y, kind='cubic')
    x_smooth = np.linspace(0, len(x) - 1, len(date_list))
    y_smooth = f(x_smooth)
    fig, ax1 = plt.subplots(figsize=(15, 6))
    ax1.fill_between(x_smooth, y_smooth, color='#F5D0B5', alpha=1, label='超额股票占比')
    ax1.set_ylabel('超额股票占比%')
    ax1.tick_params(axis='y')
    num_labels_to_display = len(date)
    step = len(x_smooth) // (num_labels_to_display - 1)
    x_ticks_to_display = x_smooth[::step]
    date_labels_to_display = date_list[::step]
    plt.xticks(x_ticks_to_display, date_labels_to_display, rotation=45)
    ax2 = ax1.twinx()
    ax2.bar(date, y_median[i], alpha=0.5, color='#DD9558', label='中位数超越幅度')
    ax2.set_ylabel('中位数超越幅度')
    ax2.tick_params(axis='y')
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines = lines1 + lines2
    labels = labels1 + labels2
    plt.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, 1), ncol=2)
    plt.title('超额股票占比（'+types[i]+'）')
    plt.savefig('output/超额股票占比（'+types[i]+'）.png')  # 保存为png图片