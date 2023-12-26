from iFinDPy import *
# from threading import Thread,Lock,Semaphore
import matplotlib.pyplot as plt
import pandas as pd
# from openpyxl import Workbook
import sys

print(sys.executable)

# 实例化
# wb = Workbook() # openpyxl库，用于创建excel，进行编辑
# # 激活 worksheet
# ws = wb.active

# sem = Semaphore(5)  # 创建名为sem的信号量实体，信号量是用于控制并发访问资源的同步性机制，这里表示最多有5个线程可以同时获取该信号。
# dllock = Lock()  # 创建锁，防止多个线程同时修改数据而引发冲突的机制。
thsLogin = THS_iFinDLogin("tfzq1556", "752862")
df = THS_HQ('399300.SZ,000905.SH,000852.SH,700008.TI', 'turnoverRatio', '', '2023-05-09', '2023-06-09')
df = df.data
df = df.pivot_table(index="time", columns="thscode", values="turnoverRatio")
df.to_csv('output/zs.csv')
df = pd.read_csv('zs.csv')
df = df.rename(columns={'000852.SH': '中证1000', '000905.SH': '中证500', '399300.SZ': '沪深300', '700008.TI': '全市场'})
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams['font.size'] = 15
fig = plt.figure(figsize=(18, 6))
ax = fig.add_subplot()
df.plot(x='time', ax=ax, color=['#FFC175', '#000000', '#FF0000', '#D5995D'])
plt.legend(loc='upper center', ncol=4, frameon=False)
plt.show()
