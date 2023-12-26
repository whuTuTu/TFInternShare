# 导入所需的库
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 定义雪球合约的参数
contract_term = 24  # 合约期限，单位为月
strike_out = 1  # 敲出水平，相对于期初价格的比例
strike_in = 0.8  # 敲入水平，相对于期初价格的比例
coupon_rate = 0.15  # 票息收益率，年化
leverage = 1  # 杠杆倍数

# 读取关联标的的历史收盘价数据，假设为中证500指数
df = pd.read_excel(r'input/雪球输入.xlsx',sheet_name='500')  # 从csv文件读取数据，您需要根据您的数据源进行修改
df['date'] = pd.to_datetime(df['date'])  # 将日期列转换为datetime格式

df['close'] = df['close'].astype(float)  # 将收盘价列转换为浮点数格式
df.set_index('date', inplace=True)# 将日期列设为索引

# 定义一个函数，根据给定的起始日期，模拟一个雪球合约的收益情况
def snowball_return(start_date):
    # 计算合约到期日期，假设每月有20个交易日
    end_date = start_date + pd.DateOffset(months=contract_term) ##DateOffset可以给定日期偏移量

    # 截取合约存续期内的标的收盘价数据
    sub_df = df.loc[start_date:end_date]
    sub_df = sub_df['close']

    # 计算期初价格和敲出/敲入价格
    initial_price = sub_df[0]
    out_price = initial_price * strike_out
    in_price = initial_price * strike_in

    # 初始化一些变量，用于记录合约的状态和收益
    knocked_out = False  # 是否发生敲出事件
    knocked_in = False  # 是否发生敲入事件
    out_date = None  # 敲出日期
    in_date = None  # 敲入日期
    holding_days = 0  # 合约持有天数


    # 遍历每个交易日，判断是否发生敲出或敲入事件，并更新合约状态和收益
    for date, close in sub_df.items():
        holding_days += 1  # 持有天数加一
        if holding_days < 40:
            continue

        if (holding_days % 20 == 0) and (close >= out_price):  # 如果是月度观察日且标的价格高于或等于敲出价格，视为发生敲出事件
            knocked_out = True  #敲出观察日是月度
            out_date = date
            break

        # 如果没有发生敲出事件，判断是否发生敲入事件
        if not knocked_in:
            if close <= in_price:  # 如果标的价格低于或等于敲入价格，视为发生敲入事件
                knocked_in = True #敲入观察日是日度
                in_date = date

        # 如果到达合约到期日，结束循环
        if date == end_date:
            break

    # 根据合约状态和收益，计算最终收益率
    if knocked_out:  # 如果发生了敲出事件，提前终止合约，获得年化票息收益
        return_rate = coupon_rate * holding_days / 240  # 假设一年有240个交易日
    elif not knocked_in:  # 如果没有发生敲入事件，到期终止合约，获得年化票息收益
        if holding_days > 480:
            return_rate = coupon_rate
        else: return_rate = None
    else:  # 如果发生了敲入事件，到期终止合约，承担标的下跌造成的损失
        if holding_days > 480:
            if close >= initial_price:
                return_rate = 0
            else: return_rate = (close - initial_price) / initial_price
        else:
            return_rate = None

    # 返回合约的状态和收益
    return knocked_out, knocked_in, out_date, in_date, holding_days, return_rate


# 定义一个空的列表，用于存储每个合约的状态和收益
results = []

# 遍历每个交易日，假设每天买入一个雪球合约，并调用上面定义的函数模拟其收益情况
for date in df.index:
    result = snowball_return(date)  # 调用函数，得到一个元组
    results.append(result)  # 将元组添加到列表中

# 将列表转换为数据框，并命名列名

results_df = pd.DataFrame(results,
                          columns=['knocked_out', 'knocked_in', 'out_date', 'in_date', 'holding_days', 'return_rate'])
results_df.to_excel(r'output/雪球回测输出.xlsx')
# 筛选出已经到期的合约，并计算一些统计指标
win_rate = results_df[results_df['return_rate'] > 0].shape[0] / results_df.shape[0]  # 胜率，即正收益的合约占比
mean_return = results_df['return_rate'].mean()  # 平均收益率
max_loss = results_df['return_rate'].min()  # 最大亏损率

# 打印统计指标
print(f'胜率：{win_rate:.2%}')
print(f'平均收益率：{mean_return:.2%}')
print(f'最大亏损率：{max_loss:.2%}')

# 绘制收益率分布的直方图
plt.hist(results_df['return_rate'], bins=20)
plt.xlabel('Return Rate')
plt.ylabel('Frequency')
plt.title('Snowball Return Distribution')
plt.show()
