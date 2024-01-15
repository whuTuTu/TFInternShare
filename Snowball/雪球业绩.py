# 导入所需的库
import pandas as pd

# 定义雪球合约的参数，函数外面定义的变量不用做特殊说明，就是全局变量
contract_term = 24  # 合约期限，单位为月
strike_out = 1  # 敲出水平，相对于期初价格的比例
strike_in = 0.7  # 敲入水平，相对于期初价格的比例
coupon_rate = 0.15  # 票息收益率，年化
leverage = 1  # 杠杆倍数
start_date = '2021-1-4'  # 回测起点
start_price = 6798.7644  # 回测起点的指数点数
end_date = '2023-08-30'  # 回测终点

# 读取关联标的的历史收盘价数据，假设为中证1000指数
df = pd.read_excel(r'input/雪球输入.xlsx', sheet_name='1000')  # 从csv文件读取数据，您需要根据您的数据源进行修改
df['date'] = pd.to_datetime(df['date'])  # 将日期列转换为datetime格式
df['close'] = df['close'].astype(float)  # 将收盘价列转换为浮点数格式
df.set_index('date', inplace=True)  # 将日期列设为索引


# 定义一个函数，根据给定的起始日期，模拟一个雪球合约的收益情况
def snowball_return(start_date, initial_price):
    # 计算合约到期日期，假设每月有20个交易日
    end_date = start_date + pd.DateOffset(months=contract_term)  # DateOffset是一个日期偏移函数？？？

    # 截取合约存续期内的标的收盘价数据
    sub_df = df.loc[start_date:end_date]
    sub_df = sub_df['close']
    price0 = sub_df[0]

    # 计算期初价格和敲出/敲入价格
    out_price = price0 * strike_out
    in_price = price0 * strike_in

    # 初始化一些变量，用于记录合约的状态和收益
    knocked_out = False  # 是否发生敲出事件
    knocked_in = False  # 是否发生敲入事件
    expired = False  # 事件结束标志，表示发生敲出事件或者期满
    expired_date = None  # 敲出日期
    in_date = None  # 敲入日期
    out_date = None
    holding_days = 0  # 合约持有天数

    # 遍历每个交易日，判断是否发生敲出或敲入事件，并更新合约状态和收益
    for date, close in sub_df.items():
        holding_days += 1  # 持有天数加一
        if holding_days < 40:
            continue  # 锁定期40天

        if (holding_days % 20 == 0) and (close >= out_price):  # 如果是月度观察日且标的价格高于或等于敲出价格，视为发生敲出事件
            knocked_out = True
            expired = True
            out_date = date
            expired_date = out_date
            break

        elif not knocked_in:
            if close <= in_price:  # 如果标的价格低于或等于敲入价格，视为发生敲入事件
                in_date = date
                knocked_in = True

        # 如果到达合约到期日，结束循环
        elif date >= end_date:
            expired = True
            expired_date = end_date
            break

    if expired:
        if knocked_out:  # 如果发生了敲出事件，提前终止合约，获得年化票息收益
            price = initial_price * coupon_rate * holding_days / 240 + initial_price  # 假设一年有240个交易日
        elif not knocked_in:  # 如果没有发生敲入事件，到期终止合约，获得年化票息收益
            price = initial_price * coupon_rate * contract_term / 12 + initial_price
        else:  # 如果发生了敲入事件，到期终止合约，承担标的下跌造成的损失
            if close >= price0:
                price = initial_price
            else:
                price = initial_price * (close / price0 - 1) + initial_price
        next_start_date = expired_date
    elif knocked_in:  # expired=0的情况，可能延后两年终止日在周末
        if close >= price0:
            price = initial_price
        else:
            price = initial_price * (close / price0 - 1) + initial_price
        next_start_date = end_date
        expired_date = start_date + pd.DateOffset(months=contract_term)
    else:
        price = initial_price * coupon_rate * contract_term / 12 + initial_price
        next_start_date = end_date
        expired_date = start_date + pd.DateOffset(months=contract_term)
    return [start_date, expired_date, holding_days, out_date, in_date, price, next_start_date]


# 定义一个空的列表，用于存储每个合约的状态和收益
results = []

def xueqiu(date, initial_price):
    result = snowball_return(date, initial_price)  # list
    date = result[-1]  # 滚动赋值
    initial_price = result[-2]
    results.append(result)

    if date < pd.to_datetime(end_date):  # 中止日期
        xueqiu(date, initial_price)
    return results


results = xueqiu(pd.to_datetime(start_date), start_price)  # 替换初始数据
results_df = pd.DataFrame(results,
                          columns=['start_date', 'expired_date', 'holding_days', 'out_date', 'in_date', 'price',
                                   'next_start_date'])
results_df.to_excel(r'output/雪球业绩输出.xlsx')
