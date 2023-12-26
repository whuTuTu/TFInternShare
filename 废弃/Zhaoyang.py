import pymysql
import pandas as pd
config = {
    'host': '106.75.45.237',
    'port': 50128,
    'user': 'simu4_tfzqzj',
    'password': 'suO9A2nfTBXS2NRV',
    'db': 'CUS_FUND_DB',
    'charset': 'utf8',
    'cursorclass': pymysql.cursors.DictCursor
 }
conn = pymysql.connect(**config)
cursor = conn.cursor()
cursor.execute("select fund_name, statistic_date, week_return, month_return, m1_return, m3_return, m6_return, year_return, total_return_a, total_max_retracement, total_sharp from t_fund_weekly_performance where statistic_date = '2023-06-02' ")
results = cursor.fetchall()
data = pd.DataFrame(list(results))
cursor.execute("select fund_name, statistic_date, total_win_rate, total_loss_to_profit, total_upmonth, total_downmonth from t_fund_monthly_risk where statistic_date = '2023-05-31' ")
results = cursor.fetchall()
data1 = pd.DataFrame(list(results))
data = pd.merge(data, data1, on='fund_name')
cursor.execute("select fund_name, total_win_rate, total_loss_to_profit, total_upweek, total_downweek from t_fund_weekly_risk where statistic_date = '2023-06-02' ")
results = cursor.fetchall()
data1 = pd.DataFrame(list(results))
month_data = pd.merge(data, data1, on='fund_name')
month_data.to_excel(r'C:\Users\86139\Desktop\risk.xlsx')