import twstock # pip install twstock, lxml
import pandas as pd # pip install pandas
target_stock = '0050'
stock = twstock.Stock(target_stock)
target_price = stock.fetch_from(2023, 5)

print(target_price)

name_attribute = ['Date', 'Capacity', 'Turnover', 'Open', 'High', 'Low', 'Close', 'Change', 'Transcation']
# 日期 總成交股數 總成交金額(Volume) 開 高 低 收 漲跌幅 成交量

df = pd.DataFrame(columns = name_attribute, data = target_price)

print(df)