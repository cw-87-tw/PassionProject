# import tkinter as tk
# from tkinter import ttk
# from UI import *
from getData import *
try:
    import openpyxl
except:
    import os
    os.system("pip install openpyxl")
    import openpyxl


def read_stock_numbers_from_excel(filename):
    stock_numbers = []
    readwb = openpyxl.load_workbook(filename)
    readws = readwb.active
    for row in readws.iter_rows(values_only=True):
        stock_numbers.append(row[2])
    # print("stock numbers:", *stock_numbers)
    print("got stock numbers", len(stock_numbers))
    return stock_numbers

wb = openpyxl.Workbook()
ws = wb.active
l = []
def save_to_excel(results : dict, name):
    if len(l) == 0: ws.append(["公司編號"] + list(results.keys()))  # 添加標題行
    l.append([name] + list(results.values()))
    name = len(l)
    ws.append([name] + list(results.values()))
    # global year
    # global season
    wb.save(f"{year}年第{season}季.xlsx")

def search(searchYear, searchSeason):
    # changeButton(text="搜尋中...請稍候", state = "disabled")
    # root.update()
    l.clear()
    global wb
    wb = openpyxl.Workbook()
    global ws
    ws = wb.active
    stock_numbers = read_stock_numbers_from_excel("stocks1.xlsx")
    cnt = 1
    login("h1110539@stu.wghs.tp.edu.tw")
    for stock in stock_numbers:
        setInformation(searchYear, searchSeason, stock)
        results = dict()
        results.update(getPrice())
        results.update(getIncome())
        results.update(getCash())
        results.update(getDebt())
        results.update(getRoeRoa())
        results.update(getAssets())
        results.update(getEps())
        save_to_excel(results, stock)
        # show_result({"目前進度" : cnt})
        print("目前進度:", cnt)
        cnt += 1
    # changeButton(text="查詢", state = "normal")
    # show_result(result={"狀態" : "成功"})
    
for y in range(2014, 2024):
    for s in range(1, 5):
        print("search", y, s)
        search(y, s)

# init()
# root.mainloop() # start program
driver.quit()
