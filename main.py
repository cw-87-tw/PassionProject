from getData import *

# the following can be modified by users
# stocksSheet = input("查詢的股票編號表格:")
# saveDir = input("存results的資料夾:")
beginYear = 2001
endYear = 2023
account = "h1110539@stu.wghs.tp.edu.tw"
password = "passion"
# end
login("h1110539@stu.wghs.tp.edu.tw")

def search(beginYear, endYear, dir, account, password):
    stock_numbers = read_stock_numbers_from_excel("stocks.xlsx")
    cnt = 1
    login(account, password)
    finalResult = {}
    for year in range(beginYear, endYear + 1):
        finalResult[year] = list()
    
    for stock in stock_numbers:
        # print("目前進度:", searchYear, searchSeason, cnt)
        cnt += 1 
        results = setInformation(beginYear, endYear, stock)
        results = getPrice(results)
        results = getIncome(results)
        results = getCash(results)
        results = getRoeRoa(results)
        results = getDebt(results)
        results = getAssets(results)
        results = getEps(results)
        results = getOther(results)
        print(stock, results)
        for year in range(beginYear, endYear + 1):
            finalResult[year].append(results[year])
    
    for year in range(beginYear, endYear + 1):
        save_to_excel(finalResult[year], year, dir)
    
def run():
    search(beginYear, endYear, saveDir, account, password)


# driver.quit()
