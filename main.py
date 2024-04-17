from getData import *


def search(beginYear, endYear):
    stock_numbers = read_stock_numbers_from_excel("stocks.xlsx")
    cnt = 1
    login("h1110539@stu.wghs.tp.edu.tw")
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
        save_to_excel(finalResult[year], year)
    
# fast search
# import sys

# _year = int(sys.argv[1])
# _month = int(sys.argv[2])
# print("search ", _year, _month)
# search(_year, _month)
                
search(2001, 2023)
driver.quit()
