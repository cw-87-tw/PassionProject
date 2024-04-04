from getData import *

def getYearROE(results):
    print("Get EPS:", target)
    setRange(f"https://statementdog.com/analysis/{target}/roe-roa")
    clickButton("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[3]/ul/li[2]")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        index = year - beginYear + 1
        results[year] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
        try: 
            if float(results[year]) < 5: results[year] = 0
        except: results[year] = 0
    return results

def ROE_save_to_excel(results : list):
    wb = openpyxl.Workbook()
    ws = wb.active
    cnt = 0
    ws.append(["公司編號"] + list(results[0].keys()))  # 添加標題行
    for i in results:
        cnt += 1
        ws.append([cnt] + list(i.values()))
    # global year
    # global season
    # print("save year", year, "season", season)
    wb.save(f"./results/年ROE.xlsx")

def search(beginYear, endYear):
    stock_numbers = read_stock_numbers_from_excel("stocks.xlsx")
    cnt = 1
    login("h1110539@stu.wghs.tp.edu.tw")
    finalResult = [[] for _ in range(beginYear, endYear + 1)] 
    # finalResult: [every year]
    # in every year: {year : ROE}
    
    for stock in stock_numbers:
        # print("目前進度:", searchYear, searchSeason, cnt)
        cnt += 1
        results = setInformation(beginYear, endYear, stock)
        results = {}
        results = getYearROE(results)
        print(cnt, stock, results)
        for year in range(beginYear, endYear + 1):
            finalResult.append(results)
    
    ROE_save_to_excel(results)

search(2004, 2023)
driver.quit()