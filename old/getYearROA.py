from getData import *



def ROA_save_to_excel(results : list):
    print("\n\n\n")
    print(results)
    print("\n\n\n")
    print(results[0])
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
    wb.save(f"./results/年ROA.xlsx")

def search(beginYear, endYear):
    stock_numbers = read_stock_numbers_from_excel("stocks.xlsx")
    cnt = 1
    login("h1110539@stu.wghs.tp.edu.tw")
    finalResult = [] 
    # finalResult: [every year]
    # in every year: {year : ROE}
    
    for stock in stock_numbers:
        # print("目前進度:", searchYear, searchSeason, cnt)
        cnt += 1
        results = setInformation(beginYear, endYear, stock)
        results = {}
        results = getYearROA(results)
        print(cnt, stock, results)
        finalResult.append(results)
    
    ROA_save_to_excel(finalResult)

search(2004, 2023)
driver.quit()