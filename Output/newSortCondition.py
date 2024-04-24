import openpyxl

beginYear = 2013
# beginYear = int(input("beginYear:"))
endYear = 2022
# endYear = int(input("endYear:"))
factors = [("營業利益率", 1)] # 每個 factor: (name, 要算幾份)
numbers = 5


def getData() -> dict:
    """
    the function can get data from the results crawled on the Net
    the workbook must be located at the desinated directory(./results/)
    it would get all the data in the following format
    results = { 
        beginYear: {
            1: {
                factor1: 123,
                factor2: "Error"...
            },
            2: {
            
            }...
        }...
    }
    """
    results = dict()
    for year in range(beginYear, endYear + 1):
        results[year] = dict()
        readwb = openpyxl.load_workbook(f"./results/{year}年.xlsx")
        readws = readwb.active

        for col in readws.iter_cols(values_only = True, min_col = 2):
            factor = col[0]
            col = col[1:]
            for idx, i in enumerate(col, 1): # 到時候可能不能用 enumerate 因為要用真的股票編號
                if results[year].get(idx) == None: results[year][idx] = dict()
                try: results[year][idx][factor] = float(i)
                except: results[year][idx][factor] = "Error"
                # print(idx, prices[idx])
    return results

def customSort(results: dict, year: int, factors: list, numbers: int = 5) -> list:
    """
    This function can return the best stock numbers to buy
    parameters:
        results: results from getData function
        year: the target year
        factors:
            a list of factor, each factor is formatted like (factorName, multiple)
            the index would
    
    """
    dataList = [(stock, data) for stock, data in results[year].items()]
    
    indexed = []
    for factor, multi in factors:
        tmp = set()
        for stock, data in dataList:
            if type[data[factor]] != str: tmp.add(data[factor])
        indexed[factor] = sorted(list(tmp))

    def customKey(item: tuple): # 用來排序的判準
        key = 0
        for factor in factors:
            if type(item[1][factor]) != str: key +=

    dataList.sort(key = customKey, reverse = True) # 大到小
    buyList = list()
    for stock, data in dataList[:numbers]:
        buyList.append(stock)
    return buyList


wb = openpyxl.Workbook()
ws = wb.active
def addData(buyList: list, year: int) -> None:
    if year == beginYear: ws.append(["年份", "以下為購買的對象編號"])
    ws.append([year] + buyList)

results = getData()
for year in range(beginYear, endYear + 1):
    buyList = customSort(results, year, factors = factors, numbers = numbers) # 每年購買
    addData(buyList, year) # add new results to sheet


wb.save(f"購買標的({','.join(factors)}).xlsx")