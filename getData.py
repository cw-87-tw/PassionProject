from selenium import webdriver
from twstock import Stock
import openpyxl
from time import sleep
from selenium.webdriver.firefox.options import Options

options = Options()
# options.add_argument("--headless") # 隱藏瀏覽器視窗
driver = webdriver.Firefox(options = options)
# print("begin web driver")

# tools
def clickButton(xpath): 
    for i in range(10):
        try: 
            driver.find_element(by = "xpath", value = xpath).click()
            return  
        except: 
            sleep(0.3)
    print("unexpected press button\n\n\n")
def boxType(xpath, msg):
    for i in range(10):
        try: 
            box = driver.find_element(by = "xpath", value = xpath)
            box.send_keys(msg)
            return
        except: 
            sleep(0.3)
    print("unexpected type\n\n\n")
def getText(xpath): 
    for i in range(10):
        try: 
            return driver.find_element(by = "xpath", value = xpath).text
        except: 
            sleep(0.3)
    print("unexpected get text\n\n\n")
    return "Error"
# basic information
target = "2330"
beginYear = 2004
index = 1
beginYear = 2004
endYear = 2023
season = 1
delay = 0.2

# process
def read_stock_numbers_from_excel(filename):
    stock_numbers = []
    readwb = openpyxl.load_workbook(filename)
    readws = readwb.active
    for row in readws.iter_rows(values_only=True):
        stock_numbers.append(row[2])
    # print("stock numbers:", *stock_numbers)
    print("got stock numbers", len(stock_numbers))
    return stock_numbers

def save_to_excel(results : list, searchYear, searchSeason):
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
    wb.save(f"./results/{searchYear}年第{searchSeason}季.xlsx")

def setInformation(_beginYear, _endYear, _target):
    global target
    global index
    global beginYear
    global endYear
    target = _target
    beginYear = _beginYear
    endYear = _endYear
    results = dict()
    for year in range(beginYear, endYear + 1):
        results[year] = dict()
        for season in range(1, 5):
            results[year][season] = dict()
    # index = (int(year) - beginYear) * 4 + int(season)
    # print("set", year, season, target, index)
    return results

def login(id):
    driver.get('https://statementdog.com/users/sign_in')
    try:
        boxType(r'//*[@id="user_email"]', id)
        boxType(r'//*[@id="user_password"]', "passion")
        clickButton("/html/body/div[3]/div[1]/form/div/button")
    except:
        clickButton("/html/body/div[1]/nav/div/div[2]/ul/li[3]/button/i[2]")
        clickButton("/html/body/div[1]/nav/div/div[2]/ul/li[3]/button/div/div[2]/ul/li[4]/button")
        driver.get('https://statementdog.com/users/sign_in')
        boxType(r'//*[@id="user_email"]', id)
        boxType(r'//*[@id="user_password"]', "passion")
        clickButton("/html/body/div[3]/div[1]/form/div/button")
        # driver.
    sleep(1)

def setRange(url):
    driver.get(url)
    sleep(delay)
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[1]/i")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/ul/li[4]")
    #2004
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[1]/select/option[04]")
    #2023
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[2]/select/option[23]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[3]/div")

def getPrice(results: dict):
    print("Get Price:", target)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            try:
                stock = Stock(str(target))
                results[year][season]["股價"] = stock.fetch(year, season * 3)[0].close
            except:
                results[year][season]["股價"] = "Error"
    return results

def getIncome(results : dict):
    print("Get Income:", target)
    setRange(f"https://statementdog.com/analysis/{target}/income-statement")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["營收"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
            results[year][season]["營業利益"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[8]/td[{index}]")
            results[year][season]["稅後淨利"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[10]/td[{index}]")
    return results


def getCash(results: dict):
    print("Get Cash:", target)
    setRange(f"https://statementdog.com/analysis/{target}/cash-flow-statement")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["營業現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[4]/td[{index}]")
            results[year][season]["投資現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[5]/td[{index}]")
    return results

def getRoeRoa(results: dict):
    print("Get ROE ROA:", target)
    setRange(f"https://statementdog.com/analysis/{target}/roe-roa")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["ROE"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
            results[year][season]["ROA"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
    return results

def getDebt(results: dict):
    print("Get Debt:", target)
    setRange(f"https://statementdog.com/analysis/{target}/liabilities-and-equity")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["總負債"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[11]/td[{index}]")
            results[year][season]["流動負債"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[8]/td[{index}]")
    return results

def getAssets(results: dict):
    print("Get Assets:", target)
    setRange(f"https://statementdog.com/analysis/{target}/assets")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["固定資產"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[9]/td[{index}]")
            results[year][season]["總資產"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[11]/td[{index}]")
            results[year][season]["流動資產"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[7]/td[{index}]")
    return results

def getEps(results: dict):
    print("Get EPS:", target)
    setRange(f"https://statementdog.com/analysis/{target}/eps")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            index = (int(year) - beginYear) * 4 + int(season)
            results[year][season]["單季EPS"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
    return results

    
def safe_divide(numerator, denominator):
    numerator = str(numerator)
    denominator = str(denominator)
    try:
        return float(numerator.replace(",", "")) / float(denominator.replace(",", ""))
    except Exception:
        # print(numerator, denominator, "Error")
        # print(Exception)
        return "Error"

def safe_subtract(a, b):
    a = str(a)
    b = str(b)
    try:
        return float(a.replace(",", "")) - float(b.replace(",", ""))
    except:
        return "Error"

def getOther(results):
    print("Get Other:", target)
    global beginYear
    global endYear
    
    for year in range(beginYear, endYear + 1):
        for season in range(1, 5):
            results[year][season]["股數"] = safe_divide(results[year][season]["稅後淨利"], results[year][season]["單季EPS"])
            results[year][season]["淨資產"] = safe_subtract(results[year][season]["總資產"], results[year][season]["總負債"])
            results[year][season]["每股淨值"] = safe_divide(results[year][season]["淨資產"], results[year][season]["股數"])
            results[year][season]["淨流動資產"] = safe_subtract(results[year][season]["流動資產"], results[year][season]["總負債"])
            results[year][season]["長期負債"] = safe_subtract(results[year][season]["總負債"], results[year][season]["流動負債"])
            results[year][season]["營業利率"] = safe_divide(results[year][season]["稅後淨利"], results[year][season]["營收"])
            results[year][season]["自由資本比率"] = safe_divide(results[year][season]["固定資產"], results[year][season]["總資產"])
            results[year][season]["負債除以本期損益"] = safe_divide(results[year][season]["總負債"], results[year][season]["營業利益"])
            results[year][season]["本益比"] = safe_divide(results[year][season]["單季EPS"], results[year][season]["股價"])
            results[year][season]["股價淨值比"] = safe_divide(results[year][season]["股價"], results[year][season]["每股淨值"])
    return results

def getYearROA(results):
    print("Get ROA:", target)
    setRange(f"https://statementdog.com/analysis/{target}/roe-roa")
    sleep(delay)
    clickButton("/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[3]/ul/li[2]")
    sleep(delay)
    global beginYear
    global endYear
    for year in range(beginYear, endYear + 1):
        index = year - beginYear + 1
        results[year] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
        try: 
            results[year] = float(results[year])
            if results[year] < 5: results[year] = 0
        except: results[year] = 0
    return results
            

# getInformation()
# getDebt()
# getIncome()
# getCash()
# getRoeRoa()
# getAssets()
# getPrice()


# driver.quit()    