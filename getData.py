from selenium import webdriver
from twstock import Stock

    

    
from time import sleep
from selenium.webdriver.firefox.options import Options

options = Options()
# options.add_argument("--headless") # 隱藏瀏覽器視窗
driver = webdriver.Firefox(options = options)
# print("begin web driver")

# tools
def clickButton(xpath): 
    try: driver.find_element(by = "xpath", value = xpath).click()
    except: pass
def boxType(xpath, msg):
    try:
        box = driver.find_element(by = "xpath", value = xpath)
        box.send_keys(msg)
    except: pass
def getText(xpath): 
    while 1:
        try: 
            return driver.find_element(by = "xpath", value = xpath).text
        except: 
            sleep(0.3)

# basic information
target = "2330"
beginYear = 2004
index = 1
beginYear = 2004
endYear = 2023
season = 1
delay = 1

# process
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


    
# getInformation()
# getDebt()
# getIncome()
# getCash()
# getRoeRoa()
# getAssets()
# getPrice()


# driver.quit()    