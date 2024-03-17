try:
    from selenium import webdriver
except:
    import os
    os.system("pip install selenium")
    from selenium import webdriver
from time import sleep
from selenium.webdriver.firefox.options import Options

options = Options()
options.add_argument("--headless") # 隱藏瀏覽器視窗
driver = webdriver.Firefox(options = options)
print("start web driver")

# tools
def clickButton(xpath): driver.find_element(by = "xpath", value = xpath).click()
def boxType(xpath, msg):
    box = driver.find_element(by = "xpath", value = xpath)
    box.send_keys(msg)
def getText(xpath): return driver.find_element(by = "xpath", value = xpath).text

# basic information
target = "2330"
beginYear = 2010
index = 1
year = 2010
season = 1

# process
def setInformation(_year, _season, _target):
    global target
    global index
    global year
    global season
    target = _target
    year = _year
    season = _season
    # target = input("輸入查詢編號:")
    # year = int(input("查詢年份:"))
    # season = int(input("查詢季度:"))
    index = (int(year) - 2010) * 4 + int(season)

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
    sleep(1)
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[1]/i")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/ul/li[4]")
    #2010
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[1]/select/option[10]")
    #2023
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[2]/select/option[23]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[3]/div")

def getPrice():
    driver.get("https://www.twse.com.tw/zh/trading/historical/stock-day.html")
    clickButton(f"/html/body/div[1]/div/div[2]/main/form/div/div[1]/div[1]/span/select[1]/option[{beginYear - year + 15}]")
    clickButton(f"/html/body/div[1]/div/div[2]/main/form/div/div[1]/div[1]/span/select[2]/option[{season * 3}]")
    boxType(f"/html/body/div[1]/div/div[2]/main/form/div/div[1]/div[2]/input", target)
    clickButton("/html/body/div[1]/div/div[2]/main/form/div/div[1]/div[3]/button")
    result = {"股價" : getText("/html/body/div[1]/div/div[2]/main/div[2]/div[2]/div[2]/table/tbody/tr[1]/td[7]")}
    print(result)
    return result

def getIncome():
    setRange(f"https://statementdog.com/analysis/{target}/income-statement")
    sleep(1)
    result = dict()
    result["營收"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
    result["毛利"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
    result["營業利益"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[8]/td[{index}]")
    result["稅後利益"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[10]/td[{index}]")
    print(result)
    return result

def getCash():
    setRange(f"https://statementdog.com/analysis/{target}/cash-flow-statement")
    sleep(1)
    result = dict()
    result["營業現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[4]/td[{index}]")
    result["投資現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[5]/td[{index}]")
    result["融資現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[6]/td[{index}]")
    result["自由現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[8]/td[{index}]")
    print(result)
    return result

def getRoeRoa():
    setRange(f"https://statementdog.com/analysis/{target}/roe-roa")
    sleep(1)
    result = dict()
    result["ROE"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
    result["ROA"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
    print(result)
    return result

def getDebt():
    setRange(f"https://statementdog.com/analysis/{target}/liabilities-and-equity")
    sleep(1)
    result = dict()
    result["總負債"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[11]/td[{index}]")
    print(result)
    return result

def getAssets():
    setRange(f"https://statementdog.com/analysis/{target}/assets")
    sleep(1)
    result = dict()
    result["固定資產"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[9]/td[{index}]")
    result["總資產"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[11]/td[{index}]")
    print(result)
    return result

def getEps():
    setRange(f"https://statementdog.com/analysis/{target}/eps")
    sleep(1)
    result = dict()
    result["單季EPS"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
    print(result)
    return result
    
# getInformation()
# login("charlesennor@igdinhcao.com")
# getDebt()
# getIncome()
# getCash()
# getRoeRoa()
# getAssets()
getPrice()


# driver.quit()    

# test()
# sys("pause")
# cookies = driver.get_cookies()
# print(cookies)