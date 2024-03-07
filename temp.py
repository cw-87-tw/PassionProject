from selenium import webdriver
from os import system as sys
from time import sleep
from selenium.webdriver.firefox.options import Options

options = Options()
# options.add_argument("--headless") # 顯示瀏覽器視窗
driver = webdriver.Firefox(options = options)

# tools
def clickButton(xpath): driver.find_element(by = "xpath", value = xpath).click()
def boxType(xpath, msg):
    box = driver.find_element(by = "xpath", value = xpath)
    box.send_keys(msg)
def getText(xpath): return driver.find_element(by = "xpath", value = xpath).text

# basic information
target = "2330"
beginYear = 2009
index = 1

# process
def getInformation():
    global target
    global index
    target = input("輸入查詢編號:")
    year = int(input("查詢年份:"))
    season = int(input("查詢季度:"))
    index = (year - 2009) * 4 + season

def login():
    driver.get('https://statementdog.com/users/sign_in')
    boxType(r'//*[@id="user_email"]', "h1110539@stu.wghs.tp.edu.tw")
    boxType(r'//*[@id="user_password"]', "passion")
    clickButton("/html/body/div[3]/div[1]/form/div/button")
    sleep(1)

def setRange(url):
    driver.get(url)
    sleep(1)
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[1]/i")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/ul/li[4]")
    #2014
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[1]/select/option[09]")
    #2023
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[2]/select/option[23]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[3]/div")

def getIncome():
    setRange(f"https://statementdog.com/analysis/{target}/income-statement")
    sleep(2)
    result = dict()
    result["營收"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[{index}]")
    result["毛利"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[3]/td[{index}]")
    result["營業利益"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[8]/td[{index}]")
    result["稅前利益"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/ul/li[2]/table/tr[9]/td[{index}]")
    print(result)
    return result

def getCash():
    setRange(f"https://statementdog.com/analysis/{target}/cash-flow-statement")
    sleep(2)
    result = dict()
    result["營業現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[4]/td[{index}]")
    result["投資現金流"] = getText(f"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[5]/td[{index}]")
    result[""]

# getInformation()
login()
getIncome()

# driver.quit()    

# test()
# sys("pause")
# cookies = driver.get_cookies()
# print(cookies)