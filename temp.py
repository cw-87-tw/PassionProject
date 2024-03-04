from selenium import webdriver
from os import system as sys
from time import sleep
driver = webdriver.Firefox()

# tools
def clickButton(xpath): driver.find_element(by = "xpath", value = xpath).click()
def boxType(xpath, msg):
    box = driver.find_element(by = "xpath", value = xpath)
    box.send_keys(msg)
def getText(xpath): return driver.find_element(by = "xpath", value = xpath).text

# process
def login():
    driver.get('https://statementdog.com/users/sign_in')
    boxType(r'//*[@id="user_email"]', "h1110539@stu.wghs.tp.edu.tw")
    boxType(r'//*[@id="user_password"]', "passion")
    clickButton("/html/body/div[3]/div[1]/form/div/button")
    sleep(1)

def setSeasonly():
    driver.get("https://statementdog.com/analysis/2330/monthly-revenue")
    sleep(1)
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[1]/i")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/ul/li[4]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[1]/select/option[14]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[2]/select/option[23]")
    clickButton(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[1]/div[1]/table/tr/td[1]/div[2]/div[3]/div")

def test():
    driver.get("https://statementdog.com/analysis/2330/income-statement")
    sleep(1)
    print(getText(r"/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[1]/ul/li[2]/table/tr[2]/td[1]"))


login()
setSeasonly()
test()
# sys("pause")
# cookies = driver.get_cookies()
# print(cookies)