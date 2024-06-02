from os import system as sys

try:
    sys("pip install selenium")
    sys("pip install twstock")
    sys("pip install openpyxl")

except:
    print("Python 安裝未成功，或者是 pip 有安裝問題")
    sys("pause")
    exit()

try:
    from selenium import webdriver
    from twstock import Stock
    import openpyxl
    from selenium.webdriver.firefox.options import Options
    options = Options()
    options.add_argument("--headless") # 隱藏瀏覽器視窗
    driver = webdriver.Firefox(options = options)
    driver.quit()
except:
    print("環境未設置成功，可能是 Python, Firefox, pip 有安裝問題")
    sys("pause")
    exit()

print("環境架設已完成，可以進到下個步驟")
sys("pause")