try:
    from selenium import webdriver
except:
    import os
    os.system("pip install selenium")
    from selenium import webdriver
from time import sleep
from selenium.webdriver.firefox.options import Options
from os import system


options = Options()
options.add_argument("--headless") # 隱藏瀏覽器視窗
driver = webdriver.Firefox(options = options)

# tools
def clickButton(xpath): driver.find_element(by = "xpath", value = xpath).click()
def boxType(xpath, msg):
    box = driver.find_element(by = "xpath", value = xpath)
    box.send_keys(msg)
def getText(xpath): return driver.find_element(by = "xpath", value = xpath).text


emails = []

for i in range(100):
    while True:
        try:
            driver.delete_all_cookies()
            # get new email
            driver.get("https://generator.email/")
            clickButton("/html/body/div[3]/div/div/p/a[1]/button")
            email = getText("/html/body/div[3]/div/div/b/span")
            # sign up
            try:
                driver.get("https://statementdog.com/feeds")
                clickButton("/html/body/div[1]/nav/div/div[2]/ul/li[3]/button/i[2]")
                clickButton("/html/body/div[1]/nav/div/div[2]/ul/li[3]/button/div/div[2]/ul/li[4]/button")
            except: pass
            driver.get("https://statementdog.com/users/sign_up")
            clickButton("/html/body/div[2]/form/div/div[3]")
            boxType(r'//*[@id="user_email"]', email)
            boxType(r'//*[@id="user_password"]', "passion")
            boxType(r'//*[@id="user_password_confirmation"]', "passion")
            clickButton("/html/body/div[2]/form/div/div[4]/button")
            if driver.current_url == "https://statementdog.com/users/confirmation_instruction": break
            else: 
                print("Failed\n")
                with open("output.txt", "w", encoding = "utf-8") as file:
                    for i in emails:
                        file.write(i + "\n")
                driver.quit()
                driver = webdriver.Firefox(options = options)
        except: pass

    emails.append(email)
    sleep(0.5)
    # get confirmation
      
    while 1:
        try:
            sleep(2)
            driver.get("https://generator.email/")
            clickButton("/html/body/div[3]/div/div/p/a[2]/button/span")  
            text = getText("/html/body/div[4]/div/div/div/div[2]/div[2]/div[4]/div[3]/div/div/div/p[2]")
            driver.get(text[text.find('h'):])
            break
        except: pass
    
    # clickButton("/html/body/div[4]/div/div/div/div[2]/div[2]/div[4]/div[3]/div/div/div/p[1]/span/a")

    print(email)
with open("output.txt", "a", encoding = "utf-8") as file:
    for i in emails:
        file.write(i + "\n")

print("\n\n\n\n\n")
print(*emails, sep = "\n")
