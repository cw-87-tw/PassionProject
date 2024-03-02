from copyreg import clear_extension_cache
from datetime import datetime as dt
from time import sleep
from pause import until as ut
import webbrowser
import os
from msedge.selenium_tools import Edge
from msedge.selenium_tools import EdgeOptions
from termcolor import colored
from os import system as sys
import webbrowser
import threading
def wait(): sleep(5)

#Edge Setup
opt = EdgeOptions()
opt.use_chromium = True
opt.add_argument("--disable-infobars")
opt.add_argument("start-maximized")
opt.add_argument("--disable-extensions")
opt.add_experimental_option("prefs", { \
    "profile.default_content_setting_values.media_stream_mic": 1, 
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.media_stream_notifications": 1,
    "profile.default_content_setting_values.notifications": 1
  })

def openClass(driver, url):
    global status
    driver.get(url)
    status = 1
    return

# The screen clear function
def screen_clear():
       # for mac and linux(here, os.name is 'posix')
   if os.name == 'posix':
          _ = os.system('clear')
   else:
          # for windows platfrom
      _ = os.system('cls')

user = "0"


def main():
    driver = Edge(options = opt)

    #Login
    driver.get("https://www.google.com")
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div/div/div[2]/a").click()
    wait()
    address = "1061134@stu.wghs.tp.edu.tw" # 帳號
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[2]/div/c-wiz/div/div[2]/div/div[1]/div/form/span/section/div/div/div[1]/div/div[1]/div/div[1]/input").send_keys(address + "\n")
    wait()
    pw = "31974625" # 密碼
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/div[2]/div/c-wiz/div/div[2]/div/div[1]/div/form/span/div[1]/div[1]/div/div/div/div/div[1]/div/div[1]/input").send_keys(pw + "\n")
    wait()

    with open(f".\\data\\{dt.now().isoweekday()}.txt", mode = "r", encoding = "utf-8") as file:
        classes = file.read().splitlines()
    with open(".\\data\\classroom.txt", mode = "r", encoding = "utf-8") as file:
        l = file.read().splitlines()
    classroom = {}
    for i in l:
        temp = i.split()
        classroom[temp[0]] = temp[1]
    with open(".\\data\\meet.txt", mode = "r", encoding = "utf-8") as file:
        l = file.read().splitlines()
    meet = {}
    for i in l:
        temp = i.split()
        meet[temp[0]] = temp[1]
    with open(".\\data\\time.txt", mode = "r", encoding = "utf-8") as file:
        time = file.read().splitlines()
    td = dt.now()




    for i in range(len(time)):
        nowt = time[i].split()
        if dt.now().hour > int(nowt[0]): continue
        if classes[i] == "0": continue
        while 1:
            if dt.now().hour == int(nowt[0]) and dt.now().minute >= int(nowt[1]): break
            sleep(1)
            screen_clear()
            print(f"\r現在時間 {dt.now().hour}:{dt.now().minute}:{dt.now().second}")
            print(colored("下堂課:", "cyan"), colored(classes[i], "yellow"))
        if(meet[classes[i]][0:23] == "https://us02web.zoom.us"): 
            webbrowser.open(meet[classes[i]])
            continue
        
        # open meet
        global status
        status = 0
        thread = threading.Thread(target = openClass, args = (driver, meet[classes[i]] + "&hs=179"))
        thread.start()
        sleep(5)
        if status == 0: return True
        wait()
        cnt = 5
        while 1:
            if (cnt == 0): break
            cnt -= 1
            # driver.find_element_by_xpath("/html/body/c-wiz/div/div[2]/div[3]/div[1]/button/span").click()
            # wait()
            try: 
                driver.find_element_by_xpath("/html/body/div[1]/c-wiz/div/div/div[10]/div[3]/div/div[1]/div[4]/div/div/div[1]/div[1]/div/div[4]/div[1]/div/div/div[1]").click()
                print(colored("Mic off", "green"))
                driver.find_element_by_xpath("/html/body/div[1]/c-wiz/div/div/div[10]/div[3]/div/div[1]/div[4]/div/div/div[1]/div[1]/div/div[4]/div[2]/div/div[1]").click()
                print(colored("Camera off", "green"))
                driver.find_element_by_xpath("/html/body/div[1]/c-wiz/div/div/div[10]/div[3]/div/div[1]/div[4]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/button").click()
                print(colored("Joined call", "green"))
                wait()
                break
            except: 
                print(colored("Cannot join", "red"))
                # sleep(55)
        # 第一版的code
        # try:
        #     webbrowser.open(classroom[classes[i]])
        # except:
        #     pass
        # try:
        #     webbrowser.open(meet[classes[i]])
        #     sleep(2)
        #     pyautogui.press("f11")
        #     sleep(5)
        #     pyautogui.click(550, 780, duration = 0.2)
        #     pyautogui.click(670, 780, duration = 0.2)
        #     pyautogui.click(1394, 586, duration = 0.2)
        # except:
        #     pass
while(main()): pass
sleep(3600)