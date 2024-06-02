# Passion Project 最終成品: 股票分析工具

這是個幫助分析股票的工具，在說明功能之前，必須要先聲明一點重要資訊: **以下的工具雖然可以幫助你在投資上有更清楚的方向和方法，但這世界並不存在任何穩賺不賠的方法，且在過去找到的最佳方法並不會在未來也是最佳**

## 功能:
1. 蒐集從 2000 年開始的財報 (配合[財報狗](https://statementdog.com/)上的資料)、股價 (來自模組[twstock](https://twstock.readthedocs.io/zh-tw/latest/))
2. 用可調整權重的指標做各年的排序
3. 透過模擬買賣來分析漲跌與損益

## 在開始之前
1. 請先確認電腦是否安裝 Python 3 (建議安裝 Python 3.10 以上版本) 可點[此連結](https://www.python.org/downloads/)下載最新版 Python
2. 需要安裝 Firefox 瀏覽器
   1. 首先，必須要安裝好 Firefox，[此為下載連結](https://www.mozilla.org/en-US/firefox/new/)
   2. 接著，需要安裝爬蟲驅動，即是目前已經安裝好在資料夾內的 "getckodriver.exe"，只是若有版本問題，可以在[此連結](https://github.com/mozilla/geckodriver/releases)找到對應的其他版本
3. 可以先執行 "0. setup.py" 確認環境是否完全安裝成功
## 工具的細項說明

### 工具 1: 網路爬蟲