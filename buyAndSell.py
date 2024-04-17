import openpyxl

buySheet = "./template1.xlsx"
outputName = ""
money = 5000000

wb = openpyxl.load_workbook(buySheet)
ws = wb.active

allTargets = dict()
for col in ws.iter_cols(values_only=True, min_col = 2):
    allTargets[col[0]] = list()
    for cell in col[1:]:
        if cell != None: allTargets[col[0]].append(cell)

print(allTargets)

def getPrices(year: int) -> dict:
    readwb = openpyxl.load_workbook(f"./results/{year}年.xlsx")
    readws = readwb.active
    idx = 1
    prices = dict()
    for col in readws.iter_cols(values_only = True, min_col = 3, max_col = 3):
        for i in col[1:]:
            prices[idx] = float(i) if str(i) != "Error" else 1e20
            idx += 1
    return prices

hold = [] # (number, 股數)
wb = openpyxl.Workbook()
ws = wb.active

for year, targets in allTargets.items():
    prices = getPrices(year) # get this year's prices

    # sell old
    for no, shares in hold:
        money += shares * prices[no]
    hold.clear()

    ws.append([f"{year}年"])
    ws.append(["原有資金", money])
    ws.append([])
    ws.append(["購買編號", "股價", "股數", "總價"])

    # buy new
    perMoney = money / len(targets)
    for stock in targets:
        shares = perMoney // prices[stock]
        money -= prices[stock] * shares
        ws.append([stock, prices[stock], shares, prices[stock] * shares])
        print(f"{year}購買了{stock}，買了{shares}股")
        hold.append((stock, year))

    ws.append([])
    ws.append(["剩餘資金", money])
    ws.append([])
    ws.append([])
    ws.append([])
    
wb.save("./購買分析結果.xlsx")