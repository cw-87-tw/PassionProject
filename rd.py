from random import sample
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.append(["年", "購買1", "購買2", "購買3", "購買4", "購買5"])

l = range(1, 101)

for year in range(2014, 2024):
    ws.append([year] + sample(l, 5))

wb.save("random.xlsx")