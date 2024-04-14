import openpyxl

readwb = openpyxl.load_workbook(r"P:\Code\PassionProject\results\年ROA.xlsx")
readws = readwb.active
data = []
for col in readws.iter_cols(min_col = 2, values_only=False):
    data.append([])
    cnt = 1
    for cell in col[1:]:
        x = cell.value
        if x == None: x = 0
        data[-1].append((x, cnt))
        cnt += 1

year = 2014
for yearData in data:
    print(f"{year}年的最佳5個:")
    year += 1
    cnt = 0
    for i in sorted(yearData, reverse = True):
        print(f"編號 {i[1]}, ROA {i[0]}")
        cnt += 1
        if cnt == 5: break
# print("stock numbers:", *stock_numbers)
# print(data)
