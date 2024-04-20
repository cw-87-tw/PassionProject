import openpyxl

beginYear = 2014 
beginYear = int(input("beginYear:"))
endYear = 2023 
endYear = int(input("endYear:"))
results = dict()

def getData():
    for year in range(beginYear, endYear + 1):
        results[year] = dict()
        readwb = openpyxl.load_workbook(f"./results/{year}年.xlsx")
        readws = readwb.active

        for col in readws.iter_cols(values_only = True, min_col = 2):
            factor = col[0]
            col = col[1:]
            for idx, i in enumerate(col, 1):
                if results[year].get(idx) == None: results[year][idx] = dict()
                results[year][idx][factor] = i if type(i) != str else 1e20
                # print(idx, prices[idx])
    
