import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from getData import *
import tkinter.messagebox as mb

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("股票查詢與儲存工具")

        self.stocks_sheet = ""
        self.save_dir = ""
        self.begin_year = 2001
        self.end_year = 2023
        self.account = ""
        self.password = ""

        self.label1 = ttk.Label(root, text="查詢的股票編號表格:")
        self.label1.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.stocks_sheet_entry = ttk.Entry(root)
        self.stocks_sheet_entry.grid(row=0, column=1, padx=5, pady=5)
        
        self.browse_stocks_button = ttk.Button(root, text="瀏覽...", command=self.browse_stocks_sheet)
        self.browse_stocks_button.grid(row=0, column=2, padx=5, pady=5)

        self.label2 = ttk.Label(root, text="存 results 的資料夾:")
        self.label2.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.save_dir_entry = ttk.Entry(root)
        self.save_dir_entry.grid(row=1, column=1, padx=5, pady=5)
        
        self.browse_save_button = ttk.Button(root, text="瀏覽...", command=self.browse_save_dir)
        self.browse_save_button.grid(row=1, column=2, padx=5, pady=5)

        self.label3 = ttk.Label(root, text="開始年份:")
        self.label3.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.begin_year_entry = ttk.Entry(root)
        self.begin_year_entry.grid(row=2, column=1, padx=5, pady=5)
        self.begin_year_entry.insert(0, self.begin_year)

        self.label4 = ttk.Label(root, text="結束年份:")
        self.label4.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.end_year_entry = ttk.Entry(root)
        self.end_year_entry.grid(row=3, column=1, padx=5, pady=5)
        self.end_year_entry.insert(0, self.end_year)

        self.label5 = ttk.Label(root, text="帳號:")
        self.label5.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.account_entry = ttk.Entry(root)
        self.account_entry.grid(row=4, column=1, padx=5, pady=5)

        self.label6 = ttk.Label(root, text="密碼:")
        self.label6.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.password_entry = ttk.Entry(root, show="*")
        self.password_entry.grid(row=5, column=1, padx=5, pady=5)

        self.run_button = ttk.Button(root, text="執行", command=self.run)
        self.run_button.grid(row=6, column=1, padx=5, pady=5)

    def browse_stocks_sheet(self):
        self.stocks_sheet = filedialog.askopenfilename()
        self.stocks_sheet_entry.delete(0, tk.END)
        self.stocks_sheet_entry.insert(0, self.stocks_sheet)

    def browse_save_dir(self):
        self.save_dir = filedialog.askdirectory()
        self.save_dir_entry.delete(0, tk.END)
        self.save_dir_entry.insert(0, self.save_dir)

    def run(self):
        self.stocks_sheet = self.stocks_sheet_entry.get()
        self.save_dir = self.save_dir_entry.get()
        self.begin_year = int(self.begin_year_entry.get())
        self.end_year = int(self.end_year_entry.get())
        self.account = self.account_entry.get()
        self.password = self.password_entry.get()
        search(self.stocks_sheet, self.begin_year, self.end_year, self.save_dir, self.account, self.password)

def search(stockSheet, beginYear, endYear, dir, account, password):
    try:
        stock_numbers = read_stock_numbers_from_excel(stockSheet)
        cnt = 1
        login(account, password)
        finalResult = {}
        for year in range(beginYear, endYear + 1):
            finalResult[year] = list()
        
        for stock in stock_numbers:
            # print("目前進度:", searchYear, searchSeason, cnt)
            cnt += 1 
            results = setInformation(beginYear, endYear, stock)
            results = getPrice(results)
            results = getIncome(results)
            results = getCash(results)
            results = getRoeRoa(results)
            results = getDebt(results)
            results = getAssets(results)
            results = getEps(results)
            results = getOther(results)
            print(stock, results)
            for year in range(beginYear, endYear + 1):
                finalResult[year].append(results[year])
        
        for year in range(beginYear, endYear + 1):
            save_to_excel(finalResult[year], year, dir)
        mb.showinfo("成功", f"結果已儲存在: {dir}/results 資料夾內")
    except Exception as e:
        mb.showerror("錯誤", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
    driver.quit()
