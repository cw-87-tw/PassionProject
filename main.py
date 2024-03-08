import tkinter as tk
from tkinter import ttk
from getData import *

root = tk.Tk()

def init():
    root.title("Passion 財務報表查詢")
    root.geometry("800x600")

    for i in range(6):
        root.rowconfigure(i, weight=1)
    for i in range(2):
        root.columnconfigure(i, weight=1)

    global entry_target
    global entry_year
    global entry_season
    global entry_user
    label_target = tk.Label(root, text="輸入查詢編號:", font=("Helvetica", 12))  # 設置字型大小為12
    entry_target = tk.Entry(root)
    label_year = tk.Label(root, text="查詢年份:", font=("Helvetica", 12))  # 設置字型大小為12
    entry_year = tk.Entry(root)
    label_season = tk.Label(root, text="查詢季度:", font=("Helvetica", 12))  # 設置字型大小為12
    entry_season = tk.Entry(root)
    label_user = tk.Label(root, text="使用者:", font=("Helvetica", 12))  # 設置字型大小為12
    entry_user= tk.Entry(root)
    global button_search
    button_search = tk.Button(root, text="查詢", command=search)
    global tree
    tree = ttk.Treeview(root, columns=("Attribute", "Value"), show="headings")
    tree.heading("Attribute", text="Attribute")
    tree.heading("Value", text="Value")
    global previousResult
    previousResult = list()

    label_target.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    entry_target.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
    label_year.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
    entry_year.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")
    label_season.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
    entry_season.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")
    label_user.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
    entry_user.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
    button_search.grid(row=4, columnspan=2, padx=10, pady=10, sticky="nsew")
    tree.grid(row=5, columnspan=2, padx=10, pady=10, sticky="nsew")

    # 設置Treeview文字大小
    style = ttk.Style()
    style.configure("Treeview", font=("Helvetica", 14))  # 設置字型大小為14

def search():
    button_search.config(text="搜尋中...請稍候", state = "disabled")
    root.update()
    setInformation(entry_year.get(), entry_season.get(), entry_target.get())
    login(entry_user.get())
    results = dict()
    try:
        results.update(getIncome())
        results.update(getCash())
        results.update(getDebt())
        results.update(getRoeRoa())
        pass
    except:
        results["Error"] = "出現錯誤"
    button_search.config(text="查詢", state="normal")  # 還原按鈕文字並啟用按鈕
    show_result(result=results)
    root.update()
    
    
def show_result(result):
    global previousResult
    for i in previousResult: tree.remove(i)
    previousResult.clear()
    for key, value in result.items():
        previousResult.append(tree.insert("", "end", values=(key, value)))
    # previousResult = result


init()
root.mainloop() # start program
driver.quit()
