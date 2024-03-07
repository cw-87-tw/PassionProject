import tkinter as tk
from tkinter import ttk
from getData import *

root = tk.Tk()

def init():
    root.title("Passion 財務報表查詢")
    global entry_target
    global entry_year
    global entry_season
    label_target = tk.Label(root, text = "輸入查詢編號:")
    entry_target = tk.Entry(root)
    label_year = tk.Label(root, text = "查詢年份:")
    entry_year = tk.Entry(root)
    label_season = tk.Label(root, text = "查詢季度:")
    entry_season = tk.Entry(root)
    button_search = tk.Button(root, text = "查詢", command = search)
    global tree
    tree = ttk.Treeview(root, columns = ("Attribute", "Value"), show = "headings")
    tree.heading("Attribute", text = "Attribute")
    tree.heading("Value", text = "Value")

    label_target.grid(row = 0, column = 0)
    entry_target.grid(row = 0, column = 1)
    label_year.grid(row = 1, column = 0)
    entry_year.grid(row = 1, column = 1)
    label_season.grid(row = 2, column = 0)
    entry_season.grid(row = 2, column = 1)
    button_search.grid(row = 3, columnspan = 2)
    tree.grid(row = 4, columnspan = 2)

def search():
    setInformation(entry_year.get(), entry_season.get(), entry_target.get())
    login()
    results = dict()
    results.update(getIncome())
    results.update(getCash())
    results.update(getDebt())
    results.update(getRoeRoa())
    show_result(result = results)
    
def show_result(result):
    for key, value in result.items():
        tree.insert("", "end", values=(key, value))


init()
root.mainloop() # start program