import tkinter as tk

dict = {1: "一", 2: "二"}

root = tk.Tk()
root.title("投入产出表标准化数据分析")
root.geometry("800x400")

listbox = tk.Listbox(
    root,
    selectmode=tk.MULTIPLE,
    bd=0,
    selectbackground="#3A7FF6"
    )
for key in dict.keys():
    listbox.insert(tk.END, str(key) + ' ' + dict[key])

listbox.pack()
root.mainloop()


