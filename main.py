import os
import sys

import tkinter as tk
from tkinter import filedialog
import pandas as pd

from function import plot_industry_charts, plot_province_charts
from table import table_output

def relative_to_assets(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

dataframe = None
icon_button = None

elements = [start_button, data_button, load_button, back_start_button, back_coefficient_button, back_industry_button, listbox, \
industry_button, province_button, \
select_reaction_button, select_demand_button] = \
[None] * 11


is_reaction_global = True

# 跟新开始页面
def on_back_start_button_click():

    global icon_button, entry_1, entry_bg_1, text_1, text_3

    # 删除不需要元素
    if select_reaction_button is not None: select_reaction_button.destroy()
    if select_demand_button is not None: select_demand_button.destroy()
    if industry_button is not None: industry_button.destroy()
    if province_button is not None: province_button.destroy()

    if back_coefficient_button is not None: back_coefficient_button.destroy()
    if back_start_button is not None: back_start_button.destroy()
    if back_industry_button is not None: back_industry_button.destroy()

    # 添加文字格
    entry_image_1 = tk.PhotoImage(
    file=relative_to_assets("entry_1.png"))
    entry_bg_1 = canvas.create_image(
        571.4999999999999,
        190.5,
        image=entry_image_1
    )
    entry_1 = tk.Entry(
        bd=0,
        bg="#F1F5FF",
        fg="#000716",
        highlightthickness=0
    )
    entry_1.image = entry_image_1
    entry_1.place(
        x=410.9999999999999,
        y=160.0,
        width=321.0,
        height=59.0
    )

    # 添加图标按钮
    icon_button_image = tk.PhotoImage(
        file=relative_to_assets("button_2.png"))
    icon_button = tk.Button(
        image=icon_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=load_tables,
        relief="flat"
    )
    icon_button.image = icon_button_image
    icon_button.place(
        x=698.9999999999999,
        y=180.0,
        width=24.0,
        height=22.0
    )
    
    # 添加相关文字说明
    text_1 = \
    canvas.create_text(
        39.999999999999886,
        109.0,
        anchor="nw",
        text="投入产出表数据分析",
        fill="#FCFCFC",
        font=("Roboto Bold", 24 * -1)
    )

    text_3 = \
    canvas.create_text(
        454.9999999999999,
        72.0,
        anchor="nw",
        text="请点击按钮以选取文件",
        fill="#505485",
        font=("Roboto Bold", 24 * -1)
    )


# 跟新选择文件路径页面
def load_tables():
    global industry_dict, province_dict
    global input_file_path
    global load_button

    input_file_path = filedialog.askdirectory()
    output_file_path = input_file_path + "/分析结果"
    industry_dict, province_dict = table_output(input_file_path, output_file_path)

    # 删除不需要元素
    if icon_button is not None: icon_button.destroy()
    if entry_1 is not None: entry_1.destroy()
    if back_start_button is not None: back_start_button.destroy()

    canvas.delete(entry_bg_1)
    canvas.delete(text_3)

    # 添加"标准化分析"按钮
    load_button_image = tk.PhotoImage(
        file=relative_to_assets("button_1.png")
    )
    load_button = tk.Button(
        image=load_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_back_coefficient_button_click,
        relief="flat"
    )
    load_button.image = load_button_image
    load_button.place(
        x=484.9999999999999,
        y=173.0,
        width=180.0,
        height=55.0
    )


# 跟新感应/需求系数选择页面
def on_back_coefficient_button_click():
    print("标准化文档分析中...")

    global select_reaction_button, select_demand_button, back_start_button

    # 删除不需要元素
    if load_button is not None: load_button.destroy()
    if industry_button is not None: industry_button.destroy()
    if province_button is not None: province_button.destroy()
    if back_start_button is not None: back_start_button.destroy()
    if back_coefficient_button is not None: back_coefficient_button.destroy()
    if back_industry_button is not None: back_industry_button.destroy()


    # 添加"感应系数"按钮
    select_reaction_button_image = tk.PhotoImage(
    file=relative_to_assets("button_4.png")
    )
    select_reaction_button = tk.Button(
        image=select_reaction_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=lambda: on_select_dataframe_click(True),
        relief="flat"
    )
    select_reaction_button.image = select_reaction_button_image
    select_reaction_button.place(
        x=484.9999999999999,
        y=76.0,
        width=180.0,
        height=55.0
    )

    # 添加"需求系数"按钮
    select_demand_button_image = tk.PhotoImage(
        file=relative_to_assets("button_3.png")
    )
    select_demand_button = tk.Button(
        image=select_demand_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=lambda: on_select_dataframe_click(False),
        relief="flat"
    )
    select_demand_button.image = select_demand_button_image
    select_demand_button.place(
        x=484.9999999999999,
        y=173.0,
        width=180.0,
        height=55.0
    )

    # 添加“返回”按钮
    back_start_button_image = tk.PhotoImage(
        file=relative_to_assets("button_7.png")
    )
    back_start_button = tk.Button(
        image=back_start_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_back_start_button_click,
        relief="flat"
    )
    back_start_button.image = back_start_button_image
    back_start_button.place(
        x=484.9999999999999,
        y=270.0,
        width=180.0,
        height=55.0
    )        


# 跟新读取感应/需求系数页面
def on_select_dataframe_click(is_reaction):
    print("读取感应系数中...") if is_reaction else print("读取需求系数中...")

    global dataframe, is_reaction_global

    # 删除不需要元素
    if select_reaction_button is not None: select_reaction_button.destroy()
    if select_demand_button is not None: select_demand_button.destroy()

    excel_path = input_file_path + "/分析结果/标准化.xlsx"
    
    dataframe = pd.read_excel(excel_path, sheet_name=0, header=None) if is_reaction \
                else pd.read_excel(excel_path, sheet_name=1, header=None)
    
    is_reaction_global = is_reaction

    on_back_industry_button_click()


# 跟新省份/行业选择页面
def on_back_industry_button_click():

    if dataframe is not None:
        global industry_button, province_button, back_coefficient_button

        if listbox is not None: listbox.destroy()
        if data_button is not None: data_button.destroy()
        if load_button is not None: load_button.destroy()
        
        if industry_button is not None: industry_button.destroy()
        if province_button is not None: province_button.destroy()
        if back_industry_button is not None: back_industry_button.destroy()
        if back_coefficient_button is not None: back_coefficient_button.destroy()
        if back_start_button is not None: back_start_button.destroy()

        # 添加“产业系数”按钮
        industry_button_image = tk.PhotoImage(
        file=relative_to_assets("button_5.png")
        )
        industry_button = tk.Button(
            image=industry_button_image,
            borderwidth=0,
            highlightthickness=0,
            command=on_industry_button_click,
            relief="flat"
        )
        industry_button.image = industry_button_image
        industry_button.place(
            x=484.9999999999999,
            y=76.0,
            width=180.0,
            height=55.0
        )

        # 添加“省份系数”按钮
        province_button_image = tk.PhotoImage(
            file=relative_to_assets("button_6.png")
        )
        province_button = tk.Button(
            image=province_button_image,
            borderwidth=0,
            highlightthickness=0,
            command=on_province_button_click,
            relief="flat"
        )
        province_button.image = province_button_image
        province_button.place(
            x=484.9999999999999,
            y=173.0,
            width=180.0,
            height=55.0
        )

        # 添加“返回”按钮
        back_coefficient_button_image = tk.PhotoImage(
            file=relative_to_assets("button_7.png")
        )
        back_coefficient_button = tk.Button(
            image=back_coefficient_button_image,
            borderwidth=0,
            highlightthickness=0,
            command=on_back_coefficient_button_click,
            relief="flat"
        )
        back_coefficient_button.image = back_coefficient_button_image
        back_coefficient_button.place(
            x=484.9999999999999,
            y=270.0,
            width=180.0,
            height=55.0
        )        


# 更新"不同省份同一产业"页面
def on_industry_button_click():
    print("已选择：不同省份同一行业")

    global listbox, data_button, back_industry_button

    if back_industry_button is not None: back_industry_button.destroy()

    listbox = tk.Listbox(
        root,
        selectmode=tk.MULTIPLE,
        bd=0,
        background="#FCFCFC",
        fg="black",
        selectbackground="#3A7FF6"
    )
    for key in industry_dict.keys():
        listbox.insert(tk.END, str(key) + ' ' + industry_dict[key])
    listbox.place(
        x=484.9999999999999,
        y=76.0
    )

    # 添加“生成表格”按钮
    data_button_image = tk.PhotoImage(
        file=relative_to_assets("button_8.png")
    )
    data_button = tk.Button(
        image=data_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_industry_data_button_click,
        relief="flat"
    )
    data_button.image = data_button_image
    data_button.place(
        x=484.9999999999999,
        y=10.0,
        width=180.0,
        height=55.0
    )

    # 添加“返回”按钮
    back_industry_button_image = tk.PhotoImage(
        file=relative_to_assets("button_7.png")
    )
    back_industry_button = tk.Button(
        image=back_industry_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_back_industry_button_click,
        relief="flat"
    )
    back_industry_button.image = back_industry_button_image
    back_industry_button.place(
        x=484.9999999999999,
        y=270.0,
        width=180.0,
        height=55.0
    )


# 更新"不同行业同一省份"页面
def on_province_button_click():
    print("已选择：不同行业同一省份")

    global listbox, data_button, back_industry_button

    if back_industry_button is not None: back_industry_button.destroy()

    listbox = tk.Listbox(
        root,
        selectmode=tk.MULTIPLE,
        bd=0,
        background="#FCFCFC",
        fg="black",
        selectbackground="#3A7FF6"
    )
    for key in province_dict.keys():
        listbox.insert(tk.END, str(key) + ' ' + province_dict[key])
    listbox.place(
        x=484.9999999999999,
        y=76.0
    )

    # 添加“生成表格”按钮
    data_button_image = tk.PhotoImage(
        file=relative_to_assets("button_8.png")
    )
    data_button = tk.Button(
        image=data_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_province_data_button_click,
        relief="flat"
    )
    data_button.image = data_button_image
    data_button.place(
        x=484.9999999999999,
        y=10.0,
        width=180.0,
        height=55.0
    )

    # 添加“返回”按钮
    back_industry_button_image = tk.PhotoImage(
        file=relative_to_assets("button_7.png")
    )
    back_industry_button = tk.Button(
        image=back_industry_button_image,
        borderwidth=0,
        highlightthickness=0,
        command=on_back_industry_button_click,
        relief="flat"
    )
    back_industry_button.image = back_industry_button_image
    back_industry_button.place(
        x=484.9999999999999,
        y=270.0,
        width=180.0,
        height=55.0
    )


# 生成产业原数据表格
def on_industry_data_button_click():
    global selected_industries
    selected_indices = listbox.curselection()
    selected_industries = [int(listbox.get(idx).split(" ")[0]) for idx in selected_indices]
    print(f"产业选择成功：{selected_industries}")
    plot_industry_charts(dataframe, selected_industries, industry_dict, province_dict, is_reaction_global)


# 生成省份原数据表格
def on_province_data_button_click():
    global selected_provinces
    selected_indices = listbox.curselection()
    selected_provinces = [int(listbox.get(idx).split(" ")[0]) for idx in selected_indices]
    print(f"省份选择成功：{selected_provinces}")
    plot_province_charts(dataframe, selected_provinces, industry_dict, province_dict, is_reaction_global)

def main():
    global root
    global canvas, icon_button, text_2
    root = tk.Tk()
    root.title("投入产出表标准化数据分析")
    root.geometry("800x400")
    root.configure(bg = "#3A7FF6")

    canvas = tk.Canvas(
        root,
        bg = "#3A7FF6",
        height = 400,
        width = 800,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )

    canvas.place(x = 0, y = 0)
    canvas.create_rectangle(
        349.9999999999999,
        7.105427357601002e-15,
        799.9999999999999,
        400.0,
        fill="#FCFCFC",
        outline="")
    
    on_back_start_button_click()

    canvas.create_rectangle(
        39.999999999999886,
        160.0,
        99.99999999999989,
        165.0,
        fill="#FCFCFC",
        outline="")

    text_2 = \
    canvas.create_text(
        39.999999999999886,
        191.0,
        anchor="nw",
        text="中国（深圳）综合开发研究院",
        fill="#FCFCFC",
        font=("RobotoRoman CondensedRegular", 16 * -1)
    )
    root.resizable(False, False)
    root.mainloop()

if __name__ == "__main__":
    main()
