import os
import pandas as pd
import numpy as np

input_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/2017年各省份投入产出表" # 输入路径 请替换原有路径
output_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/分析结果" # 输出路径 请替换原有路径

files = os.listdir(input_file_path)
excel_files = [file for file in files if file.endswith(('.xlsx', '.xls'))]
file = excel_files[0]
file_path = os.path.join(input_file_path, file)

dataframe = pd.read_excel(file_path, sheet_name=0, header=None, engine="xlrd" if file.endswith(('.xls')) else "openpyxl")

dict_input = dict()

def findCellIndex(value: list) -> list:

        res = []
        for index_row, _ in dataframe.iterrows():
            for col in dataframe.columns:
                cell_value = dataframe.at[index_row, col]
                if cell_value in value:
                    res.append([index_row, col])
                if len(res) == 2:
                    break
        if not 0 < len(res) <= 2:
            raise ValueError(f"错误：{file} 无法定位表格行列位置！") # 找不到对应参数 无法对应表格位置
        return res

daima_cells = findCellIndex(["代码"])

input_row_idx = daima_cells[1][0] + 1
input_col_idx = daima_cells[0][1]

for index_row, _ in dataframe.iterrows():
    if index_row < input_row_idx: continue

    daima_input = dataframe.at[index_row, input_col_idx]
    if daima_input == "TII": break
    dict_input[daima_input] = dataframe.at[index_row, input_col_idx-1]

print(dict_input)