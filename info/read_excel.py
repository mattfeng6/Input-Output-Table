import pandas as pd
import numpy as np

dataframe = pd.read_excel("/Users/zhennanfeng/Downloads/internship/投入产出表/运算结果/广东.xlsx", sheet_name="42部门", header=None)
region = dataframe.iloc[10:52, 3:45]
matrix = np.array(region)

print("--------------------")
print("读取excel资料:")
print(matrix)


# 另一种读取excel表格的方法

# from openpyxl import load_workbook

# workbook = load_workbook("/Users/zhennanfeng/Downloads/internship/投入产出表/运算结果/广东.xlsx")
# sheet = workbook["42部门"]
# value = sheet["B11"].value
# print(value)