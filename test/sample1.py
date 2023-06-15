import pandas as pd
import numpy as np

### 读取表格信息

dataframe = pd.read_excel("/Users/zhennanfeng/Downloads/internship/投入产出表/运算结果/广东.xlsx", sheet_name="42部门", header=None)

# 产品部门联系表
product_industry_matrix = np.array(dataframe.iloc[10:52, 3:45])

# 供给表
supply_table = np.array(dataframe.iloc[53:57, 3:45])

# 使用表（直接取总产出）
temp = np.array(dataframe.iloc[10:52, 60])
use_table = np.reshape(temp, (len(temp), 1))

### 数据分析

# 总投入
total_input = np.add(np.sum(product_industry_matrix, axis=0), np.sum(supply_table, axis=0))

# 总产出
# total_output = np.add(np.sum(product_industry_matrix, axis=1), np.sum(use_table, axis=1))
total_output = use_table

# 感应系数
multiplier_coefficient = np.divide(product_industry_matrix, total_output)

# 感应系数汇总
I = np.eye(multiplier_coefficient.shape[0])
I_M = I - multiplier_coefficient
sum_multiplier_coefficient = np.linalg.inv(I_M.astype(float))

# 需求系数
demand_coefficient = np.divide(product_industry_matrix, total_input)

# 需求系数汇总
I = np.eye(demand_coefficient.shape[0])
I_M = I - demand_coefficient
sum_demand_coefficient = np.linalg.inv(I_M.astype(float))


### 转换表格

dataframe_multiplier_coefficient = pd.DataFrame(multiplier_coefficient)
dataframe_sum_mutltiplier_coefficient = pd.DataFrame(sum_multiplier_coefficient)
dataframe_demand_coefficient = pd.DataFrame(demand_coefficient)
dataframe_sum_demand_coefficient = pd.DataFrame(sum_demand_coefficient)


### 导出表格

filepath = "/Users/zhennanfeng/Downloads/internship/投入产出表/运算结果/test.xlsx"
with pd.ExcelWriter(filepath) as writer:

    dataframe_multiplier_coefficient.to_excel(writer, sheet_name="感应系数")
    dataframe_sum_mutltiplier_coefficient.to_excel(writer, sheet_name="感应系数汇总")
    dataframe_demand_coefficient.to_excel(writer, sheet_name="需求系数")
    dataframe_sum_demand_coefficient.to_excel(writer, sheet_name="需求系数汇总")
