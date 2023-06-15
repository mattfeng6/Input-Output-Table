# https://www.bilibili.com/video/BV1mG4y1z7Yb/?spm_id_from=333.999.0.0&vd_source=31148325763a0f09ae8963fc28f5a71b
# https://peter-puszko.medium.com/solving-leontiefs-input-output-model-in-python-a0a29455b2d8
# https://www.oecd.org/sti/ind/input-outputtables.htm

import numpy as np

# 产品部门联系表
product_industry_matrix = np.array([[300, 400, 100],
                                [300, 1000, 350],
                                [200, 400, 200]])

# 供给表
supply_table = np.array([[80, 120, 120],
                          [450, 350, 380],
                          [170, 330, 300],
                          [100, 200, 350]])

# 使用表
use_table = np.array([[500, 200, 100],
                         [600, 400, 150],
                         [500, 400, 100]])

print("--------------------")
print("产品部门联系表:")
print(product_industry_matrix)
print("--------------------")
print("供给表:")
print(supply_table)
print("--------------------")
print("使用表:")
print(use_table)

# 总投入
total_input = np.add(np.sum(product_industry_matrix, axis=0), np.sum(supply_table, axis=0))

# 总产出
total_output = np.add(np.sum(product_industry_matrix, axis=1), np.sum(use_table, axis=1))

print("--------------------")
print("总投入:")
print(total_input)
print("--------------------")
print("总产出:")
print(total_output)

# 感应系数
multiplier_coefficient = np.round(np.divide(product_industry_matrix, total_output), decimals=3)

print("--------------------")
print("感应系数:")
print(multiplier_coefficient)

# 感应系数汇总
I = np.eye(multiplier_coefficient.shape[0])
I_M = I - multiplier_coefficient
L = np.linalg.inv(I_M)
sum_multiplier_coefficient = np.round(L, decimals=3)

print("--------------------")
print("感应系数汇总:")
print(sum_multiplier_coefficient)

# 需求系数
demand_coefficient = np.round(np.divide(product_industry_matrix, total_input), decimals=3)

print("--------------------")
print("需求系数:")
print(demand_coefficient)

# 需求系数汇总

I = np.eye(demand_coefficient.shape[0])
I_M = I - demand_coefficient
L = np.linalg.inv(I_M)
sum_demand_coefficient = np.round(L, decimals=3)

print("--------------------")
print("需求系数汇总:")
print(sum_demand_coefficient)