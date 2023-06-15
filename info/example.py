# https://www.bilibili.com/video/BV1mG4y1z7Yb/?spm_id_from=333.999.0.0&vd_source=31148325763a0f09ae8963fc28f5a71b
# https://peter-puszko.medium.com/solving-leontiefs-input-output-model-in-python-a0a29455b2d8
# https://www.oecd.org/sti/ind/input-outputtables.htm

import numpy as np

# 产品部门联系表
intermediate_demand = np.array([[300, 400, 100],
                                [300, 1000, 350],
                                [200, 400, 200]])

# 供给表
primary_input = np.array([[80, 120, 120],
                          [450, 350, 380],
                          [170, 330, 300],
                          [100, 200, 350]])

# 使用表
final_demand = np.array([[500, 200, 100],
                         [600, 400, 150],
                         [500, 400, 100]])

print("--------------------")
print("产品部门联系表:")
print(intermediate_demand)
print("--------------------")
print("供给表:")
print(primary_input)
print("--------------------")
print("使用表:")
print(final_demand)

# 总投入
total_input = np.add(np.sum(intermediate_demand, axis=0), np.sum(primary_input, axis=0))

# 总产出
total_output = np.add(np.sum(intermediate_demand, axis=1), np.sum(final_demand, axis=1))

print("--------------------")
print("总投入:")
print(total_input)
print("--------------------")
print("总产出:")
print(total_output)

# 感应系数
multiplier_coefficient = np.round(np.divide(intermediate_demand, total_output), decimals=3)

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
demand_coefficient = np.round(np.divide(intermediate_demand, total_input), decimals=3)

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