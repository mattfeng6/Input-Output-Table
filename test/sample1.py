import os
import pandas as pd
import numpy as np

"""
注意事项：
    1. 请修改输入路径及输出路径以方便读取并存储Excel表格
    2. 请确认每个Excel表格中的投入产出表在第一个工作表

"""

# 输入路径
input_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/2017年各省份投入产出表"

# 输出路径
output_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/result"

# 列出所有当前文件夹内的Excel文件
files = os.listdir(input_file_path)
excel_files = [file for file in files if file.endswith(('.xlsx', '.xls'))]

# 如果输出路径文件夹不存在，新建文件夹
if not os.path.exists(output_file_path):
    os.mkdir(output_file_path)

print(excel_files)

### 读取表格信息

for file in excel_files:
    file_path = os.path.join(input_file_path, file)
    dataframe = pd.read_excel(file_path, sheet_name=0, header=None)

    ### 定位对应单元格

    # 寻找特定内容的单元格的行列并输出所有满足该条件的单元格
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
            raise ValueError(f"错误：{file} 无法定位表格行列位置！")
        return res

    # 找到产品部门联系表对应起始和结束行列
    def industryMatrixIndex() -> list:

        # 根据“代码”单元格位置定位左上数据单元格
        daima_cells = findCellIndex(["代码"])
        start_index = [daima_cells[1][0] + 1, daima_cells[0][1] + 1]

        # 根据中间投入/使用合计单元格位置定位右下单元格
        heji_cells = findCellIndex(["TII", "TIU"]) # 中间投入合计 & 中间使用合计
        end_index = [heji_cells[1][0], heji_cells[0][1]]

        return [start_index, end_index]

    # 找到供给表对应起始和结束行列
    def supplyTableIndex() -> list:
        
        # 根据“TII”单元格位置定位左上数据单元格
        tii_cell = findCellIndex(["TII"]) # 中间投入合计
        start_index = [tii_cell[0][0] + 1, tii_cell[0][1] + 1]

        # 根据“TVA” / “TIU”单元格位置定位右下单元格
        heji_cells = findCellIndex(["TVA", "TIU"]) # 增加值合计 & 中间使用合计
        end_index = [heji_cells[1][0], heji_cells[0][1]]

        return [start_index, end_index]

    # 找到产出表对应起始和结束行列
    def useTableIndex() -> list:

        # 根据“GO”单元格位置定位列
        use_cells = findCellIndex(["GO", "TII"]) # 总产出 & 中间投入合计
        start_index = [use_cells[0][0] + 1, use_cells[0][1]]
        end_index = [use_cells[1][0], use_cells[0][1]]
        return [start_index, end_index]

    # 列出对应表格的起始以及结束单元格位置
    try:
        matrix_start, matrix_end = industryMatrixIndex()
        supply_start, supply_end = supplyTableIndex()
        use_start, use_end = useTableIndex()
    except ValueError:
        continue

    # 产品部门联系表
    product_industry_matrix = np.array(dataframe.iloc[matrix_start[0]:matrix_end[0], matrix_start[1]:matrix_end[1]])

    # 供给表
    supply_table = np.array(dataframe.iloc[supply_start[0]:supply_end[0], supply_start[1]:supply_end[1]])

    # 使用表（直接取总产出）
    temp = np.array(dataframe.iloc[use_start[0]:use_end[0], use_start[1]])
    use_table = np.reshape(temp, (len(temp), 1))

    ### 数据分析

    # 总投入
    total_input = np.add(np.sum(product_industry_matrix, axis=0), np.sum(supply_table, axis=0))

    # 总产出
    # total_output = np.add(np.sum(product_industry_matrix, axis=1), np.sum(use_table, axis=1))
    total_output = use_table

    # 矩阵除法除零做特殊结果处理
    def divide_matrices(matrix1, matrix2):
        return np.divide(matrix1, matrix2, out=np.zeros_like(matrix1), where=matrix2!= 0)

    # 感应系数
    multiplier_coefficient = divide_matrices(product_industry_matrix, total_output)

    # 感应系数汇总
    I = np.eye(multiplier_coefficient.shape[0])
    I_M = I - multiplier_coefficient
    sum_multiplier_coefficient = np.linalg.inv(I_M.astype(float))

    # 需求系数
    demand_coefficient = divide_matrices(product_industry_matrix, total_input)

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

    test_file_path = os.path.join(output_file_path, f'{file}')
    with pd.ExcelWriter(test_file_path) as writer:

        # 调整系数使得输出代码指数从1开始
        def modifyDFIndex(dataframe: pd.DataFrame):
            dataframe.index += 1
            dataframe.columns = range(1, len(dataframe.columns) + 1)

        modifyDFIndex(dataframe_multiplier_coefficient)
        modifyDFIndex(dataframe_sum_mutltiplier_coefficient)
        modifyDFIndex(dataframe_demand_coefficient)
        modifyDFIndex(dataframe_sum_demand_coefficient)

        dataframe_multiplier_coefficient.to_excel(writer, sheet_name="感应系数")
        dataframe_sum_mutltiplier_coefficient.to_excel(writer, sheet_name="感应系数汇总")
        dataframe_demand_coefficient.to_excel(writer, sheet_name="需求系数")
        dataframe_sum_demand_coefficient.to_excel(writer, sheet_name="需求系数汇总")