import os
import pandas as pd
import numpy as np

"""
注意事项：
    1. 该程序支持以下五方面数据分析计算
        I.  感应系数
        II. 需求系数
        III.增加值率
        IV. 劳动者报酬
        V.  感应/需求系数标准化
    2. 请确认每个Excel表格中的投入产出表在第一个工作表
    3. 目前仅支持'.xlsx'和'.xls'后缀的表格文件读取
    4. 表格文件会以'.xlsx'后缀形式导出
    5. 确保本地有numpy, pandas, openpyxl库以读取表格
        e.g. 在本地终端输入'pip install numpy pandas openpyxl'以下载对应库
"""


# input_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/2017年各省份投入产出表" # 输入路径 请替换原有路径
# output_file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/分析结果" # 输出路径 请替换原有路径


def table_output(input_file_path, output_file_path):

    ### ------------------------------------------------- 导入 & 导出文件整理 -------------------------------------------------
    ### --------------------------------------------------------------------------------------------------------------------

    """
    导入 & 导出文件整理：
        1. 读取当前文件夹内的所有Excel文件
        2. 目前仅支持'.xlsx'和'.xls'后缀的表格文件
    """

    # 如输入路径文件夹不存在 则系统报错
    if not os.path.exists(input_file_path):
        raise OSError(f"输入路径不存在：{input_file_path}")

    # 列出所有当前文件夹内的Excel文件
    files = os.listdir(input_file_path)
    excel_files = [file for file in files if file.endswith(('.xlsx', '.xls'))]
    excel_files = sorted(excel_files)

    # 如果输出路径文件夹不存在 则新建文件夹
    if not os.path.exists(output_file_path):
        os.mkdir(output_file_path)

    increment_dict = {}
    labor_dict = {}

    standardized_reaction = {}
    standardized_demand = {}

    # 生成产业和省份信息
    industry_dict = {}
    province_dict = {}

    ### ---------------------------------------------------- 读取表格信息 ----------------------------------------------------
    ### --------------------------------------------------------------------------------------------------------------------

    province_dict_idx = 1

    for file in excel_files:
        print(f"读取文档：{file}") # 读取文档进度显示

        file_name = os.path.splitext(file)[0]

        file_arr = str(file).split(".")
        province_dict[province_dict_idx] = file_arr[0]
        province_dict_idx += 1

        file_path = os.path.join(input_file_path, file)

        # 读取'.xls'用到'xlrd'库
        # 读取'.xlsx'用到'openpyxl'库
        dataframe = pd.read_excel(file_path, sheet_name=0, header=None, engine="xlrd" if file.endswith(('.xls')) else "openpyxl")

        ## -------------------------------------------------- 定位对应单元格 --------------------------------------------------
        ## ------------------------------------------------------------------------------------------------------------------

        """
        定位对应单元格：
            1. 单元格定位基于投入产出表表格基本信息 凭特定单元格确定各表位置及大小
                I.  产品部门联系表 -- 根据两个"代码"单元格确定左上位置 "TII"&"TIU"单元格确定右下位置
                II. 供给表 -- 根据"TII"单元格确定左上位置 "TVA"&"TIU"单元格确定右下位置
                III.产出表 -- 根据"GO"单元格确定对应列

            2. 如当前文档读取失败 则跳过当前文档进入下一文档读取
        """

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
                raise ValueError(f"错误：{file} 无法定位表格行列位置！") # 找不到对应参数 无法对应表格位置
            return res
        
        ## --------------------------------------------------- 读取产业信息 ---------------------------------------------------
        ## ------------------------------------------------------------------------------------------------------------------

        # 存储对应投入产出表的产业
        def findIndustryDict():
            
            # 根据"代码"单元格位置定位各产业
            daima_cells = findCellIndex(["代码"])
            row_idx = daima_cells[1][0] + 1
            col_idx = daima_cells[0][1]

            for index_row, _ in dataframe.iterrows():
                if index_row < row_idx: continue
                industry_key = dataframe.at[index_row, col_idx]
                if industry_key == "TII": break
                industry_dict[int(industry_key)] = dataframe.at[index_row, col_idx-1]

        if not industry_dict:
            findIndustryDict()
        
        # 找到产品部门联系表对应起始和结束行列
        def industryMatrixIndex() -> list:

            # 根据"代码"单元格位置定位左上数据单元格
            daima_cells = findCellIndex(["代码"])
            start_index = [daima_cells[1][0] + 1, daima_cells[0][1] + 1]

            # 根据中间投入/使用合计单元格位置定位右下单元格
            heji_cells = findCellIndex(["TII", "TIU"]) # 中间投入合计 & 中间使用合计
            end_index = [heji_cells[1][0], heji_cells[0][1]]

            return [start_index, end_index]

        # 找到供给表对应起始和结束行列
        def supplyTableIndex() -> list:
            
            # 根据"TII"单元格位置定位左上数据单元格
            tii_cell = findCellIndex(["TII"]) # 中间投入合计
            start_index = [tii_cell[0][0] + 1, tii_cell[0][1] + 1]

            # 根据"TVA" / "TIU"单元格位置定位右下单元格
            heji_cells = findCellIndex(["TVA", "TIU"]) # 增加值合计 & 中间使用合计
            end_index = [heji_cells[1][0], heji_cells[0][1]]

            return [start_index, end_index]

        # 找到产出表对应起始和结束行列
        def useTableIndex() -> list:

            # 根据"GO"单元格位置定位列
            use_cells = findCellIndex(["GO", "TII"]) # 总产出 & 中间投入合计
            start_index = [use_cells[0][0] + 1, use_cells[0][1]]
            end_index = [use_cells[1][0], use_cells[0][1]]
            return [start_index, end_index]

        # 列出对应表格的起始以及结束单元格位置
        try:
            matrix_start, matrix_end = industryMatrixIndex()
            supply_start, supply_end = supplyTableIndex()
            use_start, use_end = useTableIndex()
        # 如出现无法定位表格位置时会保错 跳过当前文档进入下一文档读取
        except ValueError:
            continue

        ## ---------------------------------------------------- 对应表格定位 ---------------------------------------------------
        ## -------------------------------------------------------------------------------------------------------------------

        # 产品部门联系表
        product_industry_matrix = np.array(dataframe.iloc[matrix_start[0]:matrix_end[0], matrix_start[1]:matrix_end[1]])

        # 供给表
        supply_table = np.array(dataframe.iloc[supply_start[0]:supply_end[0], supply_start[1]:supply_end[1]])

        # 使用表
        temp = np.array(dataframe.iloc[use_start[0]:use_end[0], use_start[1]])
        use_table = np.reshape(temp, (len(temp), 1))

        ## ------------------------------------------------- 投入产出表系数计算 -------------------------------------------------
        ## -------------------------------------------------------------------------------------------------------------------

        """
        投入产出表系数计算：
            1. 该程序支持以下五方面数据分析计算
                I.  感应系数
                II. 需求系数
                III.增加值率
                IV. 劳动者报酬
                V.  感应/需求系数标准化
            2. 感应系数矩阵计算除法进行特殊结果处理 以避免除零情况
        """

        # 总投入
        total_input = np.add(np.sum(product_industry_matrix, axis=0), np.sum(supply_table, axis=0))

        # 总产出
        total_output = use_table

        # --------------------------- 感应系数 ---------------------------
        # ---------------------------------------------------------------
        reaction_coefficient = divide_matrices(product_industry_matrix, total_output)

        # ------------------------- 感应系数汇总 -------------------------
        # ---------------------------------------------------------------
        I = np.eye(reaction_coefficient.shape[0])
        I_M = I - reaction_coefficient
        sum_reaction_coefficient = np.linalg.inv(I_M.astype(float))

        # --------------------------- 需求系数 ---------------------------
        # ---------------------------------------------------------------
        demand_coefficient = divide_matrices(product_industry_matrix, total_input)

        # ------------------------- 需求系数汇总 -------------------------
        # ---------------------------------------------------------------
        I = np.eye(demand_coefficient.shape[0])
        I_M = I - demand_coefficient
        sum_demand_coefficient = np.linalg.inv(I_M.astype(float))

        # ------------------------- 增加值率 ---------------------------
        # -------------------------------------------------------------
        tva = np.sum(supply_table, axis=0)
        increment = divide_matrices(tva, total_input)
        increment_dict[file_name] = increment

        # ------------------------ 劳动者报酬 --------------------------
        # -------------------------------------------------------------
        labor = divide_matrices(supply_table[0], tva)
        labor_dict[file_name] = labor

        # -------------------------- 标准化 ----------------------------
        # -------------------------------------------------------------

        # 感应系数标准化
        reaction_coefficient_sum = np.sum(sum_reaction_coefficient, axis=1)
        reaction_coefficient_avg = np.average(reaction_coefficient_sum)
        standardized_reaction_coefficient = divide_matrices(reaction_coefficient_sum, reaction_coefficient_avg)
        
        standardized_reaction[file_name] = standardized_reaction_coefficient
        
        # 需求系数标准化
        demand_coefficient_sum = np.sum(sum_demand_coefficient, axis=0)
        demand_coefficient_avg = np.average(demand_coefficient_sum)
        standardized_demand_coefficient = divide_matrices(demand_coefficient_sum, demand_coefficient_avg)
        
        standardized_demand[file_name] = standardized_demand_coefficient

        ## --------------------------------------------------- 对应表格导出 ----------------------------------------------------
        ## -------------------------------------------------------------------------------------------------------------------

        """
        对应表格导出：
            1. 输出表格文件会转化为".xlsx"文件并储存在输出路径文件夹内
            2. 输出表格会有多个工作表以存储对应系数
            3. 调整输出表格系数指数使代码指数对应原投入产出表
        """

        dataframe_reaction_coefficient = pd.DataFrame(reaction_coefficient)
        dataframe_sum_reaction_coefficient = pd.DataFrame(sum_reaction_coefficient)
        dataframe_demand_coefficient = pd.DataFrame(demand_coefficient)
        dataframe_sum_demand_coefficient = pd.DataFrame(sum_demand_coefficient)

        # 如原文档为".xls"表格 输出文档会转化为".xlsx"表格
        if file.endswith('.xls'):
            file += 'x'

        test_file_path = os.path.join(output_file_path, f'{file}')
        with pd.ExcelWriter(test_file_path, engine="openpyxl") as writer:

            # 调整系数使得输出代码指数从1开始
            def modifyDFIndex(dataframe: pd.DataFrame):

                custom_index = [f"{i+1}-{industry_dict[i+1]}" for i in range(len(dataframe))]

                dataframe.index = custom_index
                dataframe.columns = custom_index

            modifyDFIndex(dataframe_reaction_coefficient)
            modifyDFIndex(dataframe_sum_reaction_coefficient)
            modifyDFIndex(dataframe_demand_coefficient)
            modifyDFIndex(dataframe_sum_demand_coefficient)

            # 存储相对应系数在该表格不同工作表中
            dataframe_reaction_coefficient.to_excel(writer, sheet_name="感应系数")
            dataframe_sum_reaction_coefficient.to_excel(writer, sheet_name="感应系数汇总")
            dataframe_demand_coefficient.to_excel(writer, sheet_name="需求系数")
            dataframe_sum_demand_coefficient.to_excel(writer, sheet_name="需求系数汇总")

    ### ---------------------------------------------------- 标准化数据产出 ---------------------------------------------------
    ### ---------------------------------------------------------------------------------------------------------------------

    # 重新排列城市顺序
    standardized_reaction = dict(sorted(standardized_reaction.items()))
    standardized_demand = dict(sorted(standardized_demand.items()))

    print(f"生成增加值率文档...")
    dataframe_increment = pd.DataFrame(increment_dict)

    # 生成增加值率文档
    increment_file_path = os.path.join(output_file_path, "增加值率.xlsx")
    with pd.ExcelWriter(increment_file_path, engine="openpyxl") as writer:
        customize_row_index(dataframe_increment, industry_dict)
        dataframe_increment.to_excel(writer, sheet_name="增加值率")

    print(f"生成劳动者报酬文档...")
    dataframe_labor = pd.DataFrame(labor_dict)

    # 生成劳动者报酬文档
    labor_file_path = os.path.join(output_file_path, "劳动者报酬.xlsx")
    with pd.ExcelWriter(labor_file_path, engine="openpyxl") as writer:
        customize_row_index(dataframe_labor, industry_dict)
        dataframe_labor.to_excel(writer, sheet_name="劳动者报酬")


    print(f"生成标准化文档...")
    dataframe_standardized_reaction = pd.DataFrame(standardized_reaction)
    dataframe_standardized_demand = pd.DataFrame(standardized_demand)

    # 生成标准化文档以存储所有城市对应系数
    standardized_output_file_paths = os.path.join(output_file_path, "标准化.xlsx")
    with pd.ExcelWriter(standardized_output_file_paths, engine="openpyxl") as writer:

        customize_row_index(dataframe_standardized_reaction, industry_dict)
        customize_row_index(dataframe_standardized_demand, industry_dict)

        dataframe_standardized_reaction.to_excel(writer, sheet_name="感应系数标准化汇总")
        dataframe_standardized_demand.to_excel(writer, sheet_name="需求系数标准化汇总")

    # 整理修改省份文件名称
    customize_province_dict(province_dict)

    return industry_dict, province_dict

# 矩阵除法除零做特殊结果处理
def divide_matrices(matrix1, matrix2):
    return np.divide(matrix1, matrix2, out=np.zeros_like(matrix1), where=matrix2!= 0)

def customize_row_index(dataframe, industry_dict):
    custom_index = [f"{i+1}-{industry_dict[i+1]}" for i in range(len(dataframe))]
    dataframe.index = custom_index

# 修改省份文件名称
def customize_province_dict(province_dict):
    # 整理省份名字：删除省份"-"前缀
    for key, value in province_dict.items():
        if province_dict[key].index("-") > 0:
            province_dict[key] = province_dict[key].split("-")[1]