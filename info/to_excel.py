import pandas as pd
import numpy as np

matrix = np.array([[1,2,3], [4,5,6]])
dataframe = pd.DataFrame(matrix)
dataframe.to_excel("/Users/zhennanfeng/Downloads/internship/投入产出表/运算结果/test.xlsx")