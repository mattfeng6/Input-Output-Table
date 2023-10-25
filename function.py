import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

current_directory = os.path.dirname(os.path.abspath(__file__))
font_path = current_directory + '/font.otf'
font_prop = FontProperties(fname=font_path)

"""
注意事项：
    1. 请跑GUI.py程序而非该程序
    2. 该程序使用Matplotlib进行数据可视化处理
    3. 请确认语言包'font.otf'是否在同一文件夹内 以显示中文语言
    4. 本程序默认各投入产出表产业列一致 如出现错误请确认对应产业信息无误
"""

# 生成产业数据柱状图
def plot_industry_charts(dataframe, industries, industry_dict, province_dict, is_reaction):

    coefficient = "感应系数" if is_reaction else "需求系数"

    def plot_industry_chart(industry):

        print(f"生成产业{coefficient}标准化表格：{industry_dict[industry]}")

        fig, ax = plt.subplots(figsize=(12, 6))

        # 根据行业找到不同省份数据
        arr = []
        for col in dataframe.columns:
            if col == 0: continue
            arr.append(dataframe.at[industry, col])

        # 图表样式设计
        ax.bar(np.arange(len(arr)), arr)
        ax.set_xticks(range(len(arr)))
        ax.set_xticklabels(province_dict.values(), rotation=90, fontproperties=font_prop)

        for i, v in enumerate(arr):
            v = round(v, 2)
            ax.text(i, v, str(v), ha='center', va='bottom')

        ax.set_xlabel('省份', fontproperties=font_prop)
        ax.set_ylabel(f'{coefficient}标准化系数', fontproperties=font_prop)
        ax.set_title(f'{industry_dict[industry]}', fontproperties=font_prop)

    for industry in industries:
        plot_industry_chart(industry)

    plt.show()

# 生成省份数据柱状图
def plot_province_charts(dataframe, provinces, industry_dict, province_dict, is_reaction):

    coefficient = "感应系数" if is_reaction else "需求系数"

    def plot_province_chart(province):
        print(f"生成省份{coefficient}标准化表格：{province_dict[province]}")

        fig, ax = plt.subplots(figsize=(12, 6))

        # 根据省份找到不同行业数据
        arr = []
        for index_row, _ in dataframe.iterrows():
            if index_row == 0: continue
            arr.append(dataframe.at[index_row, province])

        # 图表样式设计
        ax.bar(np.arange(len(arr)), arr)
        ax.set_xticks(range(len(arr)))
        ax.set_xticklabels(industry_dict.values(), rotation=90, fontproperties=font_prop)

        for i, v in enumerate(arr):
            v = round(v, 2)
            ax.text(i, v, str(v), ha='center', va='bottom')

        ax.set_xlabel('行业', fontproperties=font_prop)
        ax.set_ylabel(f'{coefficient}标准化系数', fontproperties=font_prop)
        ax.set_title(f'{province_dict[province]}', fontproperties=font_prop)

    for province in provinces:
        plot_province_chart(province)

    plt.show()