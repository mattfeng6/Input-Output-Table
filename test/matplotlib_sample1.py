import os
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

from matplotlib.font_manager import FontProperties
font_path = '/Users/zhennanfeng/Downloads/internship/投入产出表/font.otf'
font_prop = FontProperties(fname=font_path)

# from sample1 import output_file_path
# input_file_path = os.path.join(output_file_path, "标准化.xlsx")

industry_dict = {1: '农林牧渔产品和服务',
                 2: '煤炭采选产品',
                 3: '石油和天然气开采产品',
                 4: '金属矿采选产品',
                 5: '非金属矿和其他矿采选产品',
                 6: '食品和烟草',
                 7: '纺织品',
                 8: '纺织服装鞋帽皮革羽绒及其制品',
                 9: '木材加工品和家具',
                 10: '造纸印刷和文教体育用品',
                 11: '石油、炼焦产品和核燃料加工品',
                 12: '化学产品',
                 13: '非金属矿物制品',
                 14: '金属冶炼和压延加工品',
                 15: '金属制品',
                 16: '通用设备',
                 17: '专用设备',
                 18: '交通运输设备',
                 19: '电气机械和器材',
                 20: '通信设备、计算机和其他电子设备',
                 21: '仪器仪表',
                 22: '其他制造产品和废品废料',
                 23: '金属制品、机械和设备修理服务',
                 24: '电力、热力的生产和供应',
                 25: '燃气生产和供应',
                 26: '水的生产和供应',
                 27: '建筑',
                 28: '批发和零售',
                 29: '交通运输、仓储和邮政',
                 30: '住宿和餐饮',
                 31: '信息传输、软件和信息技术服务',
                 32: '金融',
                 33: '房地产',
                 34: '租赁和商务服务',
                 35: '研究和试验发展',
                 36: '综合技术服务',
                 37: '水利、环境和公共设施管理',
                 38: '居民服务、修理和其他服务',
                 39: '教育',
                 40: '卫生和社会工作',
                 41: '文化、体育和娱乐', 
                 42: '公共管理、社会保障和社会组织'}

province_dict = {1: '01-北京',
                 2: '02-天津',
                 3: '03-河北',
                 4: '04-山西',
                 5: '05-内蒙古',
                 6: '06-辽宁',
                 7: '07-吉林',
                 8: '08-黑龙江',
                 9: '09-上海',
                 10: '10-江苏',
                 11: '11-浙江',
                 12: '12-安徽',
                 13: '13-福建',
                 14: '14-江西',
                 15: '15-山东',
                 16: '16-河南',
                 17: '17-湖北',
                 18: '18-湖南',
                 19: '19-广东',
                 20: '20-广西',
                 21: '21-海南',
                 22: '22-重庆',
                 23: '23-四川',
                 24: '24-贵州',
                 25: '25-云南',
                 26: '26-西藏',
                 27: '27-陕西',
                 28: '28-甘肃',
                 29: '29-青海',
                 30: '30-宁夏',
                 31: '31-新疆'}

file_path = "/Users/zhennanfeng/Downloads/internship/投入产出表/分析结果/标准化.xlsx" 

dataframe = pd.read_excel(file_path, sheet_name=0, header=None)

# for index_row, _ in dataframe.iterrows():
#     if index_row == 0: continue
#     dataframe.loc[index_row, 0] = industry_dict[int(index_row)]
# dataframe.to_excel(file_path)

industry_keys = industry_dict.keys()
industry_values = industry_dict.values()

# 整理省份名字：删除省份"-"前缀
for key, value in province_dict.items():
    if province_dict[key].index("-") > 0:
        province_dict[key] = province_dict[key].split("-")[1]

province_keys = province_dict.keys()
province_values = province_dict.values()


## 行业内不同省份

industries = [1, 2, 7, 9]

def plot_industry_charts(industries):

    def plot_industry_chart(industry):

        print(f"生成行业标准化表格：{industry_dict[industry]}")

        fig, ax = plt.subplots(figsize=(12, 6))

        # 根据行业找到不同省份数据
        arr = []
        for col in dataframe.columns:
            if col == 0: continue
            arr.append(dataframe.at[industry, col])

        # 图表样式设计
        ax.bar(np.arange(len(arr)), arr)
        ax.set_xticks(range(len(arr)))
        ax.set_xticklabels(province_values, fontproperties=font_prop)

        for i, v in enumerate(arr):
            v = round(v, 2)
            ax.text(i, v, str(v), ha='center', va='bottom')

        ax.set_xlabel('省份', fontproperties=font_prop)
        ax.set_ylabel('标准化系数', fontproperties=font_prop)
        ax.set_title(f'{industry_dict[industry]}', fontproperties=font_prop)

    for industry in industries:
        plot_industry_chart(industry)

    plt.show()

# plot_industry_charts(industries)

## 省内不同行业

provinces = [1, 3, 21]

def plot_province_charts(provinces):

    def plot_province_chart(province):
        print(f"生成省份标准化表格：{province_dict[province]}")

        fig, ax = plt.subplots(figsize=(16, 8))

        # 根据省份找到不同行业数据
        arr = []
        for index_row, _ in dataframe.iterrows():
            if index_row == 0: continue
            arr.append(dataframe.at[index_row, province])

        # 图表样式设计
        ax.bar(np.arange(len(arr)), arr)
        ax.set_xticks(range(len(arr)))
        ax.set_xticklabels(industry_values, rotation=90, fontproperties=font_prop)

        for i, v in enumerate(arr):
            v = round(v, 2)
            ax.text(i, v, str(v), ha='center', va='bottom')

        ax.set_xlabel('行业', fontproperties=font_prop)
        ax.set_ylabel('标准化系数', fontproperties=font_prop)
        ax.set_title(f'{province_dict[province]}', fontproperties=font_prop)

    for province in provinces:
        plot_province_chart(province)

    plt.show()
        
plot_province_charts(provinces)

