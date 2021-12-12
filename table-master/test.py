import xlrd
import numpy as np
import pandas as pd

# 提取自评成绩
zlist = []
for i in range(1, 40):
    path = '评分表\\' + str(i) + '.xlsx'
    df = pd.read_excel(path)
    zlist.append(df.iloc[i + 1, 2:8])
z = np.array(zlist)
z = z * 0.5

hlist = []
data = []
h = np.zeros([39, 6])
# 先将每个表中的自评成绩归0
for i in range(1, 40):
    df = pd.read_excel('评分表\\' + str(i) + '.xlsx')
    df.iloc[i + 1, 2:8] = 0
    h_array = np.array(df.iloc[2:, 2:8])
    h = h + h_array
h = h * 0.5 / 39

zh = np.zeros([39, 6])
zh = z + h
df = pd.DataFrame(zh)
df.to_excel("评议表总分.xlsx", index=False)
