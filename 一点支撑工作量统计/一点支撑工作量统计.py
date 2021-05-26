import pandas as pd
import numpy as np
import re
import time

start_time = time.time()
print('开始读取明细表...')
Tj = pd.read_excel('5月一点支撑接单量统计明细表0501-0509.xlsx', sheet_name=0)
print('正在计算...')
# 计算11位数的号码数量
Wt = Tj[['问题描述']]
Wt = pd.DataFrame(Wt)
Wt['条数'] = Wt['问题描述'].apply(lambda x: len(re.findall(r'\d{11}\d|\d{10}\d', str(x))))
# 计算带有检测字符的工单
DaTax1 = Wt[Wt.条数 == 0]
DaTax1 = pd.DataFrame(DaTax1)
DaTax1['条数'] = DaTax1['问题描述'].apply(lambda x: len(re.findall(r'检测', str(x))))
DaTax1.loc[(DaTax1['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTax1, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算工单编号数量
DaTa = DaTax1[DaTax1.条数 == 0]
DaTa = pd.DataFrame(DaTa)
# kk = re.compile(r'\w*[A-Z]\w*')
DaTa['条数'] = DaTa['问题描述'].apply(lambda x: len(re.findall(r'\w*[A-Z]\D', str(x))))
Wt = pd.merge(Wt, DaTa, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)
# 计算带有转字符的工单
DaTa1 = DaTa[DaTa.条数 == 0]
DaTa1 = pd.DataFrame(DaTa1)
DaTa1['条数'] = DaTa1['问题描述'].apply(lambda x: len(re.findall(r'转',str(x))))
DaTa1.loc[(DaTa1['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa1, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算带有派字符的工单
DaTa2 = DaTa1[DaTa1.条数 == 0]
DaTa2 = pd.DataFrame(DaTa2)
DaTa2['条数'] = DaTa2['问题描述'].apply(lambda x: len(re.findall(r'派', str(x))))
DaTa2.loc[(DaTa2['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa2, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算带有回字符的工单
DaTa3 = DaTa2[DaTa2.条数 == 0]
DaTa3 = pd.DataFrame(DaTa3)
DaTa3['条数'] = DaTa3['问题描述'].apply(lambda x: len(re.findall(r'回', str(x))))
DaTa3.loc[(DaTa3['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa3, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算带有通过字符的工单
DaTa4 = DaTa3[DaTa3.条数 == 0]
DaTa4 = pd.DataFrame(DaTa4)
DaTa4['条数'] = DaTa4['问题描述'].apply(lambda x: len(re.findall(r'通过', str(x))))
DaTa4.loc[(DaTa4['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa4, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算带有调字符的工单
DaTa5 = DaTa4[DaTa4.条数 == 0]
DaTa5 = pd.DataFrame(DaTa5)
DaTa5['条数'] = DaTa5['问题描述'].apply(lambda x: len(re.findall(r'调', str(x))))
DaTa5.loc[(DaTa5['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa5, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

# 计算带有点字符的工单
DaTa6 = DaTa5[DaTa5.条数 == 0]
DaTa6 = pd.DataFrame(DaTa6)
DaTa6['条数'] = DaTa6['问题描述'].apply(lambda x: len(re.findall(r'点', str(x))))
DaTa6.loc[(DaTa6['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa6, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)
# 计算带有质检字符的工单
DaTa7 = DaTa6[DaTa6.条数 == 0]
DaTa7 = pd.DataFrame(DaTa7)
DaTa7['条数'] = DaTa7['问题描述'].apply(lambda x: len(re.findall(r'质检', str(x))))
DaTa7.loc[(DaTa7['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa7, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)
# 计算带有质检字符的工单
DaTa8 = DaTa7[DaTa7.条数 == 0]
DaTa8 = pd.DataFrame(DaTa8)
DaTa8['条数'] = DaTa8['问题描述'].apply(lambda x: len(re.findall(r'激活', str(x))))
DaTa8.loc[(DaTa8['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa8, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)
# 计算带有预约字符的工单
DaTa9 = DaTa8[DaTa8.条数 == 0]
DaTa9 = pd.DataFrame(DaTa9)
DaTa9['条数'] = DaTa9['问题描述'].apply(lambda x: len(re.findall(r'预约', str(x))))
DaTa9.loc[(DaTa9['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa9, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)
# 计算带有首响字符的工单
DaTa10 = DaTa9[DaTa9.条数 == 0]
DaTa10 = pd.DataFrame(DaTa10)
DaTa10['条数'] = DaTa10['问题描述'].apply(lambda x: len(re.findall(r'首响', str(x))))
DaTa10.loc[(DaTa10['条数'] >= 2), '条数'] = 1
Wt = pd.merge(Wt, DaTa10, on='问题描述', how='left', right_index=True, left_index=True)
Wt = Wt.fillna(0)
Wt['条数'] = Wt.条数_x + Wt.条数_y
Wt = Wt.drop(['条数_x'], axis=1)
Wt = Wt.drop(['条数_y'], axis=1)

print('计算完毕,正在写入数据...')
with pd.ExcelWriter('一点支撑统计' + '.xlsx') as writer:  # 写入结果为当前路径
    Wt.to_excel(writer, sheet_name='一点支撑统计', startcol=0, index=False, header=True)
    DaTa1.to_excel(writer, sheet_name='1', startcol=0, index=False, header=True)
    # DaTa3.to_excel(writer, sheet_name='1', startcol=0, index=False, header=True)
end_time = time.time()
print('处理完毕!!!总耗时%0.0f秒钟' % (end_time - start_time))
