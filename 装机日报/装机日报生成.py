# 更新日期:2021-7-21
# 更新选取时间需求,优化运行时间
import pandas as pd
import time, datetime
import re
import numpy as np
from pandas import Series
import xlwings as xw
import os

last_month = '5月'  # 用于选取几月拍照值

yestday1 = '6月30日'  # 用于选取昨日数据,这里根据表格日期更改

bf_yestday1 = '6月29日'  # 指定要删除的前天数据,这里根据表格的日期更改

start_time = time.time()
todaytime = datetime.datetime.strptime('2021-7-1', '%Y-%m-%d')  # 手动输入日期
todaytime1 = str(todaytime.month) + '月' + todaytime.strftime('%d') + '日'  # 日期带0,比如7月01日
todaytime2 = str(todaytime.month) + '月' + str(todaytime.day) + '日'  # 日期带0,比如7月1日


# 新读取表格方法
def GetDataFrame(file, sheets=0):
    if os.path.exists(file):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(file)  # 打开Excel文件

            sht_all = wb.sheets[sheets]

            info = sht_all.used_range

            nrows = info.last_cell.row

            ncols = info.last_cell.column
            index1 = sht_all.range((1, 1), (1, ncols)).value

            index2 = Series(index1)

            Data = sht_all.range((2, 1), (nrows, ncols)).value

            Data = pd.DataFrame(Data, columns=index2)
            wb.close()
            app.quit()
            return Data
        except:
            # 报错自动结束进程
            print('运行错误')
            os.system('taskkill/IM et.exe /F')
            quit()
    else:
        print('文件名不存在')
        quit()


# 批量修改表格格式
def format_str(data, name, ty=str):
    data_types_dict = {name: ty}
    data = data.astype(data_types_dict)
    return data


# 文件路径名

HISTORY_PATH = '调度中心日报0601-0630.xlsx'  # 选取历史数据的表名,在此修改即可

DATA_PATH = f'广西移动装机工单分时间段通报2021年{todaytime2}24点.xlsx'
JG_PATH = f'7月01日-{todaytime2}家宽详表（集团模板）.xlsx'
GX_PATH = f'7月1-{todaytime2}（广西）新家宽装机及时率报表.xlsx'
CONTRAST_PATH = '程序专用对照表.xlsx'  # 程序对照表名
# 读取表格

print('正在读取数据表...')
Data_Quxian = pd.read_excel(DATA_PATH, sheet_name='表2区县超长超时')
Data_Wangge = pd.read_excel(DATA_PATH, sheet_name='表3网格超长超时')
Data_all1 = pd.read_excel(DATA_PATH, sheet_name='月日报', skiprows=1,
                          skipfooter=15)
Data_all1 = Data_all1.iloc[:15, 1:21]

print('正在读取广西竣工详表...')
Data_gx = GetDataFrame(GX_PATH)
# Data_gx = pd.read_excel(GX_PATH)
print('正在读取集团竣工详表...')
# Data_jt = pd.read_excel(JG_PATH)
Data_jt = GetDataFrame(JG_PATH)

print('正在读取数据表...')
Data_all = pd.read_excel(HISTORY_PATH, sheet_name=4, skiprows=1)
QX_data = pd.read_excel(CONTRAST_PATH, sheet_name='代维区县对照表')
WG_data = pd.read_excel(CONTRAST_PATH, sheet_name='网格信息')

# 改列名
Data_gx = Data_gx[Data_gx['工单标题'].str.contains('1000')]  # 筛选千兆数据

Data_gx.rename(columns={'首次响应时长（H）': '首次响应时长'}, inplace=True)
Data_jt.rename(columns={'首次预约时长(H)': '首次预约时长'}, inplace=True)
Data_jt.rename(columns={'工单历时减去总时长分化时': '工单时长'}, inplace=True)

Data_gx = pd.DataFrame(Data_gx)
Data_gx['首次响应时长'] = Data_gx['首次响应时长'].apply(pd.to_numeric, downcast='float')
Data_gx['全流程工作时长'] = Data_gx['全流程工作时长'].apply(pd.to_numeric, downcast='float')
Data_jt['首次预约时长'] = Data_jt['首次预约时长'].apply(pd.to_numeric, downcast='float')
Data_jt['工单时长'] = Data_jt['工单时长'].apply(pd.to_numeric, downcast='float')
# format_str(Data_gx, '首次响应时长', np.float64)
# format_str(Data_jt, '首次预约时长', np.float64)
# format_str(Data_jt, '工单时长', np.float64)

WG_data.drop('备注', axis=1, inplace=True)  # 删除不需要的数据
# 匹配网格信息
# 广西详表
Data_gx['区县'].fillna('区县信息为空', inplace=True)  # 批量替换nan
Data_gx['所属网格'].fillna('网格信息为空', inplace=True)
map1 = pd.merge(Data_gx, WG_data, on=["地市", '所属网格'])
map1.rename(columns={'boss工单号': 'boss工单号1'}, inplace=True)  # 改列名
Data_gx = pd.merge(Data_gx, map1, how='left')
Data_gx['boss工单号1'].fillna('未匹配', inplace=True)  # 批量替换nan
Data_gx.loc[(Data_gx['boss工单号1'] == '未匹配'), '所属网格'] = '网格信息为空'  # 网格进行判断，得到所筛选的

# 集团详表
Data_jt['区县'].fillna('区县信息为空', inplace=True)  # 批量替换nan
Data_jt['所属网格'].fillna('网格信息为空', inplace=True)
map1 = pd.merge(Data_jt, WG_data, on=["地市", '所属网格'])
map1.rename(columns={'boss工单号': 'boss工单号1'}, inplace=True)  # 改列名
Data_jt = pd.merge(Data_jt, map1, how='left')
Data_jt['boss工单号1'].fillna('未匹配', inplace=True)  # 批量替换nan
Data_jt.loc[(Data_jt['boss工单号1'] == '未匹配'), '所属网格'] = '网格信息为空'  # 网格进行判断，得到所筛选的

# 筛选所需要的数据
Gx_drop = Data_gx[Data_gx.首次响应时长 >= 0]  # 剔除首响时长为负数的数据
Jt_drop = Data_jt[Data_jt.首次预约时长 >= 0]  # 剔除首响时长为负数的数据
# 替换表格nan为0
Data_Quxian = Data_Quxian.fillna(0)  # 批量替换nan 为数字 0
Data_Wangge = Data_Wangge.fillna(0)  # 批量替换nan 为数字 0
Quxian = pd.DataFrame(Data_Quxian)
Wangge = pd.DataFrame(Data_Wangge)
Dishi = pd.DataFrame(Data_all1)
Data_all = pd.DataFrame(Data_all)
# 提取数据
Quxian = Quxian[['地市', '区县', '未归档工单总数', '超72小时工单数', '在途工单压单比']]
Wangge = Wangge[['地市', '所属网格', '未归档工单总数', '超72小时工单数', '在途工单压单比']]
Dishi = Dishi[['地市', '家宽未归档工单总数', '超72小时工单数', '在途工单压单比']]
# 改列名

Dishi.rename(columns={'家宽未归档工单总数': '家宽未归档总数'}, inplace=True)

Df = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']})
# 计算广西区县千兆首响平均时长
Gx_sx = Gx_drop[['地市', '区县', '首次响应时长']].groupby(['地市', '区县']).mean().reset_index()
Quxian = pd.merge(Quxian, Gx_sx, on=['地市', '区县'], how='left')  # 拼接

# 计算广西区县千兆平均工作时长
Gx_time = Data_gx[['地市', '区县', '全流程工作时长']].groupby(['地市', '区县']).mean().reset_index()
Quxian = pd.merge(Quxian, Gx_time, on=['地市', '区县'], how='left')  # 拼接
# 计算广西区县普通家宽首响平均时长
Jt_sx = Jt_drop[['地市', '区县', '首次预约时长']].groupby(['地市', '区县']).mean().reset_index()
Quxian = pd.merge(Quxian, Jt_sx, on=['地市', '区县'], how='left')  # 拼接
# 计算广西区县普通家宽平均工作时长
Jt_time = Data_jt[['地市', '区县', '工单时长']].groupby(['地市', '区县']).mean().reset_index()
Quxian = pd.merge(Quxian, Jt_time, on=['地市', '区县'], how='left')  # 拼接
# 计算广西区县装机未归档超72小时工单占比
Quxian['装机未归档超72小时工单占比'] = Quxian.超72小时工单数 / Quxian.未归档工单总数
# 最后整改
Quxian = Quxian.round({'在途工单压单比': 2})  # 四舍五入
Quxian = Quxian.round({'首次响应时长': 2})  # 四舍五入
Quxian = Quxian.round({'全流程工作时长': 2})  # 四舍五入
Quxian = Quxian.round({'首次预约时长': 2})  # 四舍五入
Quxian = Quxian.round({'工单时长': 2})  # 四舍五入
Quxian = Quxian.round({'装机未归档超72小时工单占比': 4})  # 四舍五入
Quxian['装机未归档超72小时工单占比'] = Quxian['装机未归档超72小时工单占比'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
# 排序
Quxian = Quxian[['地市', '区县', '首次响应时长', '全流程工作时长', '首次预约时长', '工单时长', '未归档工单总数', '超72小时工单数', '装机未归档超72小时工单占比',
                 '在途工单压单比']]
# 改列名
Quxian.rename(columns={'首次响应时长': '千兆装移机首响'}, inplace=True)
Quxian.rename(columns={'全流程工作时长': '千兆装移机时长'}, inplace=True)
Quxian.rename(columns={'首次预约时长': '普通装移机首响'}, inplace=True)
Quxian.rename(columns={'工单时长': '普通装移机装机时长'}, inplace=True)
Quxian.rename(columns={'未归档工单总数': '家宽未归档总数'}, inplace=True)
# 计算广西网格千兆首响平均时长
Wg_sx = Gx_drop[['地市', '所属网格', '首次响应时长']].groupby(['地市', '所属网格']).mean().reset_index()
Wangge = pd.merge(Wangge, Wg_sx, on=['地市', '所属网格'], how='left')  # 拼接
# 计算广西网格千兆平均工作时长
Wg_time = Data_gx[['地市', '所属网格', '全流程工作时长']].groupby(['地市', '所属网格']).mean().reset_index()
Wangge = pd.merge(Wangge, Wg_time, on=['地市', '所属网格'], how='left')  # 拼接
# 计算广西网格普通家宽首响平均时长
Wgjt_sx = Jt_drop[['地市', '所属网格', '首次预约时长']].groupby(['地市', '所属网格']).mean().reset_index()
Wangge = pd.merge(Wangge, Wgjt_sx, on=['地市', '所属网格'], how='left')  # 拼接
# 计算广西网格普通家宽平均工作时长
Wgjt_time = Data_jt[['地市', '所属网格', '工单时长']].groupby(['地市', '所属网格']).mean().reset_index()
Wangge = pd.merge(Wangge, Wgjt_time, on=['地市', '所属网格'], how='left')  # 拼接
# 计算广西网格装机未归档超72小时工单占比
Wangge['装机未归档超72小时工单占比'] = Wangge.超72小时工单数 / Wangge.未归档工单总数
# 最后整改
Wangge = Wangge.round({'在途工单压单比': 2})  # 四舍五入
Wangge = Wangge.round({'首次响应时长': 2})  # 四舍五入
Wangge = Wangge.round({'全流程工作时长': 2})  # 四舍五入
Wangge = Wangge.round({'首次预约时长': 2})  # 四舍五入
Wangge = Wangge.round({'工单时长': 2})  # 四舍五入
Wangge = Wangge.round({'装机未归档超72小时工单占比': 4})  # 四舍五入
Wangge = Wangge.fillna(0)  # 批量替换nan 为数字 0
Wangge['装机未归档超72小时工单占比'] = Wangge['装机未归档超72小时工单占比'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
# 排序
Wangge = Wangge[['地市', '所属网格', '首次响应时长', '全流程工作时长', '首次预约时长', '工单时长', '未归档工单总数', '超72小时工单数', '装机未归档超72小时工单占比',
                 '在途工单压单比']]
# 改列名
Wangge.rename(columns={'首次响应时长': '千兆装移机首响'}, inplace=True)
Wangge.rename(columns={'全流程工作时长': '千兆装移机时长'}, inplace=True)
Wangge.rename(columns={'首次预约时长': '普通装移机首响'}, inplace=True)
Wangge.rename(columns={'工单时长': '普通装移机装机时长'}, inplace=True)
Wangge.rename(columns={'未归档工单总数': '家宽未归档总数'}, inplace=True)
# 计算广西全区千兆首响平均时长
Dishi_sx = Gx_drop[['地市', '首次响应时长']].groupby(['地市']).mean().reset_index()
Dishi_sx = Dishi_sx.append(
    [{'地市': '全区', '首次响应时长': Gx_drop[['地市', '首次响应时长']].mean().reset_index(name='首次响应时长').at[0, '首次响应时长']}],
    ignore_index=True)  # 计算全区
Dishi = pd.merge(Dishi, Dishi_sx, on=['地市'], how='left')  # 拼接
##计算广西全区千兆平均工作时长
Dishi_time = Data_gx[['地市', '全流程工作时长']].groupby(['地市']).mean().reset_index()
Dishi_time = Dishi_time.append(
    [{'地市': '全区', '全流程工作时长': Data_gx[['地市', '全流程工作时长']].mean().reset_index(name='全流程工作时长').at[0, '全流程工作时长']}],
    ignore_index=True)  # 计算全区
Dishi = pd.merge(Dishi, Dishi_time, on=['地市'], how='left')  # 拼接
# 计算广西全区普通家宽首响平均时长
Dsjt_sx = Jt_drop[['地市', '首次预约时长']].groupby(['地市']).mean().reset_index()
Dsjt_sx = Dsjt_sx.append(
    [{'地市': '全区', '首次预约时长': Jt_drop[['地市', '首次预约时长']].mean().reset_index(name='首次预约时长').at[0, '首次预约时长']}],
    ignore_index=True)  # 计算全区
Dishi = pd.merge(Dishi, Dsjt_sx, on=['地市'], how='left')  # 拼接
# 计算广西全区普通家宽平均工作时长
Dsjt_time = Data_jt[['地市', '工单时长']].groupby(['地市']).mean().reset_index()
Dsjt_time = Dsjt_time.append(
    [{'地市': '全区', '工单时长': Data_jt[['地市', '工单时长']].mean().reset_index(name='工单时长').at[0, '工单时长']}],
    ignore_index=True)  # 计算全区
Dishi = pd.merge(Dishi, Dsjt_time, on=['地市'], how='left')  # 拼接

# 计算广西全区装机未归档超72小时工单占比
Dishi['装机未归档超72小时工单占比'] = Dishi.超72小时工单数 / Dishi.家宽未归档总数
# 最后整改

Dishi = Dishi.fillna(0)  # 批量替换nan 为数字 0
# 排序
Dishi = Dishi[['地市', '首次响应时长', '全流程工作时长', '首次预约时长', '工单时长', '家宽未归档总数', '超72小时工单数', '装机未归档超72小时工单占比',
               '在途工单压单比']]

# 选择读取区域
Data_all = Data_all.iloc[:16, :31]
Data_all = Data_all.replace(r'\n', '', regex=True)  # 替换换行
Data_all.fillna('地市', inplace=True)
# #重命名表头
list0 = list(Data_all.iloc[0])
Data_all.columns = list0  ###重命名表头
Data_all = Data_all.drop(Data_all.head(1).index)
Data_all = Data_all.reset_index(drop=True)
#
Data_all.drop([bf_yestday1], axis=1, inplace=True)  # 删除前天数据
Data_all.drop('环比昨日', axis=1, inplace=True)  # 删除不需要的数据
Data_all.drop('环比上月', axis=1, inplace=True)  # 删除不需要的数据
Data_all.rename(columns={last_month + '拍照值': '拍照值'}, inplace=True)
# 批量修改列名
for i, vl in enumerate(Data_all.columns.values):
    if re.search(vl, vl):
        # 重命名列名
        column_names = Data_all.columns.values
        column_names[i] = vl + str(i)
        Data_all.columns = column_names
Data_all.rename(columns={'地市0': '地市'}, inplace=True)
Data_all.rename(columns={'拍照值2': '千兆首响拍照值'}, inplace=True)
Data_all.rename(columns={'拍照值4': '千兆时长拍照值'}, inplace=True)
Data_all.rename(columns={'拍照值6': '普通首响拍照值'}, inplace=True)
Data_all.rename(columns={'拍照值8': '普通时长拍照值'}, inplace=True)
Data_all.rename(columns={'拍照值10': '超72小时拍照值'}, inplace=True)
Data_all.rename(columns={'拍照值12': '压单比拍照值'}, inplace=True)
# 给予单独变量
today_xs1 = Dishi[['地市', '首次响应时长']]
today_xs2 = Dishi[['地市', '全流程工作时长']]
today_xs3 = Dishi[['地市', '首次预约时长']]
today_xs4 = Dishi[['地市', '工单时长']]
today_xs5 = Dishi[['地市', '装机未归档超72小时工单占比']]
today_xs6 = Dishi[['地市', '在途工单压单比']]
yestday_xs1 = Data_all[['地市', yestday1 + '1']]
yestday_xs2 = Data_all[['地市', yestday1 + '3']]
yestday_xs3 = Data_all[['地市', yestday1 + '5']]
yestday_xs4 = Data_all[['地市', yestday1 + '7']]
yestday_xs5 = Data_all[['地市', yestday1 + '9']]
yestday_xs6 = Data_all[['地市', yestday1 + '11']]
yestday_pz1 = Data_all[['地市', '千兆首响拍照值']]
yestday_pz2 = Data_all[['地市', '千兆时长拍照值']]
yestday_pz3 = Data_all[['地市', '普通首响拍照值']]
yestday_pz4 = Data_all[['地市', '普通时长拍照值']]
yestday_pz5 = Data_all[['地市', '超72小时拍照值']]
yestday_pz6 = Data_all[['地市', '压单比拍照值']]
# 千兆装移机首响
# 拼接数据，计算昨日环比
Data_daily = pd.merge(today_xs1, yestday_xs1, on='地市', how='left')
Data_daily['千兆首响环比昨日'] = (Data_daily['首次响应时长'] - Data_daily[yestday1 + '1']) / Data_daily[yestday1 + '1']
# 拼接数据，计算上月环比
Data_daily = pd.merge(Data_daily, yestday_pz1, on='地市', how='left')
Data_daily['千兆首响环比上月'] = (Data_daily['首次响应时长'] - Data_daily['千兆首响拍照值']) / Data_daily['千兆首响拍照值']
# 千兆装移机时长
# 拼接数据，计算昨日环比
Data_daily1 = pd.merge(today_xs2, yestday_xs2, on='地市', how='left')
Data_daily1['千兆时长环比昨日'] = (Data_daily1['全流程工作时长'] - Data_daily1[yestday1 + '3']) / Data_daily1[yestday1 + '3']
# 拼接数据，计算上月环比
Data_daily1 = pd.merge(Data_daily1, yestday_pz2, on='地市', how='left')
Data_daily1['千兆时长环比上月'] = (Data_daily1['全流程工作时长'] - Data_daily1['千兆时长拍照值']) / Data_daily1['千兆时长拍照值']
Data_daily = pd.merge(Data_daily, Data_daily1, on='地市', how='left')
# 普通装移机首响
# 拼接数据，计算昨日环比x
Data_daily2 = pd.merge(today_xs3, yestday_xs3, on='地市', how='left')
Data_daily2['普通首响环比昨日'] = (Data_daily2['首次预约时长'] - Data_daily2[yestday1 + '5']) / Data_daily2[yestday1 + '5']
# 拼接数据，计算上月环比
Data_daily2 = pd.merge(Data_daily2, yestday_pz3, on='地市', how='left')
Data_daily2['普通首响环比上月'] = (Data_daily2['首次预约时长'] - Data_daily2['普通首响拍照值']) / Data_daily2['普通首响拍照值']
Data_daily = pd.merge(Data_daily, Data_daily2, on='地市', how='left')
# 普通装移机装机时长
# 拼接数据，计算昨日环比
Data_daily3 = pd.merge(today_xs4, yestday_xs4, on='地市', how='left')
Data_daily3['普通时长环比昨日'] = (Data_daily3['工单时长'] - Data_daily3[yestday1 + '7']) / Data_daily3[yestday1 + '7']
# 拼接数据，计算上月环比
Data_daily3 = pd.merge(Data_daily3, yestday_pz4, on='地市', how='left')
Data_daily3['普通时长环比上月'] = (Data_daily3['工单时长'] - Data_daily3['普通时长拍照值']) / Data_daily3['普通时长拍照值']
Data_daily = pd.merge(Data_daily, Data_daily3, on='地市', how='left')
# 装机未归档超72小时工单占比
# 拼接数据，计算昨日环比


Data_daily4 = pd.merge(today_xs5, yestday_xs5, on='地市', how='left')
# Data_daily4[yestday1 + '9'] = Data_daily4[yestday1 + '9'].str.strip("%").astype(float) / 100
Data_daily4['超72小时环比昨日'] = (Data_daily4['装机未归档超72小时工单占比'] - Data_daily4[yestday1 + '9']) / Data_daily4[yestday1 + '9']
# 拼接数据，计算上月环比
Data_daily4 = pd.merge(Data_daily4, yestday_pz5, on='地市', how='left')
# Data_daily4['超72小时拍照值'] = Data_daily4['超72小时拍照值'].str.strip("%").astype(float) / 100
Data_daily4['超72小时环比上月'] = (Data_daily4['装机未归档超72小时工单占比'] - Data_daily4['超72小时拍照值']) / Data_daily4['超72小时拍照值']
Data_daily = pd.merge(Data_daily, Data_daily4, on='地市', how='left')
# 装机未归档压单比
# 拼接数据，计算昨日环比
Data_daily5 = pd.merge(today_xs6, yestday_xs6, on='地市', how='left')
Data_daily5['压单比环比昨日'] = (Data_daily5['在途工单压单比'] - Data_daily5[yestday1 + '11']) / Data_daily5[yestday1 + '11']
# 拼接数据，计算上月环比
Data_daily5 = pd.merge(Data_daily5, yestday_pz6, on='地市', how='left')
Data_daily5['压单比环比上月'] = (Data_daily5['在途工单压单比'] - Data_daily5['压单比拍照值']) / Data_daily5['压单比拍照值']
Data_daily = pd.merge(Data_daily, Data_daily5, on='地市', how='left')
##最后整改
Dishi['装机未归档超72小时工单占比'] = Dishi['装机未归档超72小时工单占比'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
Data_daily['装机未归档超72小时工单占比'] = Data_daily['装机未归档超72小时工单占比'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
Data_daily[yestday1 + '9'] = Data_daily[yestday1 + '9'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
Data_daily['超72小时拍照值'] = Data_daily['超72小时拍照值'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
Data_daily = pd.DataFrame(Data_daily)

Dishi = Dishi.round({'在途工单压单比': 2})  # 四舍五入
Dishi = Dishi.round({'首次响应时长': 2})  # 四舍五入
Dishi = Dishi.round({'全流程工作时长': 2})  # 四舍五入
Dishi = Dishi.round({'首次预约时长': 2})  # 四舍五入
Dishi = Dishi.round({'工单时长': 2})  # 四舍五入
Dishi = Dishi.round({'装机未归档超72小时工单占比': 4})  # 四舍五入

Dishi = Dishi.fillna(0)
Quxian = Quxian.fillna(0)
Wangge = Wangge.fillna(0)
Dishi.replace(np.inf, 0)
Quxian.replace(np.inf, 0)
Wangge.replace(np.inf, 0)

print('数据处理完毕,正在写入数据...')

with pd.ExcelWriter('装机全区重点指标日报' + '.xlsx') as writer:  # 写入结果为当前路径
    Data_daily.to_excel(writer, sheet_name='重点指标日报', startcol=0, startrow=0, index=False, header=True)
    Dishi.to_excel(writer, sheet_name='重点指标日报', startcol=0, startrow=21, index=False, header=True)
    Quxian.to_excel(writer, sheet_name='区县', startcol=0, startrow=0, index=False, header=True)
    Wangge.to_excel(writer, sheet_name='网格', startcol=0, startrow=0, index=False, header=True)
end_time = time.time()
print('处理完毕!!!总耗时%0.0f秒钟' % (end_time - start_time))
