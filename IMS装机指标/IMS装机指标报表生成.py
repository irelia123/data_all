#导入模块.
import os, glob
import time, datetime
import pandas as pd
import numpy as np
import re
#######################
df = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

IMS_Data=pd.read_excel('3月1-31IMS装机及时率报表详情.xls')



###############################################
####减免超时数据
absolve = 100
if (absolve > 0):
    print("自动减免机制--启用")
    IMS_Data.loc[(IMS_Data['工单历时'] > absolve), '是否超时'] = '未超时'
    IMS_Data.loc[(IMS_Data['工单历时'] > absolve), '工单历时'] = 0
else:
    print("自动减免机制--禁用")
##############处理完毕
GP_City = IMS_Data[(IMS_Data.区域类型=='城镇')&(IMS_Data.用户等级=='高品质')]### 筛选城镇高品质用户
GP_NC = IMS_Data[(IMS_Data.区域类型=='农村')&(IMS_Data.用户等级=='高品质')]### 筛选农村高品质用户
PT_City = IMS_Data[(IMS_Data.区域类型=='城镇')&(IMS_Data.用户等级=='普通')]### 筛选城镇普通品质用户
PT_NC = IMS_Data[(IMS_Data.区域类型=='农村')&(IMS_Data.用户等级=='普通')]### 筛选农村普通品质用户
City = IMS_Data[IMS_Data.区域类型=='城镇']### 筛选城镇用户
NC = IMS_Data[IMS_Data.区域类型=='农村']  ### 筛选农村用户
GPZ = IMS_Data[(IMS_Data.是否超时=='未超时')&(IMS_Data.用户等级=='高品质')] ### 筛选未超时，高品质用户
PTPZ = IMS_Data[(IMS_Data.是否超时=='未超时')&(IMS_Data.用户等级=='普通')] ### 筛选未超时，普通品质用户
GP = IMS_Data[IMS_Data.用户等级=='高品质']   ### 筛选高品质用户
PT = IMS_Data[IMS_Data.用户等级=='普通']    ### 筛选普通品质用户
Nocs = IMS_Data[IMS_Data.是否超时=='未超时']  ### 筛选未超时
#######################################################
ZD=IMS_Data.groupby(['地市']).size().reset_index(name='总数')
ZD=ZD.append([{'地市': '全区', '总数': ZD.apply(lambda x: x.sum()).总数}], ignore_index=True)
# 城镇高品质装移机平均时长
GP_Cityls = GP_City[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
GP_Cityls.rename(columns={'工单历时': '城镇高品质装移机平均时长'}, inplace=True)
GP_Cityls['排名'] = GP_Cityls.rank(axis=0, ascending=True,method='dense').城镇高品质装移机平均时长  # 输出排名
GP_Cityls = GP_Cityls.append(
    [{'地市': '全区', '城镇高品质装移机平均时长': GP_City[['地市', '工单历时']].mean().reset_index(name='城镇高品质装移机平均时长').at[0, '城镇高品质装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
GP_Cityls = GP_Cityls.round({'城镇高品质装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(ZD,GP_Cityls , on=['地市'], how='left')  #拼接
#############################################################
# 农村高品质装移机平均时长
GP_NCls = GP_NC[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
GP_NCls.rename(columns={'工单历时': '农村高品质装移机平均时长'}, inplace=True)
GP_NCls['排名']=GP_NCls['农村高品质装移机平均时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
GP_NCls = GP_NCls.append(
    [{'地市': '全区', '农村高品质装移机平均时长': GP_NC[['地市', '工单历时']].mean().reset_index(name='农村高品质装移机平均时长').at[0, '农村高品质装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
GP_NCls = GP_NCls.round({'农村高品质装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, GP_NCls, on=['地市'], how='left')  #拼接
#############################################################
# 城镇普通品质装移机平均时长
PT_Cityls = PT_City[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
PT_Cityls.rename(columns={'工单历时': '城镇普通品质装移机平均时长'}, inplace=True)
PT_Cityls['排名']=PT_Cityls['城镇普通品质装移机平均时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PT_Cityls = PT_Cityls.append(
    [{'地市': '全区', '城镇普通品质装移机平均时长': PT_City[['地市', '工单历时']].mean().reset_index(name='城镇普通品质装移机平均时长').at[0, '城镇普通品质装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区

PT_Cityls = PT_Cityls.round({'城镇普通品质装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, PT_Cityls, on=['地市'], how='left')  #拼接
#############################################################
# 农村普通品质装移机平均时长
PT_NCls = PT_NC[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
PT_NCls.rename(columns={'工单历时': '农村普通品质装移机平均时长'}, inplace=True)
PT_NCls['排名'] = PT_NCls.rank(axis=0, ascending=True,method='dense').农村普通品质装移机平均时长  # 输出排名
PT_NCls = PT_NCls.append(
    [{'地市': '全区', '农村普通品质装移机平均时长': PT_NC[['地市', '工单历时']].mean().reset_index(name='农村普通品质装移机平均时长').at[0, '农村普通品质装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
PT_NCls = PT_NCls.round({'农村普通品质装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, PT_NCls, on=['地市'], how='left')  #拼接
#############################################################
# 整体城镇装移机平均时长
Cityls = City[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
Cityls.rename(columns={'工单历时': '整体城镇装移机平均时长'}, inplace=True)
Cityls['排名'] = Cityls.rank(axis=0, ascending=True,method='dense').整体城镇装移机平均时长  # 输出排名
Cityls = Cityls.append(
    [{'地市': '全区', '整体城镇装移机平均时长': City[['地市', '工单历时']].mean().reset_index(name='整体城镇装移机平均时长').at[0, '整体城镇装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
Cityls = Cityls.round({'整体城镇装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, Cityls, on=['地市'], how='left')  #拼接
#############################################################
# 整体农村装移机平均时长
NCls = NC[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
NCls.rename(columns={'工单历时': '整体农村装移机平均时长'}, inplace=True)
NCls['排名'] = NCls.rank(axis=0, ascending=True,method='dense').整体农村装移机平均时长  # 输出排名
NCls = NCls.append(
    [{'地市': '全区', '整体农村装移机平均时长': NC[['地市', '工单历时']].mean().reset_index(name='整体农村装移机平均时长').at[0, '整体农村装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
NCls = NCls.round({'整体农村装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, NCls, on=['地市'], how='left')  #拼接
#############################################################
# 整体装移机平均时长
IMS_Datals = IMS_Data[['地市', '工单历时']].groupby(['地市']).mean().reset_index()
IMS_Datals.rename(columns={'工单历时': '整体装移机平均时长'}, inplace=True)
IMS_Datals['排名'] = IMS_Datals.rank(axis=0, ascending=True,method='dense').整体装移机平均时长  # 输出排名
IMS_Datals = IMS_Datals.append(
    [{'地市': '全区', '整体装移机平均时长': IMS_Data[['地市', '工单历时']].mean().reset_index(name='整体装移机平均时长').at[0, '整体装移机平均时长'], '排名': ''}],
    ignore_index=True)  # 计算全区
IMS_Datals = IMS_Datals.round({'整体装移机平均时长': 2})  # 四舍五入
GP_Cityls = pd.merge(GP_Cityls, IMS_Datals, on=['地市'], how='left')  #拼接
#############################################################
###先计算整体品质工单数
GD_NO = IMS_Data[IMS_Data.是否超时 == '未超时'][['地市', '工单历时']].groupby(['地市']).size().reset_index(name='整体未超时工单数')
GD = IMS_Data[['地市', '工单历时']].groupby(['地市']).size().reset_index(name='整体工单数')
GD_NO = pd.merge(GD_NO , GD, on=['地市'], how='left')
GD_NO['整体装移机及时率']=GD_NO.整体未超时工单数/GD_NO.整体工单数
GD_NO['排名'] = GD_NO.rank(axis=0, ascending=False,method='dense').整体装移机及时率  # 输出排名
Temp3 = IMS_Data[IMS_Data.是否超时 == '未超时'][['工单编码']].count().reset_index(name='整体未超时工单数').at[0, '整体未超时工单数'] / \
        IMS_Data[['工单编码']].count().reset_index(name='整体工单数').at[0, '整体工单数']  # 这里的Temp 临时计算全区及时率
GD_NO = GD_NO.append([{'地市': '全区', '整体工单数': IMS_Data[['工单编码']].count().reset_index(name='整体工单数').at[0, '整体工单数'],
                       '整体未超时工单数': IMS_Data[IMS_Data.是否超时 == '未超时'][['工单编码']].count().reset_index(name='整体未超时工单数').at[
                           0, '整体未超时工单数'], '整体装移机及时率': Temp3, '排名': ''}], ignore_index=True)  # 添加一行全区的数据

GD_NO = GD_NO.round({'整体装移机及时率': 4})  # 四舍五入
# ###先计算高品质工单数
GP_NO = IMS_Data[(IMS_Data.是否超时 == '未超时')&(IMS_Data.用户等级 == '高品质')][['地市', '工单历时']].groupby(['地市']).size().reset_index(name='高品质未超时工单数')
GPGD = IMS_Data[IMS_Data.用户等级 == '高品质'][['地市', '工单历时']].groupby(['地市']).size().reset_index(name='高品质工单数')
GP_NO = pd.merge(GP_NO , GPGD, on=['地市'], how='left')
GP_NO['高品质装移机及时率']=GP_NO.高品质未超时工单数/GP_NO.高品质工单数
GP_NO['排名'] = GP_NO.rank(axis=0, ascending=False,method='dense').高品质装移机及时率  # 输出排名
Temp1 = GP[GP.是否超时 == '未超时'][['工单编码']].count().reset_index(name='高品质未超时工单数').at[0, '高品质未超时工单数'] / \
        GP[['工单编码']].count().reset_index(name='高品质工单数').at[0, '高品质工单数']  # 这里的Temp 临时计算全区及时率
GP_NO = GP_NO.append([{'地市': '全区', '高品质工单数': GP[['工单编码']].count().reset_index(name='高品质工单数').at[0, '高品质工单数'],
                       '高品质未超时工单数': GP[GP.是否超时 == '未超时'][['工单编码']].count().reset_index(name='高品质未超时工单数').at[
                           0, '高品质未超时工单数'], '高品质装移机及时率': Temp1, '排名': ''}], ignore_index=True)  # 添加一行全区的数据
GP_NO = GP_NO.round({'高品质装移机及时率': 4})  # 四舍五入
GD_NO = pd.merge(GD_NO, GP_NO, on=['地市'], how='left')  #拼接


########################
###先计算普通品质工单数
PT_NO = IMS_Data[(IMS_Data.是否超时 == '未超时')&(IMS_Data.用户等级 == '普通')][['地市', '工单历时']].groupby(['地市']).size().reset_index(name='普通品质未超时工单数')
PTGD = IMS_Data[IMS_Data.用户等级 == '普通'][['地市', '工单历时']].groupby(['地市']).size().reset_index(name='普通品质工单数')
PT_NO = pd.merge(PT_NO , PTGD, on=['地市'], how='left')
PT_NO['普通品质装移机及时率']=PT_NO.普通品质未超时工单数/PT_NO.普通品质工单数
PT_NO['排名'] = PT_NO.rank(axis=0, ascending=False,method='dense').普通品质装移机及时率  # 输出排名
Temp2 = PT[PT.是否超时 == '未超时'][['工单编码']].count().reset_index(name='普通品质未超时工单数').at[0, '普通品质未超时工单数'] / \
        PT[['工单编码']].count().reset_index(name='普通品质工单数').at[0, '普通品质工单数']  # 这里的Temp 临时计算全区及时率
PT_NO = PT_NO.append([{'地市': '全区', '普通品质工单数': PT[['工单编码']].count().reset_index(name='普通品质工单数').at[0, '普通品质工单数'],
                       '普通品质未超时工单数': PT[PT.是否超时 == '未超时'][['工单编码']].count().reset_index(name='普通品质未超时工单数').at[
                           0, '普通品质未超时工单数'], '普通品质装移机及时率': Temp2, '排名': ''}], ignore_index=True)  # 添加一行全区的数据
PT_NO = PT_NO.round({'普通品质装移机及时率': 4})  # 四舍五入
GD_NO = pd.merge(GD_NO, PT_NO, on=['地市'], how='left')  #拼接


#############



GP_Cityls = pd.merge(GP_Cityls, GD_NO, on=['地市'], how='left')  #拼接
GP_Cityls = GP_Cityls.fillna(0)  # 批量替换nan 为数字 0


######################################
GP_Cityls['高品质装移机及时率']=GP_Cityls['高品质装移机及时率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
GP_Cityls['普通品质装移机及时率']=GP_Cityls['普通品质装移机及时率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
GP_Cityls['整体装移机及时率']=GP_Cityls['整体装移机及时率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
####
test = ""

for i,vl in enumerate(GP_Cityls.columns.values):
    # 匹配含有 排名 的列名，判断对该列是否需要 + 1
    if re.search(r'排名',vl):
        #print(vl)
        # 重命名列名
        column_names = GP_Cityls.columns.values
        column_names[i] = '排名'+str(i)
        GP_Cityls.columns = column_names
        # 重命名列名

        # 临时存储列名
        cs = GP_Cityls[column_names[i]]

        for bs in range(len(cs) - 1):
            # 对含有排名 0 的 单元格进行批量 + 1
            if cs[bs] == 0:
                for sb in range(len(cs) - 1):
                    cs[sb] = cs[sb] + 1

                GP_Cityls[column_names[i]] = cs
                break
#将 排名x 等列名全部替换为 排名
column_names = GP_Cityls.columns.values
for i, vl in enumerate(GP_Cityls.columns.values):
    # 匹配含有 排名 的列名，记录并修改
    if re.search(r'排名', vl):
        column_names[i] = '排名'
# 重命名列名
GP_Cityls.columns = column_names
GP_Cityls.drop(['总数'],axis=1,inplace=True)
GP_Cityls.drop(['高品质未超时工单数'],axis=1,inplace=True)
GP_Cityls.drop(['高品质工单数'],axis=1,inplace=True)
GP_Cityls.drop(['普通品质未超时工单数'],axis=1,inplace=True)
GP_Cityls.drop(['普通品质工单数'],axis=1,inplace=True)





##############################################


with pd.ExcelWriter('IMS装机指标'+'.xlsx') as writer:
    GP_Cityls.to_excel(writer, sheet_name='IMS装机指标', startcol=0, index=False, header=True)