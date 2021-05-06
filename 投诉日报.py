# -*- coding: utf-8 -*-
"""
Created on Mon Apr 19 15:08:57 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np
##初始化
Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))
Dishi1 = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','广西']},
    pd.Index(range(15)))

todaytime = datetime.datetime.strptime('2021-4-1','%Y-%m-%d') #手动输入日期
first_day = datetime.datetime(todaytime.year, todaytime.month, 1, 23, 59, 59)  # datetime类型 2019-09-01 00:00:00  取当月1日
today = str(todaytime.month) + '月' + todaytime.strftime('%d') + '日'  # 日期转换str：9月01日


ZT_data = pd.read_excel('./投诉在途.xlsx')
RJguidangliang_Quxian = pd.read_excel('./日均归档量.xlsx', sheet_name='区县')
RJguidangliang_Wangge = pd.read_excel('./日均归档量.xlsx', sheet_name='网格')
QZ_time = pd.read_excel('./时长.xlsx', sheet_name='千兆',header=1)
PT_time = pd.read_excel('./时长.xlsx', sheet_name='家宽',header=1)
zhibiao_last = pd.read_excel('./昨日（月）指标.xlsx',header=1)





ZT_data1 = ZT_data[['投诉时间']]
ZT_data1.rename(columns={'投诉时间': '工单开始时间'}, inplace=True)  # 改列名
ZT_data1[first_day] = (first_day - pd.to_datetime(ZT_data1.工单开始时间)).dt.days*24 + (first_day - pd.to_datetime(ZT_data1.工单开始时间)).dt.seconds/3600 + 8

ZT_data1.ix[(ZT_data1[first_day]> 24) & (ZT_data1[first_day]< 48),'tag1']='超24-48小时未竣工'
ZT_data1.ix[(ZT_data1[first_day]> 48) & (ZT_data1[first_day]< 100),'tag1']='超48-100小时未竣工'
ZT_data1.ix[ZT_data1[first_day]> 100,'tag1']='超100小时未竣工'
ZT_data1.ix[ZT_data['工单状态'] == '待归档','是否要剔除']='TRUE'
ZT_data1.ix[ZT_data['工单状态'] != '待归档','是否要剔除']='FALSE'
ZT_data1.ix[ZT_data['速率'] == '1000M','是否千兆']='是'
ZT_data1.ix[ZT_data['速率'] != '1000M','是否千兆']='否'
ZT_data1['工单流水号'] = ZT_data['工单流水号']
ZT_data = pd.merge(ZT_data1, ZT_data, on=['工单流水号'], how='left')


##############################################################################################################


Quxian_QZ_ZTorder = ZT_data[(ZT_data.是否千兆 == '是') & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','责任区县']).size().reset_index(name='千兆投诉在途工单数')
#千兆投诉在途工单数


Quxian_PT_ZTorder = ZT_data[(ZT_data.是否千兆 == '否') & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','责任区县']).size().reset_index(name='宽带投诉在途工单数-不含千兆')
#普通宽带投诉在途工单数

Quxian_chao48order = ZT_data[(ZT_data[first_day] > 48) & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','责任区县']).size().reset_index(name='超48小时工单数')
#超48小时工单数

QuxianOrder = pd.merge(Quxian_QZ_ZTorder, Quxian_PT_ZTorder, on=['责任地市','责任区县'], how='left').fillna(0)
QuxianOrder = pd.merge(QuxianOrder, Quxian_chao48order, on=['责任地市','责任区县'], how='left').fillna(0)

QuxianOrder['投诉在途超48小时工单占比'] = QuxianOrder['超48小时工单数'] / (QuxianOrder['千兆投诉在途工单数'] + QuxianOrder['宽带投诉在途工单数-不含千兆'])
QuxianOrder['投诉在途超48小时工单占比'] = QuxianOrder['投诉在途超48小时工单占比'].apply(lambda x: '%.2f%%' % (x*100))
QuxianOrder.rename(columns={'责任地市': '地市'}, inplace=True)  # 改列名
QuxianOrder.rename(columns={'责任区县': '区县'}, inplace=True)  # 改列名
QuxianOrder = pd.merge(QuxianOrder, RJguidangliang_Quxian, on=['地市','区县'], how='left').fillna(0)

QuxianOrder['在途压单比'] = (QuxianOrder['千兆投诉在途工单数'] + QuxianOrder['宽带投诉在途工单数-不含千兆']) / QuxianOrder['日均归档量']
QuxianOrder['在途压单比'] = QuxianOrder['在途压单比'].round(decimals=2)  #保留小数点后两位
QuxianOrder = QuxianOrder.drop(['日均归档量'],axis=1)


##############################################################################################################


Wangge_QZ_ZTorder = ZT_data[(ZT_data.是否千兆 == '是') & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','网格名称']).size().reset_index(name='千兆投诉在途工单数')
#千兆投诉在途工单数


Wangge_PT_ZTorder = ZT_data[(ZT_data.是否千兆 == '否') & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','网格名称']).size().reset_index(name='宽带投诉在途工单数-不含千兆')
#普通宽带投诉在途工单数

Wangge_chao48order = ZT_data[(ZT_data[first_day] > 48) & (ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市','网格名称']).size().reset_index(name='超48小时工单数')
#超48小时工单数

WanggeOrder = pd.merge(Wangge_QZ_ZTorder, Wangge_PT_ZTorder, on=['责任地市','网格名称'], how='left').fillna(0)
WanggeOrder = pd.merge(WanggeOrder, Wangge_chao48order, on=['责任地市','网格名称'], how='left').fillna(0)

WanggeOrder['投诉在途超48小时工单占比'] = WanggeOrder['超48小时工单数'] / (WanggeOrder['千兆投诉在途工单数'] + WanggeOrder['宽带投诉在途工单数-不含千兆'])
WanggeOrder['投诉在途超48小时工单占比'] = WanggeOrder['投诉在途超48小时工单占比'].apply(lambda x: '%.2f%%' % (x*100))
WanggeOrder.rename(columns={'责任地市': '地市'}, inplace=True)  # 改列名
WanggeOrder.rename(columns={'网格名称': '网格'}, inplace=True)  # 改列名
WanggeOrder = pd.merge(WanggeOrder, RJguidangliang_Wangge, on=['地市','网格'], how='left').fillna(0)

WanggeOrder['在途压单比'] = (WanggeOrder['千兆投诉在途工单数'] + WanggeOrder['宽带投诉在途工单数-不含千兆']) / WanggeOrder['日均归档量']
WanggeOrder['在途压单比'] = WanggeOrder['在途压单比'].round(decimals=2)  #保留小数点后两位
WanggeOrder = WanggeOrder.drop(['日均归档量'],axis=1)


##############################################################################################################


Temp_tongji = ZT_data[(ZT_data.是否要剔除 == 'FALSE')].groupby(['责任地市']).size().reset_index(name='投诉在途总工单数')
Temp_tongji.rename(columns={'责任地市': '地市'}, inplace=True)
Temp_tongji = Temp_tongji.append([{'地市':'全区','投诉在途总工单数':Temp_tongji.apply(lambda x: x.sum())['投诉在途总工单数']}], ignore_index=True)
Temp_tongji = pd.merge(Dishi, Temp_tongji, on=['地市'], how='left')
Temp_tongji_chao48 = ZT_data[(ZT_data.是否要剔除 == 'FALSE') & (ZT_data[first_day] > 48)].groupby(['责任地市']).size().reset_index(name='超48小时工单')
Temp_tongji_chao48.rename(columns={'责任地市': '地市'}, inplace=True)
Temp_tongji_chao48 = Temp_tongji_chao48.append([{'地市':'全区','超48小时工单':Temp_tongji_chao48.apply(lambda x: x.sum())['超48小时工单']}], ignore_index=True)
Temp_tongji = pd.merge(Temp_tongji, Temp_tongji_chao48, on=['地市'], how='left')
Temp_tongji['在途超48小时工单占比'] = Temp_tongji['超48小时工单'] / Temp_tongji['投诉在途总工单数']

Temp_tongji = pd.merge(Temp_tongji, zhibiao_last[['地市','日均归档量']], on=['地市'], how='left')
Temp_tongji['在途压单比'] = Temp_tongji['投诉在途总工单数'] / Temp_tongji['日均归档量']
Temp_tongji['在途压单比'] = Temp_tongji['在途压单比'].round(decimals=2)  #保留小数点后两位

QZ_FirstResponse = QZ_time[['地市','工作时.3']]
QZ_FirstResponse.rename(columns={'工作时.3': today}, inplace=True)
QZ_FirstResponse = pd.merge(QZ_FirstResponse, zhibiao_last[['地市','昨日千兆首响']], on=['地市'], how='left')
QZ_FirstResponse['还比昨日'] = (QZ_FirstResponse[today] - QZ_FirstResponse['昨日千兆首响']) / QZ_FirstResponse['昨日千兆首响']
QZ_FirstResponse = pd.merge(QZ_FirstResponse, zhibiao_last[['地市','上月千兆首响']], on=['地市'], how='left')
QZ_FirstResponse['还比上月'] = (QZ_FirstResponse[today] - QZ_FirstResponse['上月千兆首响']) / QZ_FirstResponse['上月千兆首响']
QZ_FirstResponse = QZ_FirstResponse.fillna(0)
QZ_FirstResponse = QZ_FirstResponse.round(decimals=2)  #保留小数点后两位
#千兆首响

QZ_Report = QZ_time[['地市','工作时-全流程.2']]
QZ_Report.rename(columns={'工作时-全流程.2': today}, inplace=True)
QZ_Report = pd.merge(QZ_Report, zhibiao_last[['地市','昨日千兆投诉处理时长']], on=['地市'], how='left')
QZ_Report['还比昨日'] = (QZ_Report[today] - QZ_Report['昨日千兆投诉处理时长']) / QZ_Report['昨日千兆投诉处理时长']
QZ_Report = pd.merge(QZ_Report, zhibiao_last[['地市','上月千兆投诉处理时长']], on=['地市'], how='left')
QZ_Report['还比上月'] = (QZ_Report[today] - QZ_Report['上月千兆投诉处理时长']) / QZ_Report['上月千兆投诉处理时长']
QZ_Report = QZ_Report.fillna(0)
QZ_Report = QZ_Report.round(decimals=2)  #保留小数点后两位
#千兆投诉处理时长(工作时)

PT_FirstResponse = PT_time[['地市','工作时.3']]
PT_FirstResponse.rename(columns={'工作时.3': today}, inplace=True)
PT_FirstResponse = pd.merge(PT_FirstResponse, zhibiao_last[['地市','昨日普通宽带首响']], on=['地市'], how='left')
PT_FirstResponse['还比昨日'] = (PT_FirstResponse[today] - PT_FirstResponse['昨日普通宽带首响']) / PT_FirstResponse['昨日普通宽带首响']
PT_FirstResponse = pd.merge(PT_FirstResponse, zhibiao_last[['地市','上月普通宽带首响']], on=['地市'], how='left')
PT_FirstResponse['还比上月'] = (PT_FirstResponse[today] - PT_FirstResponse['上月普通宽带首响']) / PT_FirstResponse['上月普通宽带首响']
PT_FirstResponse = PT_FirstResponse.fillna(0)
PT_FirstResponse = PT_FirstResponse.round(decimals=2)  #保留小数点后两位
#普通宽带首响

PT_Report = PT_time[['地市','小时-全流程.2']]
PT_Report.rename(columns={'小时-全流程.2': today}, inplace=True)
PT_Report = pd.merge(PT_Report, zhibiao_last[['地市','昨日普通投诉处理时长']], on=['地市'], how='left')
PT_Report['还比昨日'] = (PT_Report[today] - PT_Report['昨日普通投诉处理时长']) / PT_Report['昨日普通投诉处理时长']
PT_Report = pd.merge(PT_Report, zhibiao_last[['地市','上月普通投诉处理时长']], on=['地市'], how='left')
PT_Report['还比上月'] = (PT_Report[today] - PT_Report['上月普通投诉处理时长']) / PT_Report['上月普通投诉处理时长']
PT_Report = PT_Report.fillna(0)
PT_Report = PT_Report.round(decimals=2)  #保留小数点后两位
#普通投诉处理时长(小时)

ZT_chao48 = Temp_tongji[['地市','在途超48小时工单占比']]
ZT_chao48.rename(columns={'在途超48小时工单占比': today}, inplace=True)
ZT_chao48 = pd.merge(ZT_chao48, zhibiao_last[['地市','上月在途超48小时工单占比']], on=['地市'], how='left')
ZT_chao48['还比昨日'] = (ZT_chao48[today] - ZT_chao48['上月在途超48小时工单占比']) / ZT_chao48['上月在途超48小时工单占比']
ZT_chao48 = pd.merge(ZT_chao48, zhibiao_last[['地市','上月普通宽带首响']], on=['地市'], how='left')
ZT_chao48['还比上月'] = (ZT_chao48[today] - ZT_chao48['上月普通宽带首响']) / ZT_chao48['上月普通宽带首响']
ZT_chao48 = ZT_chao48.fillna(0)
ZT_chao48 = ZT_chao48.round(decimals=2)  #保留小数点后两位
ZT_chao48[today] = ZT_chao48[today].apply(lambda x: '%.2f%%' % (x*100))
ZT_chao48['上月在途超48小时工单占比'] = ZT_chao48['上月在途超48小时工单占比'].apply(lambda x: '%.2f%%' % (x*100))
ZT_chao48['上月普通宽带首响'] = ZT_chao48['上月普通宽带首响'].apply(lambda x: '%.2f%%' % (x*100))
#在途超48小时工单占比


ZT_yadanbi = Temp_tongji[['地市','在途压单比']]
ZT_yadanbi.rename(columns={'在途压单比': today}, inplace=True)
ZT_yadanbi = pd.merge(ZT_yadanbi, zhibiao_last[['地市','昨日在途压单比']], on=['地市'], how='left')
ZT_yadanbi['还比昨日'] = (ZT_yadanbi[today] - ZT_yadanbi['昨日在途压单比']) / ZT_yadanbi['昨日在途压单比']
ZT_yadanbi = pd.merge(ZT_yadanbi, zhibiao_last[['地市','上月在途压单比']], on=['地市'], how='left')
ZT_yadanbi['还比上月'] = (ZT_yadanbi[today] - ZT_yadanbi['上月在途压单比']) / ZT_yadanbi['上月在途压单比']
ZT_yadanbi = ZT_yadanbi.fillna(0)
ZT_yadanbi = ZT_yadanbi.round(decimals=2)  #保留小数点后两位
#上月在途压单比


title1 = pd.DataFrame({'千兆首响': ['']})
title2 = pd.DataFrame({'千兆投诉处理时长(工作时)': ['']})
title3 = pd.DataFrame({'普通宽带首响': ['']})
title4 = pd.DataFrame({'普通投诉处理时长(小时)': ['']})
title5 = pd.DataFrame({'在途超48小时工单占比': ['']})
title6 = pd.DataFrame({'在途压单比': ['']})
print('正在将数据导出到文件夹...')
with pd.ExcelWriter('投诉全区重点指标日报' + '.xlsx') as writer:  # 写入结果为当前路径
    
    title1.to_excel(writer, sheet_name='重点指标日报', startcol=0, startrow=1, index=False, header=True)
    QZ_FirstResponse.to_excel(writer, sheet_name='重点指标日报', startcol=0, startrow=2, index=False, header=True)
    
    title2.to_excel(writer, sheet_name='重点指标日报', startcol=6, startrow=1, index=False, header=True)
    QZ_Report.to_excel(writer, sheet_name='重点指标日报', startcol=6, startrow=2, index=False, header=True)
    
    title3.to_excel(writer, sheet_name='重点指标日报', startcol=12, startrow=1, index=False, header=True)
    PT_FirstResponse.to_excel(writer, sheet_name='重点指标日报', startcol=12, startrow=2, index=False, header=True)
    
    title4.to_excel(writer, sheet_name='重点指标日报', startcol=18, startrow=1, index=False, header=True)
    PT_Report.to_excel(writer, sheet_name='重点指标日报', startcol=18, startrow=2, index=False, header=True)
    
    title5.to_excel(writer, sheet_name='重点指标日报', startcol=24, startrow=1, index=False, header=True)
    ZT_chao48.to_excel(writer, sheet_name='重点指标日报', startcol=24, startrow=2, index=False, header=True)
    
    title6.to_excel(writer, sheet_name='重点指标日报', startcol=30, startrow=1, index=False, header=True)
    ZT_yadanbi.to_excel(writer, sheet_name='重点指标日报', startcol=30, startrow=2, index=False, header=True)
    QuxianOrder.to_excel(writer, sheet_name='区县', startcol=0, startrow=0, index=False, header=True)
    WanggeOrder.to_excel(writer, sheet_name='网格', startcol=0, startrow=0, index=False, header=True)
    ZT_data.to_excel(writer, sheet_name='投诉在途清单', startcol=0, startrow=0, index=False, header=True)
    
print('导出完毕！')
    
    
    
    