# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 10:55:47 2021

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
todaytime1 = str(todaytime.month) + '月' + todaytime.strftime('%d') + '日'  # 日期转换str：9月01日
yestday = todaytime - datetime.timedelta(days=1)  # 昨天日期
yestday1 = str(yestday.month) + '月' + yestday.strftime('%d') + '日'  # 日期转换str：9月01日
lastmonth = todaytime - datetime.timedelta(days=30)  # 昨天日期
lastmonth = str(lastmonth.month) + '月'
Question_order = pd.read_excel('./问题单耗时清单.xlsx')
print('正在读取基础数据......')
if list(Question_order)[0] == '查询结果':  ###修改表头
    list0 = list(Question_order.iloc[0])
    ###print('******请检查新的表头是否正确：',list0)
    Question_order.columns = list0  ###重命名表头
    Question_order.dropna(subset=['问题单号'], inplace=True)  # ,inplace=True  删除地市列为空的行
    Question_order = Question_order.drop(Question_order[Question_order.问题单号 == '问题单号'].index)  # 删除多余行


Question_order['响应时长(分钟)'] = Question_order['响应耗时（小时）'] * 60
Question_order['发起至办结'] = Question_order['处理耗时（小时）'] * 60
Question_order['发起至办结归档'] = Question_order['整体耗时（小时）'] * 60
basedata = Question_order[['响应时长(分钟)','发起至办结','发起至办结归档','问题单号']]
Question_order = Question_order.drop(['响应时长(分钟)','发起至办结','发起至办结归档'],axis=1)
basedata = pd.merge(basedata,Question_order,on=['问题单号'],how='left')

basedata_diaodu = basedata[(basedata.问题归档大类 == '操作申请') |(basedata.问题归档大类 == '工单调度') | (basedata.问题归档大类 == '工单信息错误') | (basedata.问题归档大类 == '审批申请') | (basedata.问题归档大类 == '系统异常')]
basedata_diaodu = basedata_diaodu[(basedata_diaodu.问题归档小类 == '拔测工单调度') | (basedata_diaodu.问题归档小类 == '工单驳回申请') | (basedata_diaodu.问题归档小类 == '缓装/待装申请') | (basedata_diaodu.问题归档小类 == '回单质量检测申请') | (basedata_diaodu.问题归档小类 == '其他操作申请') | (basedata_diaodu.问题归档小类 == '弱光工单调度') | (basedata_diaodu.问题归档小类 == '新装工单调度') | (basedata_diaodu.问题归档小类 == '用户联系方式错误') | (basedata_diaodu.问题归档小类 == '预约首响 / 缓装 / 待装等操作不成功')]
###调度基础数据


basedata_kadan = basedata[basedata['问题归档大类'].str.contains('卡单')]
basedata_kadan = basedata_kadan[basedata_kadan['问题归档小类'].str.contains('卡单')]
###卡单基础数据


basedata_tousu = basedata[basedata['问题归档小类'].str.contains('投诉')]
###投诉基础数据

print('正在处理整合数据......')
##############################################################################################################


basedata.rename(columns={'分公司': '地市'}, inplace=True)  # 改列名
Timestat = basedata.groupby(['地市']).size().reset_index(name='工单量')
Timestat = Timestat.append([{'地市':'广西','工单量':Timestat.apply(lambda x: x.sum()).工单量}], ignore_index=True)
Timestat = pd.merge(Dishi1, Timestat, on=['地市'],how='left')

Timestat_responsetime_all = basedata[['地市','响应时长(分钟)']]
Timestat_responsetime_all = pd.DataFrame(Timestat_responsetime_all.groupby('地市')['响应时长(分钟)'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_responsetime_all = Timestat_responsetime_all.append([{'地市':'广西','响应时长(分钟)':Timestat_responsetime_all.apply(lambda x: x.sum())['响应时长(分钟)']}], ignore_index=True)
Timestat_responsetime_all = pd.merge(Dishi1,Timestat_responsetime_all,on=['地市'],how='left')

Timestat['响应时长(分钟)'] = Timestat_responsetime_all['响应时长(分钟)'] / Timestat['工单量']
Timestat['响应时长(分钟)'] = Timestat['响应时长(分钟)'].round(decimals=2)  #保留小数点后两位


Timestat_responsetime_ontime = basedata[basedata['响应时长(分钟)']<=10].groupby(['地市']).size().reset_index(name='响应及时总数')
Timestat_responsetime_ontime = Timestat_responsetime_ontime.append([{'地市':'广西','响应及时总数':Timestat_responsetime_ontime.apply(lambda x: x.sum()).响应及时总数}], ignore_index=True)
Timestat_responsetime_ontime = pd.merge(Dishi1,Timestat_responsetime_ontime,on=['地市'],how='left')

Timestat['响应及时率（10）'] = Timestat_responsetime_ontime['响应及时总数'] / Timestat['工单量']
Timestat['响应及时率（10）'] = Timestat['响应及时率（10）'].apply(lambda x: '%.2f%%' % (x*100))       #转换百分比

Timestat_responsetime_handle = basedata[['地市','发起至办结']]
Timestat_responsetime_handle = pd.DataFrame(Timestat_responsetime_handle.groupby('地市')['发起至办结'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_responsetime_handle = Timestat_responsetime_handle.append([{'地市':'广西','发起至办结':Timestat_responsetime_handle.apply(lambda x: x.sum()).发起至办结}], ignore_index=True)
Timestat_responsetime_handle = pd.merge(Dishi1,Timestat_responsetime_handle,on=['地市'],how='left')

Timestat['处理时长-发起至办结（分钟）'] = Timestat_responsetime_handle['发起至办结'] / Timestat['工单量']
Timestat['处理时长-发起至办结（分钟）'] = Timestat['处理时长-发起至办结（分钟）'].round(decimals=2)  #保留小数点后两位

Timestat_handletime_ontime = basedata[basedata['发起至办结']<=30].groupby(['地市']).size().reset_index(name='处理及时总数')
Timestat_handletime_ontime = Timestat_handletime_ontime.append([{'地市':'广西','处理及时总数':Timestat_handletime_ontime.apply(lambda x: x.sum()).处理及时总数}], ignore_index=True)
Timestat_handletime_ontime = pd.merge(Dishi1,Timestat_handletime_ontime,on=['地市'],how='left')

Timestat['处理及时率-发起至办结（30）'] = Timestat_handletime_ontime['处理及时总数'] / Timestat['工单量']
Timestat['处理及时率-发起至办结（30）'] = Timestat['处理及时率-发起至办结（30）'].apply(lambda x: '%.2f%%' % (x*100))       #转换百分比

Timestat_Alltime = basedata[['地市','发起至办结归档']]
Timestat_Alltime = pd.DataFrame(Timestat_Alltime.groupby('地市')['发起至办结归档'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_Alltime = Timestat_Alltime.append([{'地市':'广西','发起至办结归档':Timestat_Alltime.apply(lambda x: x.sum()).发起至办结归档}], ignore_index=True)
Timestat_Alltime = pd.merge(Dishi1,Timestat_Alltime,on=['地市'],how='left')

Timestat['全流程至办结归档（分钟）'] = Timestat_Alltime['发起至办结归档'] / Timestat['工单量']
Timestat['全流程至办结归档（分钟）'] = Timestat['全流程至办结归档（分钟）'].round(decimals=2)  #保留小数点后两位

Timestat_Alltime_ontime = basedata[basedata['发起至办结归档']<=70].groupby(['地市']).size().reset_index(name='全流程及时总数')
Timestat_Alltime_ontime = Timestat_Alltime_ontime.append([{'地市':'广西','全流程及时总数':Timestat_Alltime_ontime.apply(lambda x: x.sum()).全流程及时总数}], ignore_index=True)
Timestat_Alltime_ontime = pd.merge(Dishi1,Timestat_Alltime_ontime,on=['地市'],how='left')

Timestat['全流程及时率'] = Timestat_Alltime_ontime['全流程及时总数'] / Timestat['工单量']
Timestat['全流程及时率']= Timestat['全流程及时率'].apply(lambda x: '%.2f%%' % (x*100))



Timestat_question_caozuoshenqing = basedata[basedata.问题大类 =='操作申请'].groupby(['地市']).size().reset_index(name='操作申请')
Timestat_question_caozuoshenqing = Timestat_question_caozuoshenqing.append([{'地市':'广西','操作申请':Timestat_question_caozuoshenqing.apply(lambda x: x.sum()).操作申请}], ignore_index=True)

Timestat_question_kadan = basedata[basedata.问题大类 =='卡单'].groupby(['地市']).size().reset_index(name='卡单')
Timestat_question_kadan = Timestat_question_kadan.append([{'地市':'广西','卡单':Timestat_question_kadan.apply(lambda x: x.sum()).卡单}], ignore_index=True)

Timestat_question_shenpishenqing = basedata[basedata.问题大类 =='审批申请'].groupby(['地市']).size().reset_index(name='审批申请')
Timestat_question_shenpishenqing = Timestat_question_shenpishenqing.append([{'地市':'广西','审批申请':Timestat_question_shenpishenqing.apply(lambda x: x.sum()).审批申请}], ignore_index=True)

Timestat_question_gongdandiaodu = basedata[basedata.问题大类 =='工单调度'].groupby(['地市']).size().reset_index(name='工单调度')
Timestat_question_gongdandiaodu = Timestat_question_gongdandiaodu.append([{'地市':'广西','工单调度':Timestat_question_gongdandiaodu.apply(lambda x: x.sum()).工单调度}], ignore_index=True)

Timestat_question_zhanghaooyichang = basedata[basedata.问题大类 =='账号异常'].groupby(['地市']).size().reset_index(name='账号异常')
Timestat_question_zhanghaooyichang = Timestat_question_zhanghaooyichang.append([{'地市':'广西','账号异常':Timestat_question_zhanghaooyichang.apply(lambda x: x.sum()).账号异常}], ignore_index=True)

Timestat_question_zhongduanwenti = basedata[basedata.问题大类 =='终端问题'].groupby(['地市']).size().reset_index(name='终端问题')
Timestat_question_zhongduanwenti = Timestat_question_zhongduanwenti.append([{'地市':'广西','终端问题':Timestat_question_zhongduanwenti.apply(lambda x: x.sum()).终端问题}], ignore_index=True)

Timestat_question_jihuoshibai = basedata[basedata.问题大类 =='激活失败'].groupby(['地市']).size().reset_index(name='激活失败')
Timestat_question_jihuoshibai = Timestat_question_jihuoshibai.append([{'地市':'广西','激活失败':Timestat_question_jihuoshibai.apply(lambda x: x.sum()).激活失败}], ignore_index=True)

Timestat_question_ziyuancuowu = basedata[basedata.问题大类 =='资源错误'].groupby(['地市']).size().reset_index(name='资源错误')
Timestat_question_ziyuancuowu = Timestat_question_ziyuancuowu.append([{'地市':'广西','资源错误':Timestat_question_ziyuancuowu.apply(lambda x: x.sum()).资源错误}], ignore_index=True)

Timestat_question_shebeixianluguzhang = basedata[basedata.问题大类 =='设备或线路故障'].groupby(['地市']).size().reset_index(name='设备或线路故障')
Timestat_question_shebeixianluguzhang = Timestat_question_shebeixianluguzhang.append([{'地市':'广西','设备或线路故障':Timestat_question_shebeixianluguzhang.apply(lambda x: x.sum()).设备或线路故障}], ignore_index=True)

Timestat_question_kuorong = basedata[basedata.问题大类 =='扩容'].groupby(['地市']).size().reset_index(name='扩容')
Timestat_question_kuorong = Timestat_question_kuorong.append([{'地市':'广西','扩容':Timestat_question_kuorong.apply(lambda x: x.sum()).扩容}], ignore_index=True)

Timestat_question_gongdanxinxicuowu = basedata[basedata.问题大类 =='工单信息错误'].groupby(['地市']).size().reset_index(name='工单信息错误')
Timestat_question_gongdanxinxicuowu = Timestat_question_gongdanxinxicuowu.append([{'地市':'广西','工单信息错误':Timestat_question_gongdanxinxicuowu.apply(lambda x: x.sum()).工单信息错误}], ignore_index=True)

Timestat_question_guzhanggongdanhuidanguifan = basedata[basedata.问题大类 =='故障工单回单规范'].groupby(['地市']).size().reset_index(name='故障工单回单规范')
Timestat_question_guzhanggongdanhuidanguifan = Timestat_question_guzhanggongdanhuidanguifan.append([{'地市':'广西','故障工单回单规范':Timestat_question_guzhanggongdanhuidanguifan.apply(lambda x: x.sum()).故障工单回单规范}], ignore_index=True)
 
Timestat = pd.merge(Timestat,Timestat_question_caozuoshenqing,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_kadan,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_shenpishenqing,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_gongdandiaodu,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_zhanghaooyichang,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_zhongduanwenti,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_jihuoshibai,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_ziyuancuowu,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_shebeixianluguzhang,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_kuorong,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_gongdanxinxicuowu,on=['地市'],how='left').fillna(0)
Timestat = pd.merge(Timestat,Timestat_question_guzhanggongdanhuidanguifan,on=['地市'],how='left').fillna(0)

Timestat['操作申请(百分比)'] = Timestat.操作申请 / Timestat.工单量
Timestat['操作申请(百分比)'] = Timestat['操作申请(百分比)'].apply(lambda x: '%.2f%%' % (x*100))       #转换百分比
Timestat['卡单(百分比)'] = Timestat.卡单 / Timestat.工单量
Timestat['卡单(百分比)'] = Timestat['卡单(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['审批申请(百分比)'] = Timestat.审批申请 / Timestat.工单量
Timestat['审批申请(百分比)'] = Timestat['审批申请(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['工单调度(百分比)'] = Timestat.工单调度 / Timestat.工单量
Timestat['工单调度(百分比)'] = Timestat['工单调度(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['账号异常(百分比)'] = Timestat.账号异常 / Timestat.工单量
Timestat['账号异常(百分比)'] = Timestat['账号异常(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['终端问题(百分比)'] = Timestat.终端问题 / Timestat.工单量
Timestat['终端问题(百分比)'] = Timestat['终端问题(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['激活失败(百分比)'] = Timestat.激活失败 / Timestat.工单量
Timestat['激活失败(百分比)'] = Timestat['激活失败(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['资源错误(百分比)'] = Timestat.资源错误 / Timestat.工单量
Timestat['资源错误(百分比)'] = Timestat['资源错误(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['设备或线路故障(百分比)'] = Timestat.设备或线路故障 / Timestat.工单量
Timestat['设备或线路故障(百分比)'] = Timestat['设备或线路故障(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['扩容(百分比)'] = Timestat.扩容 / Timestat.工单量
Timestat['扩容(百分比)'] = Timestat['扩容(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['工单信息错误(百分比)'] = Timestat.工单信息错误 / Timestat.工单量
Timestat['工单信息错误(百分比)'] = Timestat['工单信息错误(百分比)'].apply(lambda x: '%.2f%%' % (x*100))      
Timestat['故障工单回单规范(百分比)'] = Timestat.故障工单回单规范 / Timestat.工单量
Timestat['故障工单回单规范(百分比)'] = Timestat['故障工单回单规范(百分比)'].apply(lambda x: '%.2f%%' % (x*100))


##############################################################################################################


Diaodu_last = pd.read_excel('./昨日（月）问题工单.xlsx',sheet_name='调度')
Diaodu_last = Diaodu_last.round(decimals=2)  #保留小数点后两位
Zhengti_Question_response = Timestat[['地市','响应时长(分钟)']]
Zhengti_Question_response.rename(columns={'响应时长(分钟)': todaytime1}, inplace=True)  # 改列名
Zhengti_Question_response = pd.merge(Zhengti_Question_response, Diaodu_last[['地市','整体接单昨日问题工单响应时长']],on=['地市'],how='left')
Zhengti_Question_response.rename(columns={'整体接单昨日问题工单响应时长': yestday1}, inplace=True)  # 改列名
Zhengti_Question_response['环比昨日'] = (Zhengti_Question_response[todaytime1] - Zhengti_Question_response[yestday1]) / Zhengti_Question_response[yestday1]
Zhengti_Question_response = pd.merge(Zhengti_Question_response, Diaodu_last[['地市','整体接单上月问题工单响应时长']],on=['地市'],how='left')
Zhengti_Question_response.rename(columns={'整体接单上月问题工单响应时长': lastmonth}, inplace=True)  # 改列名
Zhengti_Question_response['环比上月'] = (Zhengti_Question_response[todaytime1] - Zhengti_Question_response[lastmonth]) / Zhengti_Question_response[lastmonth]
Zhengti_Question_response = Zhengti_Question_response.round(decimals=2)  #保留小数点后两位

Zhengti_Question_handle = Timestat[['地市','处理时长-发起至办结（分钟）']]
Zhengti_Question_handle.rename(columns={'处理时长-发起至办结（分钟）': todaytime1}, inplace=True)  # 改列名
Zhengti_Question_handle = pd.merge(Zhengti_Question_handle, Diaodu_last[['地市','整体接单昨日问题工单处理时长']],on=['地市'],how='left')
Zhengti_Question_handle.rename(columns={'整体接单昨日问题工单处理时长': yestday1}, inplace=True)  # 改列名
Zhengti_Question_handle['环比昨日'] = (Zhengti_Question_handle[todaytime1] - Zhengti_Question_handle[yestday1]) / Zhengti_Question_handle[yestday1]
Zhengti_Question_handle = pd.merge(Zhengti_Question_handle, Diaodu_last[['地市','整体接单上月问题工单处理时长']],on=['地市'],how='left')
Zhengti_Question_handle.rename(columns={'整体接单上月问题工单处理时长': lastmonth}, inplace=True)  # 改列名
Zhengti_Question_handle['环比上月'] = (Zhengti_Question_handle[todaytime1] - Zhengti_Question_handle[lastmonth]) / Zhengti_Question_handle[lastmonth]
Zhengti_Question_handle = Zhengti_Question_handle.round(decimals=2)  #保留小数点后两位
Yidian_zhengti = pd.merge(Zhengti_Question_response, Zhengti_Question_handle, on=['地市'], how='left')

basedata_diaodu.rename(columns={'分公司': '地市'}, inplace=True)  # 改列名
Timestat_responsetime_diaodu = basedata_diaodu[['地市','响应时长(分钟)']]
Timestat_responsetime_diaodu = pd.DataFrame(Timestat_responsetime_diaodu.groupby('地市')['响应时长(分钟)'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_responsetime_diaodu = Timestat_responsetime_diaodu.append([{'地市':'广西','响应时长(分钟)':Timestat_responsetime_diaodu.apply(lambda x: x.sum())['响应时长(分钟)']}], ignore_index=True)
Timestat_responsetime_diaodu = pd.merge(Dishi1,Timestat_responsetime_diaodu,on=['地市'],how='left')

Timestat_diaodu = basedata_diaodu.groupby(['地市']).size().reset_index(name='工单量')
Timestat_diaodu = Timestat_diaodu.append([{'地市':'广西','工单量':Timestat_diaodu.apply(lambda x: x.sum()).工单量}], ignore_index=True)

Timestat_diaodu = pd.merge(Dishi1, Timestat_diaodu, on=['地市'],how='left')
Timestat_diaodu['响应时长(分钟)'] = Timestat_responsetime_diaodu['响应时长(分钟)'] / Timestat_diaodu['工单量']
Timestat_diaodu['响应时长(分钟)'] = Timestat_diaodu['响应时长(分钟)'].round(decimals=2)  #保留小数点后两位
#计算调度基础数据的工单量及其响应时长

Timestat_handletime_diaodu = basedata_diaodu[['地市','发起至办结']]
Timestat_handletime_diaodu = pd.DataFrame(Timestat_handletime_diaodu.groupby('地市')['发起至办结'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_handletime_diaodu = Timestat_handletime_diaodu.append([{'地市':'广西','发起至办结':Timestat_handletime_diaodu.apply(lambda x: x.sum()).发起至办结}], ignore_index=True)
Timestat_handletime_diaodu = pd.merge(Dishi1,Timestat_handletime_diaodu,on=['地市'],how='left')

Timestat_diaodu['处理时长-发起至办结（分钟）'] = Timestat_handletime_diaodu['发起至办结'] / Timestat_diaodu['工单量']
Timestat_diaodu['处理时长-发起至办结（分钟）'] = Timestat_diaodu['处理时长-发起至办结（分钟）'].round(decimals=2)  #保留小数点后两位
#计算调度基础数据的处理时长

DiaoduShenpi_Question_response = Timestat_diaodu[['地市','响应时长(分钟)']]
DiaoduShenpi_Question_response.rename(columns={'响应时长(分钟)': todaytime1}, inplace=True)  # 改列名
DiaoduShenpi_Question_response = pd.merge(DiaoduShenpi_Question_response, Diaodu_last[['地市','调度、审批类昨日问题工单响应时长']],on=['地市'],how='left')
DiaoduShenpi_Question_response.rename(columns={'调度、审批类昨日问题工单响应时长': yestday1}, inplace=True)  # 改列名
DiaoduShenpi_Question_response['环比昨日'] = (DiaoduShenpi_Question_response[todaytime1] - DiaoduShenpi_Question_response[yestday1]) / DiaoduShenpi_Question_response[yestday1]
DiaoduShenpi_Question_response = pd.merge(DiaoduShenpi_Question_response, Diaodu_last[['地市','调度、审批类上月问题工单响应时长']],on=['地市'],how='left')
DiaoduShenpi_Question_response.rename(columns={'调度、审批类上月问题工单响应时长': lastmonth}, inplace=True)  # 改列名
DiaoduShenpi_Question_response['环比上月'] = (DiaoduShenpi_Question_response[todaytime1] - DiaoduShenpi_Question_response[lastmonth]) / DiaoduShenpi_Question_response[lastmonth]
DiaoduShenpi_Question_response = DiaoduShenpi_Question_response.round(decimals=2)  #保留小数点后两位

DiaoduShenpi_Question_handle = Timestat_diaodu[['地市','处理时长-发起至办结（分钟）']]
DiaoduShenpi_Question_handle.rename(columns={'处理时长-发起至办结（分钟）': todaytime1}, inplace=True)  # 改列名
DiaoduShenpi_Question_handle = pd.merge(DiaoduShenpi_Question_handle, Diaodu_last[['地市','调度、审批类昨日问题工单处理时长']],on=['地市'],how='left')
DiaoduShenpi_Question_handle.rename(columns={'调度、审批类昨日问题工单处理时长': yestday1}, inplace=True)  # 改列名
DiaoduShenpi_Question_handle['环比昨日'] = (DiaoduShenpi_Question_handle[todaytime1] - DiaoduShenpi_Question_handle[yestday1]) / DiaoduShenpi_Question_handle[yestday1]
DiaoduShenpi_Question_handle = pd.merge(DiaoduShenpi_Question_handle, Diaodu_last[['地市','调度、审批类上月问题工单处理时长']],on=['地市'],how='left')
DiaoduShenpi_Question_handle.rename(columns={'调度、审批类上月问题工单处理时长': lastmonth}, inplace=True)  # 改列名
DiaoduShenpi_Question_handle['环比上月'] = (DiaoduShenpi_Question_handle[todaytime1] - DiaoduShenpi_Question_handle[lastmonth]) / DiaoduShenpi_Question_handle[lastmonth]
DiaoduShenpi_Question_handle = DiaoduShenpi_Question_handle.round(decimals=2)  #保留小数点后两位
Yidian_DiaoduShenpi = pd.merge(DiaoduShenpi_Question_response, DiaoduShenpi_Question_handle, on=['地市'], how='left')

Questionorder_zhanbi = Timestat[['地市','工单量','操作申请(百分比)','卡单(百分比)','审批申请(百分比)','工单调度(百分比)','账号异常(百分比)','终端问题(百分比)','激活失败(百分比)','资源错误(百分比)','设备或线路故障(百分比)']]

Questionorder_zhanbi.rename(columns={'操作申请(百分比)': '操作申请'}, inplace=True)  # 改列名
Questionorder_zhanbi.rename(columns={'卡单(百分比)': '卡单'}, inplace=True)
Questionorder_zhanbi.rename(columns={'审批申请(百分比)': '审批申请'}, inplace=True)
Questionorder_zhanbi.rename(columns={'工单调度(百分比)': '工单调度'}, inplace=True)
Questionorder_zhanbi.rename(columns={'账号异常(百分比)': '账号异常'}, inplace=True)
Questionorder_zhanbi.rename(columns={'终端问题(百分比)': '终端问题'}, inplace=True)
Questionorder_zhanbi.rename(columns={'激活失败(百分比)': '激活失败'}, inplace=True)
Questionorder_zhanbi.rename(columns={'资源错误(百分比)': '资源错误'}, inplace=True)
Questionorder_zhanbi.rename(columns={'设备或线路故障(百分比)': '设备或线路故障'}, inplace=True)
Questionorder_zhanbi['扩容'] =  Timestat.扩容 / Timestat.工单量
Questionorder_zhanbi['工单信息错误'] =  Timestat.工单信息错误 / Timestat.工单量
Questionorder_zhanbi['故障工单回单规范'] =  Timestat.故障工单回单规范 / Timestat.工单量
Questionorder_zhanbi['其他'] =  Questionorder_zhanbi.扩容 + Questionorder_zhanbi.工单信息错误 + Questionorder_zhanbi.故障工单回单规范
Questionorder_zhanbi['其他'] = Questionorder_zhanbi['其他'].apply(lambda x: '%.1f%%' % (x*100))
Questionorder_zhanbi = Questionorder_zhanbi.drop(['扩容','工单信息错误','故障工单回单规范'],axis=1)


##############################################################################################################


Kadan_last = pd.read_excel('./昨日（月）问题工单.xlsx', sheet_name='卡单')
Kadan_last = Kadan_last.round(decimals=2)  #保留小数点后两位
basedata_kadan.rename(columns={'分公司': '地市'}, inplace=True)  # 改列名
Timestat_responsetime_kadan = basedata_kadan[['地市','响应时长(分钟)']]
Timestat_responsetime_kadan = pd.DataFrame(Timestat_responsetime_kadan.groupby('地市')['响应时长(分钟)'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_responsetime_kadan = Timestat_responsetime_kadan.append([{'地市':'广西','响应时长(分钟)':Timestat_responsetime_kadan.apply(lambda x: x.sum())['响应时长(分钟)']}], ignore_index=True)
Timestat_responsetime_kadan = pd.merge(Dishi1,Timestat_responsetime_kadan,on=['地市'],how='left')

Timestat_kadan = basedata_kadan.groupby(['地市']).size().reset_index(name='工单量')
Timestat_kadan = Timestat_kadan.append([{'地市':'广西','工单量':Timestat_kadan.apply(lambda x: x.sum()).工单量}], ignore_index=True)

Timestat_kadan = pd.merge(Dishi1, Timestat_kadan, on=['地市'],how='left')
Timestat_kadan['响应时长(分钟)'] = Timestat_responsetime_kadan['响应时长(分钟)'] / Timestat_kadan['工单量']
Timestat_kadan['响应时长(分钟)'] = Timestat_kadan['响应时长(分钟)'].round(decimals=2)  #保留小数点后两位
#计算卡单基础数据的工单量及其响应时长

Timestat_handletime_kadan = basedata_kadan[['地市','发起至办结']]
Timestat_handletime_kadan = pd.DataFrame(Timestat_handletime_kadan.groupby('地市')['发起至办结'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_handletime_kadan = Timestat_handletime_kadan.append([{'地市':'广西','发起至办结':Timestat_handletime_kadan.apply(lambda x: x.sum()).发起至办结}], ignore_index=True)
Timestat_handletime_kadan = pd.merge(Dishi1,Timestat_handletime_kadan,on=['地市'],how='left')

Timestat_kadan['处理时长-发起至办结（分钟）'] = Timestat_handletime_kadan['发起至办结'] / Timestat_kadan['工单量']
Timestat_kadan['处理时长-发起至办结（分钟）'] = Timestat_kadan['处理时长-发起至办结（分钟）'].round(decimals=2)  #保留小数点后两位
#计算卡单基础数据的处理时长

Kadan_Question_response = Timestat_kadan[['地市','响应时长(分钟)']]
Kadan_Question_response.rename(columns={'响应时长(分钟)': todaytime1}, inplace=True)  # 改列名
Kadan_Question_response = pd.merge(Kadan_Question_response, Kadan_last[['地市','卡单类昨日问题工单响应时长']],on=['地市'],how='left')
Kadan_Question_response.rename(columns={'卡单类昨日问题工单响应时长': yestday1}, inplace=True)  # 改列名
Kadan_Question_response['环比昨日'] = (Kadan_Question_response[todaytime1] - Kadan_Question_response[yestday1]) / Kadan_Question_response[yestday1]
Kadan_Question_response = pd.merge(Kadan_Question_response, Kadan_last[['地市','卡单类上月问题工单响应时长']],on=['地市'],how='left')
Kadan_Question_response.rename(columns={'卡单类上月问题工单响应时长': lastmonth}, inplace=True)  # 改列名
Kadan_Question_response['环比上月'] = (Kadan_Question_response[todaytime1] - Kadan_Question_response[lastmonth]) / Kadan_Question_response[lastmonth]
Kadan_Question_response = Kadan_Question_response.round(decimals=2)  #保留小数点后两位

Kadan_Question_handle = Timestat_kadan[['地市','处理时长-发起至办结（分钟）']]
Kadan_Question_handle.rename(columns={'处理时长-发起至办结（分钟）': todaytime1}, inplace=True)  # 改列名
Kadan_Question_handle = pd.merge(Kadan_Question_handle, Kadan_last[['地市','卡单类昨日问题工单处理时长']],on=['地市'],how='left')
Kadan_Question_handle.rename(columns={'卡单类昨日问题工单处理时长': yestday1}, inplace=True)  # 改列名
Kadan_Question_handle['环比昨日'] = (Kadan_Question_handle[todaytime1] - Kadan_Question_handle[yestday1]) / Kadan_Question_handle[yestday1]
Kadan_Question_handle = pd.merge(Kadan_Question_handle, Kadan_last[['地市','卡单类上月问题工单处理时长']],on=['地市'],how='left')
Kadan_Question_handle.rename(columns={'卡单类上月问题工单处理时长': lastmonth}, inplace=True)  # 改列名
Kadan_Question_handle['环比上月'] = (Kadan_Question_handle[todaytime1] - Kadan_Question_handle[lastmonth]) / Kadan_Question_handle[lastmonth]
Kadan_Question_handle = Kadan_Question_handle.round(decimals=2)  #保留小数点后两位
Yidian_Kadan = pd.merge(Kadan_Question_response, Kadan_Question_handle, on=['地市'], how='left')
Yidian_Kadan = Yidian_Kadan.fillna(0)


##############################################################################################################


Tousu_last = pd.read_excel('./昨日（月）问题工单.xlsx', sheet_name='投诉')
Tousu_last = Tousu_last.round(decimals=2)  #保留小数点后两位
basedata_tousu.rename(columns={'分公司': '地市'}, inplace=True)  # 改列名
Timestat_responsetime_tousu = basedata_tousu[['地市','响应时长(分钟)']]
Timestat_responsetime_tousu = pd.DataFrame(Timestat_responsetime_tousu.groupby('地市')['响应时长(分钟)'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_responsetime_tousu = Timestat_responsetime_tousu.append([{'地市':'广西','响应时长(分钟)':Timestat_responsetime_tousu.apply(lambda x: x.sum())['响应时长(分钟)']}], ignore_index=True)
Timestat_responsetime_tousu = pd.merge(Dishi1,Timestat_responsetime_tousu,on=['地市'],how='left')

Timestat_tousu = basedata_tousu.groupby(['地市']).size().reset_index(name='工单量')
Timestat_tousu = Timestat_tousu.append([{'地市':'广西','工单量':Timestat_tousu.apply(lambda x: x.sum()).工单量}], ignore_index=True)

Timestat_tousu = pd.merge(Dishi1, Timestat_tousu, on=['地市'],how='left')
Timestat_tousu['响应时长(分钟)'] = Timestat_responsetime_tousu['响应时长(分钟)'] / Timestat_tousu['工单量']
Timestat_tousu['响应时长(分钟)'] = Timestat_tousu['响应时长(分钟)'].round(decimals=2)  #保留小数点后两位
##

Timestat_handletime_tousu = basedata_tousu[['地市','发起至办结']]
Timestat_handletime_tousu = pd.DataFrame(Timestat_handletime_tousu.groupby('地市')['发起至办结'].sum()).reset_index()    #根据地市计算相对应地市的总和
Timestat_handletime_tousu = Timestat_handletime_tousu.append([{'地市':'广西','发起至办结':Timestat_handletime_tousu.apply(lambda x: x.sum()).发起至办结}], ignore_index=True)
Timestat_handletime_tousu = pd.merge(Dishi1,Timestat_handletime_tousu,on=['地市'],how='left')

Timestat_tousu['处理时长-发起至办结（分钟）'] = Timestat_handletime_tousu['发起至办结'] / Timestat_tousu['工单量']
Timestat_tousu['处理时长-发起至办结（分钟）'] = Timestat_tousu['处理时长-发起至办结（分钟）'].round(decimals=2)  #保留小数点后两位
##

Tousu_Question_response = Timestat_tousu[['地市','响应时长(分钟)']]
Tousu_Question_response.rename(columns={'响应时长(分钟)': todaytime1}, inplace=True)  # 改列名
Tousu_Question_response = pd.merge(Tousu_Question_response, Tousu_last[['地市','投诉类昨日问题工单响应时长']],on=['地市'],how='left')
Tousu_Question_response.rename(columns={'投诉类昨日问题工单响应时长': yestday1}, inplace=True)  # 改列名
Tousu_Question_response['环比昨日'] = (Tousu_Question_response[todaytime1] - Tousu_Question_response[yestday1]) / Tousu_Question_response[yestday1]
Tousu_Question_response = pd.merge(Tousu_Question_response, Tousu_last[['地市','投诉类上月问题工单响应时长']],on=['地市'],how='left')
Tousu_Question_response.rename(columns={'投诉类上月问题工单响应时长': lastmonth}, inplace=True)  # 改列名
Tousu_Question_response['环比上月'] = (Tousu_Question_response[todaytime1] - Tousu_Question_response[lastmonth]) / Tousu_Question_response[lastmonth]
Tousu_Question_response = Tousu_Question_response.round(decimals=2)  #保留小数点后两位

Tousu_Question_handle = Timestat_tousu[['地市','处理时长-发起至办结（分钟）']]
Tousu_Question_handle.rename(columns={'处理时长-发起至办结（分钟）': todaytime1}, inplace=True)  # 改列名
Tousu_Question_handle = pd.merge(Tousu_Question_handle, Tousu_last[['地市','投诉类昨日问题工单处理时长']],on=['地市'],how='left')
Tousu_Question_handle.rename(columns={'投诉类昨日问题工单处理时长': yestday1}, inplace=True)  # 改列名
Tousu_Question_handle['环比昨日'] = (Tousu_Question_handle[todaytime1] - Tousu_Question_handle[yestday1]) / Tousu_Question_handle[yestday1]
Tousu_Question_handle = pd.merge(Tousu_Question_handle, Tousu_last[['地市','投诉类上月问题工单处理时长']],on=['地市'],how='left')
Tousu_Question_handle.rename(columns={'投诉类上月问题工单处理时长': lastmonth}, inplace=True)  # 改列名
Tousu_Question_handle['环比上月'] = (Tousu_Question_handle[todaytime1] - Tousu_Question_handle[lastmonth]) / Tousu_Question_handle[lastmonth]
Tousu_Question_handle = Tousu_Question_handle.round(decimals=2)  #保留小数点后两位
Yidian_Tousu = pd.merge(Tousu_Question_response, Tousu_Question_handle, on=['地市'], how='left')
Yidian_Tousu = Yidian_Tousu.fillna(0)

print('处理数据完毕 正在导出数据......')


title = pd.DataFrame({'全区一点支撑及时率（整体接单）': ['']})
title1 = pd.DataFrame({'全区一点支撑及时率（调度、审批类）': ['']})
title2 = pd.DataFrame({'区综调一点支撑问题工单类型占比': ['']})
title3 = pd.DataFrame({'一点支撑卡单类数据统计': ['']})
title4 = pd.DataFrame({'一点支撑投诉类数据统计': ['']})

with pd.ExcelWriter('综调日报（调度、卡单、投诉）.xlsx') as writer:
     basedata.to_excel(writer,sheet_name='基础数据', startcol=0, index=False, header=True)
     basedata_diaodu.to_excel(writer,sheet_name='调度基础数据', startcol=0, index=False, header=True)
     basedata_kadan.to_excel(writer,sheet_name='卡单基础数据', startcol=0, index=False, header=True)
     basedata_tousu.to_excel(writer,sheet_name='投诉基础数据', startcol=0, index=False, header=True)
     Timestat.to_excel(writer,sheet_name='时长及时率统计', startcol=0, index=False, header=True)

     title.to_excel(writer,sheet_name='调度', startcol=0, startrow=0, index=False, header=True)
     Yidian_zhengti.to_excel(writer,sheet_name='调度', startcol=0, startrow=1, index=False, header=True)
     
     title1.to_excel(writer,sheet_name='调度', startcol=13, startrow=0, index=False, header=True)
     Yidian_DiaoduShenpi.to_excel(writer,sheet_name='调度', startcol=13, startrow=1, index=False, header=True)

     title2.to_excel(writer,sheet_name='调度', startcol=0, startrow=20, index=False, header=True)
     Questionorder_zhanbi.to_excel(writer,sheet_name='调度', startcol=0, startrow=21, index=False, header=True)

     title3.to_excel(writer,sheet_name='卡单', startcol=0, startrow=0, index=False, header=True)
     Yidian_Kadan.to_excel(writer,sheet_name='卡单', startcol=0, startrow=1, index=False, header=True)

     title4.to_excel(writer,sheet_name='投诉', startcol=0, startrow=0, index=False, header=True)
     Yidian_Tousu.to_excel(writer,sheet_name='投诉', startcol=0, startrow=1, index=False, header=True)


print('导出数据完毕 程序运行结束')
