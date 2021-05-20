# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 09:44:43 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np

Time = input('>>请输入年-月-日格式如：xxxx-xx-xx：')#输入时间格式例如：20201-04-27
#Time = '2021-04-27'
Gd_Zb = pd.read_excel('./' + Time +'归档.xlsx', inplace=True) #读取/eomsj表
Gd_Pb = pd.read_excel('./' + Time +'派发.xlsx', inplace=True)#读取eomsp表
Zt_Pb = pd.read_excel('./' + Time +'在途.xlsx', inplace=True)#读取eomszt表
Zj_Pb = pd.read_excel('./自定义数据.xlsx',sheet_name='中间数据',inplace=True) #读取中间数据
Zdysj = pd.read_excel('./自定义数据.xlsx', sheet_name='地市日报自定义数据',inplace=True) #读取自定义数据
Zd_Qx = pd.read_excel('./自定义数据.xlsx',sheet_name='区县表自定义数据' , inplace=True) #读取自定义数据
Zd_Wg = pd.read_excel('./自定义数据.xlsx',sheet_name = '网格表自定义数据', inplace=True) #读取自定网格义数据
Zd_Bxguishudi = pd.read_excel('./自定义数据.xlsx',sheet_name = '不详归属地市', inplace=True) #读取自定网格义数据


#---------------------------设置时间


todaytime = datetime.datetime.strptime(Time,'%Y-%m-%d') #---------------------手动输入日期
todaytime = datetime.datetime(todaytime.year, todaytime.month, todaytime.day, 18, 00, 00)#设置当日为18点

Sz_Time = datetime.datetime.strptime(Time,'%Y-%m-%d') #---------------------------手动输入日期
Sz_Time = datetime.datetime(Sz_Time.year, Sz_Time.month, Sz_Time.day, 23, 59, 59)#设置当日为23, 59, 59点

Ju_Time = datetime.datetime.strptime('2020-08-31','%Y-%m-%d') #---------------------------------手动输入日期
Ju_Time = datetime.datetime(Ju_Time.year, Ju_Time.month, Ju_Time.day, 23, 59, 59)#设置当日为23, 59, 59点

Chaoshi2 = datetime.datetime.strptime(Time,'%Y-%m-%d') #-------------------------------手动输入日期
Chaoshi2 = datetime.datetime(Chaoshi2.year, Chaoshi2.month, Chaoshi2.day, 23, 59, 59)#设置当日为00:00:00点
print('时间设置完成')

 #当日归档明细
 
Gd_Zb = Gd_Zb.fillna('')
Gd_Zb['安装地址'] = ''
Gd_Zb['投诉内容'] = ''
print('当日归档明细完成')
#---------剔除问题解决单$和家亲
Gd_Zb = Gd_Zb[(Gd_Zb.工单类型 != '问题解决单')&(Gd_Zb.工单类型 !='和家亲')]


Gd_Pb = Gd_Pb.fillna('')
Gd_Pb['安装地址'] = ''
Gd_Pb['投诉内容'] = ''
print('剔除问题解决单$和家亲完成')
#---------剔除和睦$和家亲
Gd_Pb = Gd_Pb[(Gd_Pb.工单类型 != '问题解决单')&(Gd_Pb.工单类型 !='和家亲')]
#---------剔除和睦$和家亲

Zt_Pb = Zt_Pb.fillna('')
Zt_Pb['安装地址'] = ''
Zt_Pb['投诉内容'] = ''
Zt_Pb = Zt_Pb[(Zt_Pb.工单类型 != '问题解决单')&(Zt_Pb.工单类型 !='和家亲')]
Zt_Pb = Zt_Pb.reset_index(drop=True) #重新排序
Gd_Zb = Gd_Zb.reset_index(drop=True) #重新排序
Gd_Pb = Gd_Pb.reset_index(drop=True) #重新排序
print('剔除和睦$和家亲完成')

#当日派单明细
Sd = Gd_Pb[['首次到单时间']]
Sd['时间差'] = (pd.to_datetime(todaytime)-pd.to_datetime(Sd.首次到单时间)).dt.days*24+(pd.to_datetime(todaytime)-pd.to_datetime(Sd.首次到单时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Sd['是否18点到达'] = Sd.时间差 > 0
Sd = Sd.drop(['首次到单时间','时间差'],axis=1)#删除列
Gd_Pb = pd.concat([Sd, Gd_Pb], axis=1)  # 插入第一列为标段信息---------当日派单明细

print('当日派单明细已完成')

#当日在途明细

if list(Zt_Pb)[0] == '查询结果':  ###修改表头
    list0 = list(Zt_Pb.iloc[0])
    ###print('******请检查新的表头是否正确：',list0)
    Zt_Pb.columns = list0  ###重命名表头
    Zt_Pb.dropna(subset=['工单流水号'], inplace=True)  # ,inplace=True  删除空行
    Zt_Pb = Zt_Pb.drop(Zt_Pb[Zt_Pb.工单流水号 == '工单流水号'].index)  # 删除多余行



Zaitu_mx = Zt_Pb[['工单流水号']]

#----------工单开始时间

Zaitu_Scddsj =Zt_Pb[['工单流水号','首次到单时间']]
Zaitu_Scddsj.rename(columns={'首次到单时间': '工单开始时间'}, inplace=True)  # 改列名

#-------2021/4/27 24:00:00

Zaitu_sjc = Zt_Pb[['工单流水号']]
Zaitu_sjc['时间差'] = (pd.to_datetime(Sz_Time)-pd.to_datetime(Zt_Pb.首次到单时间)).dt.days*24+(pd.to_datetime(Sz_Time)-pd.to_datetime(Zt_Pb.首次到单时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改

#--------------tag1

Zaitu_tag = Zaitu_sjc[['工单流水号','时间差']]
Zaitu_tag['tag1'] = ''
Zaitu_tag['tag1'][Zaitu_tag['时间差'] >= 100] = '超100小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Zaitu_tag['tag1'][(Zaitu_tag['时间差'] < 100)&(Zaitu_tag['时间差'] >= 48)] = '超48-100小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Zaitu_tag['tag1'][(Zaitu_tag['时间差'] < 48)&(Zaitu_tag['时间差'] >= 24)] = '超24-48小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Zaitu_tag= Zaitu_tag.drop(['时间差'],axis=1)#删除列

#----------是否超24小时

Zaitu_Sfcessxs = Zaitu_sjc[['工单流水号','时间差']]
Zaitu_Sfcessxs['是否超24小时']  = Zaitu_Sfcessxs['时间差'] > 24
Zaitu_Sfcessxs = Zaitu_Sfcessxs.drop(['时间差'] , axis = 1)#删除列

#-------------是否多次驳回

Zaitu_Sfdcbh = Zt_Pb[['工单流水号','重派次数']]
Zaitu_Sfdcbh['是否多次驳回'] = (Zaitu_Sfdcbh['重派次数'] -2) >= 0 
Zaitu_Sfdcbh = Zaitu_Sfdcbh.drop(['重派次数'] , axis = 1)#删除列

#------------区域类型

Zaitu_Qylx = Zt_Pb[['工单流水号','区域类型']]

#----------是否9月前工单

Zaitu_Sfjyqgd = Zaitu_Scddsj[['工单流水号','工单开始时间']]
Zaitu_Sfjyqgd['是否9月前工单']  = (pd.to_datetime(Ju_Time)-pd.to_datetime(Zaitu_Sfjyqgd.工单开始时间)).dt.days*24+(pd.to_datetime(Ju_Time)-pd.to_datetime(Zaitu_Sfjyqgd.工单开始时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Zaitu_Sfjyqgd['是否9月前工单'] = Zaitu_Sfjyqgd['是否9月前工单'] > 0
Zaitu_Sfjyqgd  = Zaitu_Sfjyqgd.drop(['工单开始时间'],axis = 1)#删除列

#-----------是否待质检

Zaitu_Sfdzj = Zt_Pb[['工单流水号','工单状态']]
Zaitu_Sfdzj['是否待质检'] = Zaitu_Sfdzj['工单状态'] == '质检'
Zaitu_Sfdzj = Zaitu_Sfdzj.drop(['工单状态'], axis = 1)#删除列

#----------是否要剔除

Zaitu_Tichu = Zt_Pb[['工单流水号','工单类型','工单状态']]
Zaitu_Tichu.ix[(Zaitu_Tichu.工单类型==''),'是否要剔除'] = False
Zaitu_Tichu.ix[(Zaitu_Tichu.工单类型=='家庭宽带'),'是否要剔除'] = False
Zaitu_Tichu.ix[(Zaitu_Tichu.工单状态=='待归档'),'是否要剔除'] = True
Zaitu_Tichu['是否要剔除'] =Zaitu_Tichu['是否要剔除'].fillna(False)
Zaitu_Tichu = Zaitu_Tichu.drop(['工单类型','工单状态'] , axis = 1)#删除列

#-----------催单次数

Zaitu_Cdcs = Zt_Pb[['工单流水号','催办次数','客服催办次数']]
Zaitu_Cdcs['催单次数'] = Zaitu_Cdcs['催办次数'] + Zaitu_Cdcs['客服催办次数']
Zaitu_Cdcs = Zaitu_Cdcs.drop(['催办次数','客服催办次数'],axis=1)#删除列

#------------被客服催单

Zaitu_Bkfcd = Zt_Pb[['工单流水号','客服催办次数']]

Zaitu_Bkfcd['被客服催单'] = (Zaitu_Bkfcd['客服催办次数'] - 1) >= 0
Zaitu_Bkfcd = Zaitu_Bkfcd.drop(['客服催办次数'] , axis = 1)#删除列

#----------------是否超时

Zaitu_Sfcs = Zt_Pb[['工单流水号','广西时限']]
Zaitu_Sfcs['是否超时'] = (pd.to_datetime(Chaoshi2)-pd.to_datetime(Zaitu_Sfcs.广西时限)).dt.days*24+(pd.to_datetime(Chaoshi2)-pd.to_datetime(Zaitu_Sfcs.广西时限)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Zaitu_Sfcs['是否超时'] = Zaitu_Sfcs['是否超时']+ 8
Zaitu_Sfcs = Zaitu_Sfcs.drop(['广西时限'], axis = 1)#删除列

# 20:00

zaitu_Esdz = Zt_Pb[['工单流水号']]
Zaitu_Bdll = Zt_Pb[['责任代维班组']]
Zhongjian_Kl = Zj_Pb[['代维班组','地市代维一']]
list1 = list()#创建三个空列表
list2 = list()
list3 = list()
for i in Zaitu_Bdll.责任代维班组:#遍历dataFrame数据再放入列表中
    list1.append(i)
for j in Zhongjian_Kl.代维班组:
    list2.append(j)   
for k in Zhongjian_Kl.地市代维一:
    list3.append(k)                

list4 = list()
for x in range(len(list1)):
    for y in range(len(list2)):
        if list1[x] == list2[y]: 
            list4.append(list3[y])
            break
    else:
        list4.append(None)    
df = pd.DataFrame(list4, columns=['责任代维班组'])#把列表转化为dataFrame
zaitu_Esdz['20:00:00'] = df[['责任代维班组']]

#------------地市

zaitu_Dishi = Zt_Pb[['工单流水号']]
list5 = list()
list6 = list()
list7 = list()
list8 = list()
list9 = list()
list10 = list()
list11 = list()
zhongjian_Lm = Zj_Pb[['地市代维一','地市二']]
for a in zhongjian_Lm.地市代维一:
    list5.append(a)
for b in zhongjian_Lm.地市二:
    list6.append(b)

for e in Zt_Pb.责任地市:
    list8.append(e)
for f in Zd_Bxguishudi.工单流水号:
    list9.append(f)
for g in Zd_Bxguishudi.责任地市:
    list10.append(g)
for p in Zt_Pb.工单流水号:
    list11.append(p)    

for c in range(len(list4)):
    for d in range(len(list5)):
        if list4[c] == list5[d]: 
            list7.append(list6[d])
            break
        elif list8[c] != '':
            list7.append(list8[c])
            break
    else:
        for h in range(len(list9)):
            if list11[c] == list9[h]:
                list7.append(list10[h])
                break
cf = pd.DataFrame(list7, columns=['地市'])#把列表转化为dataFrame
zaitu_Dishi['地市'] = cf[['地市']]
print('地市已完成')          

#------------8:00 

zaitu_Bdz = Zt_Pb[['工单流水号']]
list_g = list()
list_h = list()
list_l = list()
list_n = list()
list_i = list()
list_j = list()
list_Bo = list()
list_Br = list()
list_Bdz = list()
zhongjian_Ln = Zj_Pb[['维护区域二','代维一','地市代维一','简称一','代维公司','简称']]
for gg in zhongjian_Ln.维护区域二:
    list_g.append(gg)
for hh in zhongjian_Ln.代维一:
    list_h.append(hh)
for ll in zhongjian_Ln.地市代维一:
    list_l.append(ll)
for nn in zhongjian_Ln.简称一:
    list_n.append(nn)
for ii in zhongjian_Ln.代维公司:
    list_i.append(ii)
for jj in zhongjian_Ln.简称:
    list_j.append(jj)

for bo in Zt_Pb.责任区县:
    list_Bo.append(bo)
for br in Zt_Pb.责任代维公司:
    list_Br.append(br)
for ii in range(len(list_Bo)):
    for jj in range(len(list_g)):
        if list_Bo[ii] == list_g[jj]: 
            list_Bdz.append(list_h[jj])
            break
        elif list4[ii]==list_l[jj]:
            list_Bdz.append(list_n[jj])
            break
        elif list_Br[ii]==list_i[jj]:
            list_Bdz.append(list_j[jj])
            break
    else:
        list_Bdz.append(None)
bdz = pd.DataFrame(list_Bdz, columns=['八点整'])#把列表转化为dataFrame
zaitu_Bdz['8:00:00'] = bdz[['八点整']]
print('8:00数据处理完成')

#----------代维最终

zaitu_Dwzz = Zt_Pb[['工单流水号']]
list_Dwzz = list()
list_p = list()
list_q = list()
for pp in Zj_Pb.地市三:
    list_p.append(pp)
for qq in Zj_Pb.代维二:
    list_q.append(qq)
for xx in range(len(list4)):
    for yy in range(len(list_l)):
        if list4[xx] == list_l[yy]:
            list_Dwzz.append(list_n[yy])
            break
        elif list7[xx] == list_p[yy]:
            list_Dwzz.append(list_q[yy])
            break
    else:
        list_Dwzz.append('铁通')
dwzz = pd.DataFrame(list_Dwzz, columns = ['代维最终'])#把列表转化为dataFrame
zaitu_Dwzz['代维(最终)'] = dwzz[['代维最终']]

#------------地市代维最终

zaitu_Dsdwzz = Zt_Pb[['工单流水号']]
list_Dsdwzz = list()
for kk in range(len(list7)):
    if list_Dwzz[kk] == list_Bdz[kk]:
        if list4[kk] != None:
            list_Dsdwzz.append(list4[kk])
        else:
            list_Dsdwzz.append(list7[kk] + list_Dwzz[kk])
    else:
        list_Dsdwzz.append(list7[kk] + list_Dwzz[kk])
dsdwzz = pd.DataFrame(list_Dsdwzz, columns = ['地市代维最终'])#把列表转化为dataFrame
zaitu_Dsdwzz['地市代维(最终)'] = dsdwzz[['地市代维最终']]
#-------质检时限

zaitu_Zjsx = Zt_Pb[['工单流水号']]
gxsx = Zt_Pb[['广西时限']]
list_Zjsx = list()
for gg in gxsx.广西时限:
    strTime = gg
    startTime1 = datetime.datetime.strptime(strTime, "%Y-%m-%d %X")  # 把strTime转化为时间格式,后面的秒位自动补位的
    if '20:00:00' in gg:
        startTime2 = (startTime1 + datetime.timedelta(hours=16)).strftime("%Y-%m-%d %X") #把startTime时间加16小时
        list_Zjsx.append(startTime2)
    else:
        startTime2 = (startTime1 + datetime.timedelta(hours=4)).strftime("%Y-%m-%d %X") #把startTime时间加4小时
        list_Zjsx.append(startTime2)
zjsx = pd.DataFrame(list_Zjsx, columns = ['质检时限'])
zaitu_Zjsx['质检时限'] = zjsx[['质检时限']]    

#----------------拼接当日在途明细

Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Dishi, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Sfcessxs, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Sfdcbh, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Scddsj, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Qylx, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Dsdwzz, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Dwzz, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Bdz, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Esdz, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_sjc, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_tag, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, zaitu_Zjsx, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Sfjyqgd, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Sfdzj, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Tichu, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Cdcs, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Bkfcd, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zaitu_Sfcs, on=['工单流水号'], how='left')
Zaitu_mx =pd.merge(Zaitu_mx, Zt_Pb, on=['工单流水号'], how='left')
Drztmx = Zaitu_mx.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
Drztmx.insert(19, Drztmx.pop(Drztmx.index('工单流水号')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
Zaitu_mx = Zaitu_mx[Drztmx]      #排列后数据。
Gd_Sj = Zaitu_mx


#-----------------------------------------------地市日报

Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

#当日归档

Temp1 = Gd_Zb.groupby(['责任地市']).size().reset_index(name='当日归档')
Temp1 = Temp1.append([{'责任地市':'全区','当日归档':Temp1.apply(lambda x:x.sum()).当日归档}],ignore_index=True)
Temp1.rename(columns={'责任地市': '地市'}, inplace=True) #改列名
Temp1 = pd.merge(Dishi, Temp1, on=['地市'], how='left') #拼接

print('当日归档已完成')
#当日排单

Temp2 = Gd_Pb.groupby(['责任地市']).size().reset_index(name='当日派单')
Temp2 = Temp2.append([{'责任地市':'全区','当日派单':Temp2.apply(lambda x:x.sum()).当日派单}],ignore_index=True)
Temp2.rename(columns={'责任地市': '地市'}, inplace=True) #改列名
print('当日排单已完成')
#其中18点后派单

Temp3 = Gd_Pb[Gd_Pb.是否18点到达 == False].groupby(['责任地市']).size().reset_index(name='其中18点后派单')
Temp3 = Temp3.append([{'责任地市':'全区','其中18点后派单':Temp3.apply(lambda x:x.sum()).其中18点后派单}],ignore_index=True)
Temp3.rename(columns={'责任地市': '地市'}, inplace=True) #改列名
print('其中18点后派单已完成')
#导入自定义数据
Zdysj.分母 = Zdysj.分母.fillna('')
Zdysj.归档总量 = Zdysj.归档总量.fillna('')

#昨日在途

Temp4 = Dishi[['地市']]
Temp4['昨日在途']= Zdysj[['昨日在途']]

print('昨日在途已完成')
#24点在途

Temp5 = Gd_Sj[Gd_Sj.是否要剔除 == False].groupby(['地市']).size().reset_index(name='二十四点在途')
Temp5 = Temp5.append([{'地市':'全区','二十四点在途':Temp5.apply(lambda x:x.sum()).二十四点在途}],ignore_index=True)
Temp5 = pd.merge(Dishi, Temp5, on=['地市'], how='left') #拼接
print('24点在途已完成')
#压单比<=0.8

Temp5['压单比<=0.8'] = Temp5.二十四点在途 / Zdysj.日均归档量
Temp5['压单比<=0.8']=Temp5['压单比<=0.8'].round(decimals=2) #保留两位小数
print('压单比<=0.8已完成')
#在途目标

Temp6 = Dishi[['地市']]
Temp6['日均归档量'] =Zdysj[['日均归档量']]
Temp6 ['在途目标']= Temp6.日均归档量 * 0.8
Temp6.在途目标=Temp6.在途目标 // 1 #保留0位小数
Temp6 = Temp6.drop(['日均归档量'],axis=1)#删除列
print('在途目标已完成')
#在途目标差值

Temp7= Dishi[['地市']]
Temp7['在途目标'] =Temp6[['在途目标']]
Temp7['在途目标差值']= Temp7['在途目标'] - Temp5['二十四点在途']
Temp7= Temp7.drop(['在途目标'],axis=1)#删除列
print('在途目标差值已完成')
#超长超时工单

#超过48小时

Temp8 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['地市']).size().reset_index(name='超过四十八小时')
Temp8 = Temp8.append([{'地市':'全区','超过四十八小时':Temp8.apply(lambda x:x.sum()).超过四十八小时}],ignore_index=True)
Temp8.超过四十八小时 = Temp8.超过四十八小时.fillna(0)
Temp8= pd.merge(Dishi, Temp8, on=['地市'], how='left') #拼接
print('超过48小时已完成')

#超过7日

Temp9 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['地市']).size().reset_index(name='超过七日')
Temp9 = Temp9.append([{'地市':'全区','超过七日':Temp9.apply(lambda x:x.sum()).超过七日}],ignore_index=True)
Temp9.超过七日 = Temp9.超过七日.fillna(0)
print('超过7日已完成')
#超过15日

Temp10 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['地市']).size().reset_index(name='超过十五日')
Temp10= Temp10.append([{'地市':'全区','超过十五日':Temp10.apply(lambda x:x.sum()).超过十五日}],ignore_index=True)
Temp10.超过十五日 = Temp10.超过十五日.fillna(0)
print('超过15日已完成')
#超过30日

Temp11 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['地市']).size().reset_index(name='超过三十日')
Temp11= Temp11.append([{'地市':'全区','超过三十日':Temp11.apply(lambda x:x.sum()).超过三十日}],ignore_index=True)
Temp11.超过三十日 = Temp11.超过三十日.fillna(0)
print('超过30日已完成')
#超过60日

Temp12 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['地市']).size().reset_index(name='超过六十日')
Temp12 = Temp12.append([{'地市':'全区','超过六十日':Temp12.apply(lambda x:x.sum()).超过六十日}],ignore_index=True)
Temp12.超过六十日 = Temp12.超过六十日.fillna(0)
print('超过60日已完成')
#超48小时占比

Temp13 = Dishi[['地市']]
Temp13 ['二十四点在途']= Temp5[['二十四点在途']]
Temp13['超四十八小时占比'] = Temp8.超过四十八小时 / Temp13.二十四点在途
Temp13['超四十八小时占比']=Temp13['超四十八小时占比'].apply(lambda x: '%.2f%%' % (x*100))

print('超48小时占比已完成')

#-----------------------------------拼接地市日报

Dishi1 = Dishi
Dishi1['地市1'] = Dishi[['地市']]
Temp1 = pd.merge(Temp1, Temp2, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp3, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp4, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp5, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp6, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp7, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1,Dishi1 , on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp8, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp9, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp10, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp11, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp12, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1, Temp13, on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1,Dishi1 , on=['地市'], how='left') #拼接
Temp1 = pd.merge(Temp1,Zdysj , on=['昨日在途'], how='left') #拼接
Dishi_Rb = Temp1  #-----------------------------------------------地市日报表
Dishi_Rb = Dishi_Rb.drop(['二十四点在途_y'],axis=1)#删除列
Dishi_Rb = Dishi_Rb.fillna(0)
print('#-----------------------------------------------地市日报已完成')
#-------------------------表3-区县
print('#-----------------------------------------------表3-区县')

Zd_Qx.归档总量 = Zd_Qx.归档总量.fillna('')
#----------------地市
#----------------区县分公司
Quxian = Zd_Qx[['地市']]
Quxian['区县分公司'] = Zd_Qx[['区县分公司']]

#-------当日完成

Drwc = Quxian
Zrqu =Gd_Zb.groupby(['责任区县']).size().reset_index(name='当日完成')
Zrqu.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Drwc = pd.merge(Drwc,Zrqu , on=['区县分公司'], how='left') #拼接
Drwc.当日完成 = Drwc.当日完成.fillna(0)
print('#-------当日完成已完成')
#-------当日投诉派单

Drtspd = Quxian
Zrqu1 =Gd_Pb.groupby(['责任区县']).size().reset_index(name='当日投诉派单')
Zrqu1.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Drtspd = pd.merge(Drtspd,Zrqu1 , on=['区县分公司'], how='left') #拼接
Drtspd.当日投诉派单 = Drtspd.当日投诉派单.fillna(0)
print('当日投诉派单已完成')
#---------其中18点后派单

Sbdhpd = Quxian
Zrqu2 =Gd_Pb[(Gd_Pb.是否18点到达 == False)].groupby(['责任区县']).size().reset_index(name='其中18点后派单')
Zrqu2.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Sbdhpd = pd.merge(Sbdhpd,Zrqu2 , on=['区县分公司'], how='left') #拼接
Sbdhpd.其中18点后派单 = Sbdhpd.其中18点后派单.fillna(0)
print('其中18点后派单已完成')
#----------昨日在途

Zrzt =  Quxian[['地市','区县分公司']]
Zrzt['昨日在途'] = Zd_Qx[['昨日在途']]
Zrzt = Zrzt[['地市','区县分公司','昨日在途']]
print('昨日在途已完成')
#----------------24点在途

Essdzt = Quxian
Zrqu3 =Gd_Sj[(Gd_Sj.是否要剔除 == False)].groupby(['责任区县']).size().reset_index(name='二十四点在途')
Zrqu3.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Essdzt = pd.merge(Essdzt,Zrqu3 , on=['区县分公司'], how='left') #拼接
Essdzt.二十四点在途 = Essdzt.二十四点在途.fillna(0)
print('24点在途已完成')
#----------------压单比小于0.8

Ydbxyldb = Quxian[['地市','区县分公司']]
Ydbxyldb['压单比小于08'] = Essdzt.二十四点在途 / Zd_Qx.日均归档量
print('压单比小于0.8已完成')


#-------------在途目标

Ztmb = Quxian[['地市','区县分公司']]
Ztmb['在途目标'] = Zd_Qx.日均归档量 * 0.8
Ztmb.在途目标 = Ztmb.在途目标 // 1
print('在途目标已完成')

#-----------在途目标差值

Ztmbcz = Quxian[['地市','区县分公司']]
Ztmbcz['在途目标差值'] = Essdzt.二十四点在途 - Ztmb.在途目标
print('在途目标差值已完成')
#----------八月拍照未归档量

Bypzwgd = Quxian
Zrqu4 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.是否9月前工单 == True)].groupby(['责任区县']).size().reset_index(name='八月拍照未归档量')
Zrqu4.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Bypzwgd = pd.merge(Bypzwgd,Zrqu4 , on=['区县分公司'], how='left') #拼接
Bypzwgd.八月拍照未归档量 =Bypzwgd.八月拍照未归档量.fillna(0)
print('八月拍照未归档量已完成')
#-------->48小时

dysbxs =  Quxian
Zrqu5 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['责任区县']).size().reset_index(name='大于四十八小时')
Zrqu5.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
dysbxs = pd.merge(dysbxs,Zrqu5 , on=['区县分公司'], how='left') #拼接
dysbxs.大于四十八小时 = dysbxs.大于四十八小时.fillna(0)
print('#-------->48小时已完成')
#-----大于7天

Dyqt =  Quxian
Zrqu6 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['责任区县']).size().reset_index(name='大于7天')
Zrqu6.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyqt = pd.merge(Dyqt,Zrqu6 , on=['区县分公司'], how='left') #拼接
Dyqt.大于7天 = Dyqt.大于7天.fillna(0)
print('#-----大于7天已完成')
#------大于15天

Dyswt =  Quxian
Zrqu7 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['责任区县']).size().reset_index(name='大于15天')
Zrqu7.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyswt = pd.merge(Dyswt,Zrqu7 , on=['区县分公司'], how='left') #拼接
Dyswt.大于15天 = Dyswt.大于15天.fillna(0)
print('#------大于15天已完成')
#-----大于30天

Dysst =  Quxian
Zrqu8 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['责任区县']).size().reset_index(name='大于30天')
Zrqu8.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dysst = pd.merge(Dysst,Zrqu8 , on=['区县分公司'], how='left') #拼接
Dysst.大于30天 = Dysst.大于30天.fillna(0)
print('#-----大于30天已完成')
#------大于60天

Dylst =  Quxian
Zrqu9 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['责任区县']).size().reset_index(name='大于60天')
Zrqu9.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dylst = pd.merge(Dylst,Zrqu9 , on=['区县分公司'], how='left') #拼接
Dylst.大于60天 = Dylst.大于60天.fillna(0)
print('#------大于60天已完成')
#------大于90天

Dyjst =  Quxian
Zrqu10 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 2160)].groupby(['责任区县']).size().reset_index(name='大于90天')
Zrqu10.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyjst = pd.merge(Dyjst,Zrqu10 , on=['区县分公司'], how='left') #拼接
Dyjst.大于90天 = Dyjst.大于90天.fillna(0)
print('#------大于90天已完成')
#--------超48小时占比

Csbxszb =  Quxian[['地市','区县分公司']]
Csbxszb ['超48小时占比']= (dysbxs.大于四十八小时 / Essdzt.二十四点在途).apply(lambda x: '%.2f%%' % (x*100))
Csbxszb= Csbxszb.drop(['区县分公司'],axis=1)#删除列
print('#--------超48小时占比已完成')

#----------联接区县表

Quxianbiao = Zd_Qx[['地市','区县分公司']]
Quxianbiao['当日完成'] = Drwc[['当日完成']]
Quxianbiao['当日投诉派单'] = Drtspd[['当日投诉派单']]
Quxianbiao['其中18点后派单'] = Sbdhpd[['其中18点后派单']]
Quxianbiao['昨日在途'] = Zrzt[['昨日在途']]
Quxianbiao['二十四点在途'] = Essdzt[['二十四点在途']]
Quxianbiao['压单比小于0.8'] = Ydbxyldb[['压单比小于08']]
Quxianbiao['在途目标'] = Ztmb[['在途目标']]
Quxianbiao['在途目标差值'] = Ztmbcz[['在途目标差值']]
Quxianbiao['八月拍照未归档量'] = Bypzwgd[['八月拍照未归档量']]
Quxianbiao['区县分公司1'] =  Zd_Qx[['区县分公司']]
Quxianbiao['大于四十八小时'] = dysbxs[['大于四十八小时']]
Quxianbiao['大于7天'] = Dyqt[['大于7天']]
Quxianbiao['大于15天'] = Dyswt[['大于15天']]
Quxianbiao['大于30天'] = Dysst[['大于30天']]
Quxianbiao['大于60天'] = Dylst[['大于60天']]
Quxianbiao['大于90天'] = Dyjst[['大于90天']]
Quxianbiao['超48小时占比'] = Csbxszb[['超48小时占比']]
Quxianbiao['归档总量'] = Zd_Qx[['归档总量']]
Quxianbiao['日均归档量'] = Zd_Qx[['日均归档量']]#-------------区县
Quxianbiao.区县分公司 = Quxianbiao.区县分公司.fillna('')
Quxianbiao.区县分公司1 = Quxianbiao.区县分公司1.fillna(0)
print('区县表已完成')
#--------------------表4网格

print('表4网格开始')

Zd_Wg.日期 =Zd_Wg.日期.fillna('')
Zd_Wg.归档总量 =Zd_Wg.归档总量.fillna('')
Wangge = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]

#------------------24点在途

Wg_Essdzt1 =Gd_Sj[(Gd_Sj.是否要剔除 == False)].groupby(['网格名称']).size().reset_index(name='二十四点在途')
Wg_Essdzt1.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_Essdzt1 = pd.merge(Wangge,Wg_Essdzt1 , on=['所属网格'], how='left') #拼接
Wg_Essdzt1.二十四点在途 = Wg_Essdzt1.二十四点在途.fillna(0) #把nan填充为0
print('24点在途已完成')
#-------------------压单比小于等于0.8

Wg_Ydbxyldb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ydbxyldb ['压单比小于等于零点八']= (Wg_Essdzt1.二十四点在途 / Zd_Wg.日均归档量).round(decimals=2) #保留两位小数
Wg_Ydbxyldb.压单比小于等于零点八 = Wg_Ydbxyldb.压单比小于等于零点八.fillna(0) #把nan填充为0
print('压单比小于等于0.8已完成')
#----------------------在途目标

Wg_Ztmb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ztmb['在途目标'] = Zd_Wg.日均归档量 *0.8
Wg_Ztmb.在途目标 = Wg_Ztmb.在途目标 //1  #取整数部分
print('在途目标已完成')

#-------------------在途目标差值

Wg_Ztmbcz = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ztmbcz['在途目标差值'] = Wg_Ztmb.在途目标 - Wg_Essdzt1.二十四点在途
print('在途目标差值已完成')
#------------------八月拍照未归档


Wg_Bypzwgd =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.是否9月前工单 == True)].groupby(['网格名称']).size().reset_index(name='八月拍照未归档')

Wg_Bypzwgd.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_Bypzwgd = pd.merge(Wangge,Wg_Bypzwgd , on=['所属网格'], how='left') #拼接
Wg_Bypzwgd.八月拍照未归档 = Wg_Bypzwgd.八月拍照未归档.fillna(0) #把nan填充为0
print('八月拍照未归档已完成')
#------------------------ 大于48小时

Wg_dysbxs =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['网格名称']).size().reset_index(name='大于四十八小时')
Wg_dysbxs.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dysbxs = pd.merge(Wangge,Wg_dysbxs , on=['所属网格'], how='left') #拼接
Wg_dysbxs.大于四十八小时 =Wg_dysbxs.大于四十八小时.fillna(0) #把nan填充为0

print('大于48小时已完成')
#-----大于7天

Wg_dyqt =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['网格名称']).size().reset_index(name='大于7天')
Wg_dyqt.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyqt = pd.merge(Wangge,Wg_dyqt , on=['所属网格'], how='left') #拼接
Wg_dyqt.大于7天 =Wg_dyqt.大于7天.fillna(0) #把nan填充为0

print('大于7天已完成')
#------大于15天

Wg_dyswt =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['网格名称']).size().reset_index(name='大于15天')
Wg_dyswt.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyswt = pd.merge(Wangge,Wg_dyswt , on=['所属网格'], how='left') #拼接
Wg_dyswt.大于15天 = Wg_dyswt.大于15天.fillna(0) #把nan填充为0

print('大于15天已完成')
#-----大于30天

Wg_dysst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['网格名称']).size().reset_index(name='大于30天')
Wg_dysst.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dysst = pd.merge(Wangge,Wg_dysst , on=['所属网格'], how='left') #拼接
Wg_dysst.大于30天 = Wg_dysst.大于30天.fillna(0) #把nan填充为0

print('大于30天已完成')
#------大于60天

Wg_dylst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['网格名称']).size().reset_index(name='大于60天')
Wg_dylst.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dylst = pd.merge(Wangge,Wg_dylst , on=['所属网格'], how='left') #拼接
Wg_dylst.大于60天= Wg_dylst.大于60天.fillna(0) #把nan填充为0

print('大于60天已完成')
#------大于90天

Wg_dyjst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 2160)].groupby(['网格名称']).size().reset_index(name='大于90天')
Wg_dyjst .rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyjst  = pd.merge(Wangge,Wg_dyjst, on=['所属网格'], how='left') #拼接
Wg_dyjst.大于90天= Wg_dyjst.大于90天.fillna(0) #把nan填充为0

print('大于90天已完成')
#--------------------超48小时占比

Wg_Csbxszb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Csbxszb['超48小时占比'] = (Wg_dysbxs.大于四十八小时 / Wg_Essdzt1.二十四点在途).apply(lambda x: '%.2f%%' % (x*100))
print('超48小时占比已完成')
#----------联接区县表

Wanggebiao = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wanggebiao['二十四点在途'] = Wg_Essdzt1[['二十四点在途']]
Wanggebiao['压单比小于等于零点八'] = Wg_Ydbxyldb [['压单比小于等于零点八']]
Wanggebiao['在途目标'] = Wg_Ztmb[['在途目标']]
Wanggebiao['在途目标差值'] = Wg_Ztmbcz[['在途目标差值']]
Wanggebiao['八月拍照未归档'] = Wg_Bypzwgd[['八月拍照未归档']]
Wanggebiao['所属网格1'] = Zd_Wg[['所属网格']]
Wanggebiao['大于四十八小时'] = Wg_dysbxs.大于四十八小时
Wanggebiao['大于7天'] = Wg_dyqt.大于7天
Wanggebiao['大于15天'] =Wg_dyswt.大于15天
Wanggebiao['大于30天'] =Wg_dysst.大于30天
Wanggebiao['大于60天'] =Wg_dylst.大于60天
Wanggebiao['大于90天'] =Wg_dyjst.大于90天
Wanggebiao['超48小时占比'] =Wg_Csbxszb[['超48小时占比']]
Wanggebiao['归档总量'] =Zd_Wg[['归档总量']]
Wanggebiao['日均归档量'] =Zd_Wg[['日均归档量']]
print('区县表已完成')
#--------------------修改列名

#--------------修改当日派单明细列名

Gd_Pb.rename(columns={'是否18点到达':'2021/4/11 18:00:00'}, inplace=True)  # 改列名

#-------------修改当日在途明细列名
Gd_Sj.rename(columns={'区域类型_x':'区域类型'}, inplace=True)  # 改列名
Gd_Sj.rename(columns={'时间差':'2021/4/11 24:00:00'}, inplace=True)  # 改列名
#----------------修改地市日报列名

Dishi_Rb.rename(columns={'二十四点在途':'24点在途'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'地市1_x':'地市'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超过四十八小时':'>48小时'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超过七日':'>7天'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超过十五日':'>15天'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超过三十日':'>30日'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超过六十日':'>60日'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'超四十八小时占比':'超48小时占比＜15%'}, inplace=True)  # 改列名
Dishi_Rb.rename(columns={'地市1_y':'地市'}, inplace=True)  # 改列名

#------------------修改区县列名

Quxianbiao.rename(columns={'二十四点在途':'24点在途'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'压单比小于0.8':'压单比<=0.8'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'八月拍照未归档量':'8月拍照未归档量'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'区县分公司1':'区县分公司'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于四十八小时':'>48小时'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于7天':'>7天'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于15天':'>15天'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于30天':'>30日'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于60天':'>60日'}, inplace=True)  # 改列名
Quxianbiao.rename(columns={'大于90天':'>90日'}, inplace=True)  # 改列名

#----------------修改网格表
Wanggebiao.rename(columns={'二十四点在途':'24点在途'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'压单比小于等于零点八':'压单比<=0.8'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'八月拍照未归档':'8月拍照未归档'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'所属网格1':'所属网格'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'大于四十八小时':'>48小时'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'大于7天':'>7天'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'大于30天':'>30日'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'大于60天':'>60日'}, inplace=True)  # 改列名
Wanggebiao.rename(columns={'大于90天':'>90日'}, inplace=True)  # 改列名

#--------------------导出数据
print('导出数据')
with pd.ExcelWriter('2021年广西移动家宽投诉在途工单.xlsx') as writer:  # 写入结果为当前路径生成Excle表格文件
   Gd_Zb.to_excel(writer, sheet_name='当日归档明细', startcol=0, index=False, header=True)
   Gd_Pb.to_excel(writer, sheet_name='当日派单明细', startcol=0, index=False, header=True)
   Gd_Sj.to_excel(writer, sheet_name='当日在途明细', startcol=0, index=False, header=True)
   Dishi_Rb.to_excel(writer, sheet_name='地市日报', startcol=0, index=False, header=True)
   Quxianbiao.to_excel(writer, sheet_name='表3-区县', startcol=0, index=False, header=True)
   Wanggebiao.to_excel(writer, sheet_name='表4-网格', startcol=0, index=False, header=True)
print('导出数据完成结束')











