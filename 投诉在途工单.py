# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 09:44:43 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np


#Tabletime = datetime.datetime.strptime('2021-04-22','%Y-%m-%d') #---------------------手动输入日期
#print("打印时间" + Tabletime)
#-------手动输入时间
Time = '2021-04-27'

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


 #当日归档明细
 
Gd_Zb = Gd_Zb.fillna('')
Gd_Zb['安装地址'] = ''
Gd_Zb['投诉内容'] = ''

#---------剔除问题解决单$和家亲
Gd_Zb = Gd_Zb[(Gd_Zb.工单类型 != '问题解决单')&(Gd_Zb.工单类型 !='和家亲')]


Gd_Pb = Gd_Pb.fillna('')
Gd_Pb['安装地址'] = ''
Gd_Pb['投诉内容'] = ''

#---------剔除和睦$和家亲
Gd_Pb = Gd_Pb[(Gd_Pb.工单类型 != '问题解决单')&(Gd_Pb.工单类型 !='和家亲')]

Zt_Pb = Zt_Pb.fillna('')
Zt_Pb['安装地址'] = ''
Zt_Pb['投诉内容'] = ''

#---------剔除和睦$和家亲
Zt_Pb = Zt_Pb [(Zt_Pb .工单类型 != '问题解决单')&(Zt_Pb .工单类型 !='和家亲')]

#当日派单明细
   
#Gd_Pb = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/4-22-24eomsp.xlsx', inplace=True)#读取eomsp表

Sd = Gd_Pb[['首次到单时间']]

#todaytime = datetime.datetime.strptime('2021-04-27','%Y-%m-%d') #---------------------手动输入日期
#todaytime = datetime.datetime(todaytime.year, todaytime.month, todaytime.day, 18, 00, 00)#设置当日为18点
Sd['时间差'] = (pd.to_datetime(todaytime)-pd.to_datetime(Sd.首次到单时间)).dt.days*24+(pd.to_datetime(todaytime)-pd.to_datetime(Sd.首次到单时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Sd['是否18点到达'] = Sd.时间差 > 0
Sd = Sd.drop(['首次到单时间','时间差'],axis=1)#删除列
Gd_Pb = pd.concat([Sd, Gd_Pb], axis=1)  # 插入第一列为标段信息---------当日派单明细


#当日在途明细
'''
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

# 8:00

Zaitu_Bdll = Zt_Pb[['责任代维班组']]
Zaitu_Bdll.rename(columns={'责任代维班组': '代维班组'}, inplace=True)  # 改列名


Zhongjian_Kl = Zj_Pb[['代维班组']]
Zhongjian_Dishi = Zj_Pb[['地市代维']]












'''for i in Zaitu_Bdll['代维班组']:
    for j in Zhongjian_Kl['代维班组']:
        
        if (i == j):
            a = {'地市代维':[i]}
            
        else:
            a= {'地市代维':[j + j]}'''

'''def readexcel(exceldir):
    os.chdir(exceldir)
    try:
        f1 =Zhongjian_Kl.open_workbook('sggwhu_stationinfo.xlsx')
    except:
        print('There is no excel name sggwhu_stationinfo.xlsx\n Please cheek ! \n')
    sheet = f1.sheet_by_index(0)             # 读sheet，这里取第一个
    rows  = sheet.nrows                      # 获得行数
    data  = [[] for i in range(rows)]        # 去掉表头，从第二行读数据
    for i in range(1, rows): 
        data[i-1] = sheet.row_values(i)[1:5] # 去掉序号，取四个数据
    #for var in data:
        #print(var)
    return data
            
readexcel(Zhongjian_Kl)''' 
ww=Zt_Pb[['工单流水号','广西时限']]
height,width = ww.shape
print(height,width,type(ww))
x = np.zeros((height,width))
for i in range(0,height):
	for j in range(1,width+1): #遍历的实际下标，即excel第一行
		x[i][j-1] = ww.ix[i,j]
print(x.shape)
print(x)





aaaaa'''
#Zt_Pb = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/4-22-24eomszt.xlsx', inplace=True)#读取eomszt表

#地市

Gd_Sj = Zt_Pb[['责任地市']]

#Gd_Sj = Zt_Pb
#Gd_Sj.rename(columns={'责任地市': '地市'}, inplace=True)  # 改列名
#Gd_Sj.ix[(Gd_Sj.责任地市==''),'地市'] = False
#Drztmx = Gd_Sj.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
#Drztmx.insert(0, Drztmx.pop(Drztmx.index('地市')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
#Gd_Sj = Gd_Sj[Drztmx]

#Gd_Sj= Zt_Pb[(Zt_Pb.责任地市 != '')]



#with pd.ExcelWriter('123.xlsx') as writer:  # 写入结果为当前路径生成Excle表格文件
# Gd_Sj.to_excel(writer, sheet_name='1', startcol=0, index=False, header=True)

Gd_Sj.rename(columns={'责任地市': '地市'}, inplace=True)  # 改列名

#工单开始时间+是否超过24小时

Sj=Zt_Pb[['首次到单时间']]
Sj.rename(columns={'首次到单时间': '工单开始时间'}, inplace=True)  # 改列名
Gd_Sj = Sj[['工单开始时间']]
Gd_Sj[' '] = ''
#Sz_Time = datetime.datetime.strptime('2021-04-27','%Y-%m-%d') #---------------------------手动输入日期
#Sz_Time = datetime.datetime(Sz_Time.year, Sz_Time.month, Sz_Time.day, 23, 59, 59)#设置当日为18点
Gd_Sj['时间差'] = (pd.to_datetime(Sz_Time)-pd.to_datetime(Gd_Sj.工单开始时间)).dt.days*24+(pd.to_datetime(Sz_Time)-pd.to_datetime(Gd_Sj.工单开始时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Gd_Sj['是否超24小时'] = Gd_Sj.时间差 > 24

#是否多次驳回

Bh = Zt_Pb[['重派次数']]
Gd_Sj['是否多次驳回'] = Bh.重派次数 > 2

#区域类型

Quyu = Zt_Pb[['区域类型']]
Gd_Sj['区域类型'] = Quyu.区域类型

#区域铁通

Dw_Bz = Zt_Pb[['责任代维班组']]
Dw_Bz.rename(columns={'责任代维班组': '代维班组'}, inplace=True)  # 改列名

#读取中间数据

#Zj_Pb = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/自定义数据1.xlsx',sheet_name='中间数据', inplace=True) #读取中间数据
Zj_sj = Zj_Pb[['代维班组','地市代维一']]
ZTdata1 = pd.merge(Dw_Bz, Zj_sj, on=['代维班组'], how='left')  # 匹配代维信息
ZTdata1.rename(columns={'地市代维一': '二十点整'}, inplace=True)  # 改列名
Gd_Sj['20:00:00'] =ZTdata1.二十点整

#地市

Gd_Sj['地市'] = Zt_Pb[['责任地市']]
#Gd_Sj.地市 = Gd_Sj['地市'].replac('','南宁')





#地市代维（最终）

Dishi = Gd_Sj[['地市']]
Es = Gd_Sj[['20:00:00']]
Es.rename(columns={'20:00:00': '地市'}, inplace=True)  # 改列名
Dishidata = pd.merge(Es,Dishi, on=['地市'], how='left')  # 匹配代维信息
Gd_Sj['地市代维（最终）'] = Dishidata.地市

#代维最终

#Dw_Zz = Zj_Pb[['地市代维','简称']]
#Dw_Zz.rename(columns={'地市代维': '地市'}, inplace=True)  # 改列名
#Dw_Zz = pd.merge(Dw_Zz,Es, on=['地市'], how='left')  # 匹配代维信息
#Gd_Sj['代维（最终）'] = Dw_Zz.简称

#tag1

Chaoshi = Gd_Sj[['时间差']]
Chaoshi['tag1']= ''
Chaoshi['tag1'][Chaoshi['时间差'] >= 100] = '超100小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Chaoshi['tag1'][(Chaoshi['时间差'] < 100)&(Chaoshi['时间差'] >= 48)] = '超48-100小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Chaoshi['tag1'][(Chaoshi['时间差'] < 48)&(Chaoshi['时间差'] >= 24)] = '超24-48小时未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
 #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
#结尾需要的数据为'Df_final3'命名！
Gd_Sj['tag1'] = Chaoshi.tag1

#是否9月前工单

#Ju_Time = datetime.datetime.strptime('2020-08-31','%Y-%m-%d') #---------------------------------手动输入日期
#Ju_Time = datetime.datetime(Ju_Time.year, Ju_Time.month, Ju_Time.day, 23, 59, 59)#设置当日为18点
Gd_Sj['是否9月前工单']  = (pd.to_datetime(Ju_Time)-pd.to_datetime(Gd_Sj.工单开始时间)).dt.days*24+(pd.to_datetime(Ju_Time)-pd.to_datetime(Gd_Sj.工单开始时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Gd_Sj['是否9月前工单'] = Gd_Sj.是否9月前工单 > 0

#是否待质检

Zhijian = Zt_Pb[['工单状态']]
Zhijian['是否待质检'] = Zhijian.工单状态 == '质检'
Gd_Sj['是否待质检'] = Zhijian.是否待质检

#催单次数 +被客服催单

Cuidan = Zt_Pb[['催办次数','客服催办次数']]
Cuidan = Cuidan.fillna(0) #把nan填充为0
Cuidan['催单次数'] = Cuidan['催办次数'] + Cuidan['客服催办次数']
Cuidan['被客服催单'] = Cuidan.客服催办次数 - 1 >= 0
Gd_Sj['催单次数'] = Cuidan.催单次数
Gd_Sj['被客服催单'] = Cuidan.被客服催单

#是否超时

Chaoshi1 = Zt_Pb[['广西时限']]
#Chaoshi2 = datetime.datetime.strptime('2021-04-27','%Y-%m-%d') #-------------------------------手动输入日期
#Chaoshi2 = datetime.datetime(Chaoshi2.year, Chaoshi2.month, Chaoshi2.day, 00, 00, 00)#设置当日为18点
Chaoshi1['是否超时']  = (pd.to_datetime(Chaoshi2)-pd.to_datetime(Chaoshi1.广西时限)).dt.days*24+(pd.to_datetime(Chaoshi2)-pd.to_datetime(Chaoshi1.广西时限)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Gd_Sj['是否超时'] = Chaoshi1.是否超时

#是否要剔除


Tichu1 = Zt_Pb[['工单类型','工单状态']]

Tichu1.ix[(Tichu1.工单类型==''),'是否要剔除'] = False

Tichu1.ix[(Tichu1.工单类型=='家庭宽带'),'是否要剔除'] = False

Tichu1.ix[(Tichu1.工单状态=='待归档'),'是否要剔除'] = True

Tichu1['是否要剔除'] =Tichu1.是否要剔除.fillna(False)
#Tichu1 = Tichu1[Tichu1.是否要剔除 == 'True']
Gd_Sj['是否要剔除'] = Tichu1.是否要剔除

#Gd_Pb = Gd_Pb[(Gd_Pb.工单类型 != '问题解决单')&(Gd_Pb.工单类型 !='和家亲')]



#---------------------------8:00
#Bo =  Zt_Pb[['责任区县']]
#G = Zj_Pb[['维护区域二']]
#H = Zj_Pb[['代维一']]
#G ['代维一'] = H[['代维一']] 
#G.rename(columns={'维护区域二': '责任区县'}, inplace=True)  # 改列名
#G = pd.merge(Bo ,G[['责任区县','代维一']], on=['代维一'], how='left')  # 匹配代维信息







Zeren_Qx = Zt_Pb[['责任区县']]
Zeren_Qx['维护区域二'] = Zj_Pb[['维护区域二']]
Zeren_Qx['代维一'] = Zj_Pb[['代维一']]
Zeren_Qx['20:00:00'] = Gd_Sj[['20:00:00']]
Zeren_Qx['地市代维一'] =Zj_Pb[['地市代维一']]
Zeren_Qx['地市二'] = Zj_Pb[['地市二']]
Zeren_Qx['简称一'] = Zj_Pb[['简称一']]
Zeren_Qx['客服最后重派时间'] = Zt_Pb[['客服最后重派时间']]
Zeren_Qx['代维公司'] = Zj_Pb[['代维公司']]
Zeren_Qx['简称'] = Zj_Pb[['简称']]
Zeren_Qx.rename(columns={'20:00:00': '晚上八点'}, inplace=True)  # 改列名
Zeren_Qx.ix[(Zeren_Qx.责任区县 == Zeren_Qx.维护区域二),'八点'] = Zeren_Qx.代维一
Zeren_Qx.ix[(Zeren_Qx.晚上八点 == Zeren_Qx.地市代维一),'八点'] = Zeren_Qx.简称一
Zeren_Qx.ix[(Zeren_Qx.代维公司 == Zeren_Qx.简称),'八点'] = Zeren_Qx.简称
Zeren_Qx = Zeren_Qx.fillna('铁通 ') 
Gd_Sj['八点'] = Zeren_Qx.八点
#Zeren_Qx = Zeren_Qx.fillna('铁通 ') 

#代维（最终）

#中间数据!$L:$N,3,FALSE

Daiwei_Zz = Gd_Sj[['20:00:00']]
Daiwei_Zz.rename(columns={'20:00:00': '二十点'}, inplace=True)  # 改列名
Daiwei_Zz['地市代维一'] = Zj_Pb[['地市代维一']]
Daiwei_Zz['地市二'] = Zj_Pb[['地市二']]
Daiwei_Zz['简称一'] = Zj_Pb[['简称一']]

# A2中间数据P-Q

Daiwei_Zz['地市'] = Gd_Sj[['地市']]
Daiwei_Zz['地市三'] = Zj_Pb[['地市三']]
Daiwei_Zz['代维二'] = Zj_Pb[['代维二']]
Daiwei_Zz.ix[(Daiwei_Zz.二十点 == Daiwei_Zz.地市代维一),'代维最终'] = Daiwei_Zz.简称一
Daiwei_Zz.ix[(Daiwei_Zz.地市 == Daiwei_Zz.地市三),'代维最终'] = Daiwei_Zz.代维二
Daiwei_Zz = Daiwei_Zz.fillna('铁通 ') 
Gd_Sj['代维最终'] = Daiwei_Zz.代维最终



Gd_Sj = Gd_Sj[['地市','是否超24小时','是否多次驳回','工单开始时间','区域类型','地市代维（最终）','代维最终','八点','20:00:00','时间差','tag1','是否9月前工单','是否待质检','是否要剔除','催单次数','被客服催单','是否超时']]
Gd_Sj.rename(columns={'代维最终': '代维(最终)'}, inplace=True)  # 改列名
Gd_Sj.rename(columns={'八点': '8:00:00'}, inplace=True)  # 改列名
#Gd_Sj.rename(columns={'时间差': '2021/04/11 24:00:00'}, inplace=True)  # 改列名
Gd_Sj['工单流水号'] = Zt_Pb[['工单流水号']]
Gd_Sj=pd.merge(Gd_Sj, Zt_Pb, on=['工单流水号'], how='left')
Drztmx = Gd_Sj.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
Drztmx.insert(18, Drztmx.pop(Drztmx.index('工单流水号')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
Gd_Sj = Gd_Sj[Drztmx]      #排列后数据。

#---------------------------新加的代码

Gd_Sj = Gd_Sj [(Gd_Sj .地市 != '')]


#-----------------------------------------------地市日报

Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

#当日归档

Temp1 = Gd_Zb.groupby(['责任地市']).size().reset_index(name='当日归档')
Temp1 = Temp1.append([{'责任地市':'全区','当日归档':Temp1.apply(lambda x:x.sum()).当日归档}],ignore_index=True)
Temp1.rename(columns={'责任地市': '地市'}, inplace=True) #改列名
Temp1 = pd.merge(Dishi, Temp1, on=['地市'], how='left') #拼接


#当日排单

Temp2 = Gd_Pb.groupby(['责任地市']).size().reset_index(name='当日派单')
Temp2 = Temp2.append([{'责任地市':'全区','当日派单':Temp2.apply(lambda x:x.sum()).当日派单}],ignore_index=True)
Temp2.rename(columns={'责任地市': '地市'}, inplace=True) #改列名

#其中18点后派单

Temp3 = Gd_Pb[Gd_Pb.是否18点到达 == False].groupby(['责任地市']).size().reset_index(name='其中18点后派单')
Temp3 = Temp3.append([{'责任地市':'全区','其中18点后派单':Temp3.apply(lambda x:x.sum()).其中18点后派单}],ignore_index=True)
Temp3.rename(columns={'责任地市': '地市'}, inplace=True) #改列名

#导入自定义数据

#Zdysj = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/自定义数据1.xlsx', sheet_name='地市日报自定义数据',inplace=True) #读取自定义数据
Zdysj.分母 = Zdysj.分母.fillna('')
Zdysj.归档总量 = Zdysj.归档总量.fillna('')

#昨日在途

Temp4 = Dishi[['地市']]
Temp4['昨日在途']= Zdysj[['昨日在途']]


#24点在途

Temp5 = Gd_Sj[Gd_Sj.是否要剔除 == False].groupby(['地市']).size().reset_index(name='二十四点在途')
Temp5 = Temp5.append([{'地市':'全区','二十四点在途':Temp5.apply(lambda x:x.sum()).二十四点在途}],ignore_index=True)
Temp5 = pd.merge(Dishi, Temp5, on=['地市'], how='left') #拼接

#压单比<=0.8

Temp5['压单比<=0.8'] = Temp5.二十四点在途 / Zdysj.日均归档量
Temp5['压单比<=0.8']=Temp5['压单比<=0.8'].round(decimals=2) #保留两位小数

#在途目标

Temp6 = Dishi[['地市']]
Temp6['日均归档量'] =Zdysj[['日均归档量']]
Temp6 ['在途目标']= Temp6.日均归档量 * 0.8
Temp6.在途目标=Temp6.在途目标 // 1 #保留0位小数
Temp6 = Temp6.drop(['日均归档量'],axis=1)#删除列

#在途目标差值

Temp7= Dishi[['地市']]
Temp7['在途目标'] =Temp6[['在途目标']]
Temp7['在途目标差值']= Temp7['在途目标'] - Temp5['二十四点在途']
Temp7= Temp7.drop(['在途目标'],axis=1)#删除列

#超长超时工单

#超过48小时

Temp8 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['地市']).size().reset_index(name='超过四十八小时')
Temp8 = Temp8.append([{'地市':'全区','超过四十八小时':Temp8.apply(lambda x:x.sum()).超过四十八小时}],ignore_index=True)
Temp8.超过四十八小时 = Temp8.超过四十八小时.fillna(0)
Temp8= pd.merge(Dishi, Temp8, on=['地市'], how='left') #拼接

#超过7日

Temp9 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['地市']).size().reset_index(name='超过七日')
Temp9 = Temp9.append([{'地市':'全区','超过七日':Temp9.apply(lambda x:x.sum()).超过七日}],ignore_index=True)
Temp9.超过七日 = Temp9.超过七日.fillna(0)

#超过15日

Temp10 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['地市']).size().reset_index(name='超过十五日')
Temp10= Temp10.append([{'地市':'全区','超过十五日':Temp10.apply(lambda x:x.sum()).超过十五日}],ignore_index=True)
Temp10.超过十五日 = Temp10.超过十五日.fillna(0)

#超过30日

Temp11 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['地市']).size().reset_index(name='超过三十日')
Temp11= Temp11.append([{'地市':'全区','超过三十日':Temp11.apply(lambda x:x.sum()).超过三十日}],ignore_index=True)
Temp11.超过三十日 = Temp11.超过三十日.fillna(0)

#超过60日

Temp12 = Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['地市']).size().reset_index(name='超过六十日')
Temp12 = Temp12.append([{'地市':'全区','超过六十日':Temp12.apply(lambda x:x.sum()).超过六十日}],ignore_index=True)
Temp12.超过六十日 = Temp12.超过六十日.fillna(0)
#超48小时占比

Temp13 = Dishi[['地市']]
Temp13 ['二十四点在途']= Temp5[['二十四点在途']]
Temp13['超四十八小时占比'] = Temp8.超过四十八小时 / Temp13.二十四点在途
Temp13['超四十八小时占比']=Temp13['超四十八小时占比'].apply(lambda x: '%.2f%%' % (x*100))
#Temp13= Temp13.drop(['二十四点在途'],axis=1)#删除列

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

#-------------------------表3-区县


#Zd_Qx = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/自定义数据1.xlsx',sheet_name='区县表自定义数据' , inplace=True) #读取自定义数据
Zd_Qx.归档总量 = Zd_Qx.归档总量.fillna('')
#Zd_Qx.区县分公司 = Zd_Qx.区县分公司.fillna('')
#----------------地市

#Dishi2 = Zd_Qx[['地市']]

#----------------区县分公司

#Qx_Fgs = Zd_Qx[['区县分公司']]
Quxian = Zd_Qx[['地市']]
Quxian['区县分公司'] = Zd_Qx[['区县分公司']]

#-------当日完成

Drwc = Quxian
Zrqu =Gd_Zb.groupby(['责任区县']).size().reset_index(name='当日完成')
Zrqu.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Drwc = pd.merge(Drwc,Zrqu , on=['区县分公司'], how='left') #拼接
#Drwc= Drwc.drop(['区县分公司'],axis=1)#删除列
Drwc.当日完成 = Drwc.当日完成.fillna(0)

#-------当日投诉派单

Drtspd = Quxian
Zrqu1 =Gd_Pb.groupby(['责任区县']).size().reset_index(name='当日投诉派单')
Zrqu1.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Drtspd = pd.merge(Drtspd,Zrqu1 , on=['区县分公司'], how='left') #拼接
#Drtspd= Drtspd.drop(['区县分公司'],axis=1)#删除列
Drtspd.当日投诉派单 = Drtspd.当日投诉派单.fillna(0)

#---------其中18点后派单

Sbdhpd = Quxian
Zrqu2 =Gd_Pb[(Gd_Pb.是否18点到达 == False)].groupby(['责任区县']).size().reset_index(name='其中18点后派单')
Zrqu2.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Sbdhpd = pd.merge(Sbdhpd,Zrqu2 , on=['区县分公司'], how='left') #拼接
#Sbdhpd= Sbdhpd.drop(['区县分公司'],axis=1)#删除列
Sbdhpd.其中18点后派单 = Sbdhpd.其中18点后派单.fillna(0)

#----------昨日在途

Zrzt =  Quxian[['地市','区县分公司']]
Zrzt['昨日在途'] = Zd_Qx[['昨日在途']]
Zrzt = Zrzt[['地市','区县分公司','昨日在途']]
#Zrzt= Zrzt.drop(['区县分公司'],axis=1)#删除列

#----------------24点在途

Essdzt = Quxian
Zrqu3 =Gd_Sj[(Gd_Sj.是否要剔除 == False)].groupby(['责任区县']).size().reset_index(name='二十四点在途')
Zrqu3.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Essdzt = pd.merge(Essdzt,Zrqu3 , on=['区县分公司'], how='left') #拼接
Essdzt.二十四点在途 = Essdzt.二十四点在途.fillna(0)
#Essdzt=Essdzt.drop(['区县分公司'],axis=1)#删除列

#----------------压单比小于0.8

Ydbxyldb = Quxian[['地市','区县分公司']]
Ydbxyldb['压单比小于08'] = Essdzt.二十四点在途 / Zd_Qx.日均归档量


#Ydbxyldb = Ydbxyldb.drop(['区县分公司'],axis=1)#删除列

#-------------在途目标

Ztmb = Quxian[['地市','区县分公司']]
Ztmb['在途目标'] = Zd_Qx.日均归档量 * 0.8
Ztmb.在途目标 = Ztmb.在途目标 // 1
#Ztmb= Ztmb.drop(['区县分公司'],axis=1)#删除列

#-----------在途目标差值

Ztmbcz = Quxian[['地市','区县分公司']]
Ztmbcz['在途目标差值'] = Essdzt.二十四点在途 - Ztmb.在途目标
#Ztmbcz= Ztmbcz.drop(['区县分公司'],axis=1)#删除列

#----------八月拍照未归档量

Bypzwgd = Quxian
Zrqu4 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.是否9月前工单 == True)].groupby(['责任区县']).size().reset_index(name='八月拍照未归档量')
Zrqu4.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Bypzwgd = pd.merge(Bypzwgd,Zrqu4 , on=['区县分公司'], how='left') #拼接
Bypzwgd.八月拍照未归档量 =Bypzwgd.八月拍照未归档量.fillna(0)
#Bypzwgd= Bypzwgd.drop(['区县分公司'],axis=1)#删除列

#-------->48小时

dysbxs =  Quxian
Zrqu5 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['责任区县']).size().reset_index(name='大于四十八小时')
Zrqu5.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
dysbxs = pd.merge(dysbxs,Zrqu5 , on=['区县分公司'], how='left') #拼接
dysbxs.大于四十八小时 = dysbxs.大于四十八小时.fillna(0)
#dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#-----大于7天

Dyqt =  Quxian
Zrqu6 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['责任区县']).size().reset_index(name='大于7天')
Zrqu6.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyqt = pd.merge(Dyqt,Zrqu6 , on=['区县分公司'], how='left') #拼接
Dyqt.大于7天 = Dyqt.大于7天.fillna(0)
#Dyqt= Dyqt.drop(['区县分公司'],axis=1)#删除列

#------大于15天

Dyswt =  Quxian
Zrqu7 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['责任区县']).size().reset_index(name='大于15天')
Zrqu7.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyswt = pd.merge(Dyswt,Zrqu7 , on=['区县分公司'], how='left') #拼接
Dyswt.大于15天 = Dyswt.大于15天.fillna(0)
#Dyswt= Dyswt.drop(['区县分公司'],axis=1)#删除列

#-----大于30天

Dysst =  Quxian
Zrqu8 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['责任区县']).size().reset_index(name='大于30天')
Zrqu8.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dysst = pd.merge(Dysst,Zrqu8 , on=['区县分公司'], how='left') #拼接
Dysst.大于30天 = Dysst.大于30天.fillna(0)
#Dysst= Dysst.drop(['区县分公司'],axis=1)#删除列

#------大于60天

Dylst =  Quxian
Zrqu9 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['责任区县']).size().reset_index(name='大于60天')
Zrqu9.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dylst = pd.merge(Dylst,Zrqu9 , on=['区县分公司'], how='left') #拼接
Dylst.大于60天 = Dylst.大于60天.fillna(0)
#Dylst = Dylst .drop(['区县分公司'],axis=1)#删除列

#------大于90天

Dyjst =  Quxian
Zrqu10 =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 2160)].groupby(['责任区县']).size().reset_index(name='大于90天')
Zrqu10.rename(columns={'责任区县': '区县分公司'}, inplace=True)  # 改列名
Dyjst = pd.merge(Dyjst,Zrqu10 , on=['区县分公司'], how='left') #拼接
Dyjst.大于90天 = Dyjst.大于90天.fillna(0)
#Dyjst= Dyjst.drop(['区县分公司'],axis=1)#删除列

#--------超48小时占比

Csbxszb =  Quxian[['地市','区县分公司']]
Csbxszb ['超48小时占比']= (dysbxs.大于四十八小时 / Essdzt.二十四点在途).apply(lambda x: '%.2f%%' % (x*100))
Csbxszb= Csbxszb.drop(['区县分公司'],axis=1)#删除列


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

#--------------------表4网格






#Zd_Wg = pd.read_excel('file:///D:/2021年广西移动家宽投诉在途工单/自定义数据1.xlsx',sheet_name = '网格表自定义数据', inplace=True) #读取自定网格义数据
#Zd_Wg.日均归档量 =Zd_Wg.日均归档量.round(decimals=1)#保留一位小数 
Zd_Wg.日期 =Zd_Wg.日期.fillna('')
Zd_Wg.归档总量 =Zd_Wg.归档总量.fillna('')
Wangge = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]

#------------------24点在途

#Wg_Essdzt =Gd_Sj[(Gd_Sj.是否要剔除 == False)].groupby(['责任区县']).size().reset_index(name='24点在途')
Wg_Essdzt1 =Gd_Sj[(Gd_Sj.是否要剔除 == False)].groupby(['网格名称']).size().reset_index(name='二十四点在途')
Wg_Essdzt1.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_Essdzt1 = pd.merge(Wangge,Wg_Essdzt1 , on=['所属网格'], how='left') #拼接
Wg_Essdzt1.二十四点在途 = Wg_Essdzt1.二十四点在途.fillna(0) #把nan填充为0

#-------------------压单比小于等于0.8

Wg_Ydbxyldb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ydbxyldb ['压单比小于等于零点八']= (Wg_Essdzt1.二十四点在途 / Zd_Wg.日均归档量).round(decimals=2) #保留两位小数
Wg_Ydbxyldb.压单比小于等于零点八 = Wg_Ydbxyldb.压单比小于等于零点八.fillna(0) #把nan填充为0

#----------------------在途目标

Wg_Ztmb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ztmb['在途目标'] = Zd_Wg.日均归档量 *0.8
Wg_Ztmb.在途目标 = Wg_Ztmb.在途目标 //1  #取整数部分


#-------------------在途目标差值

Wg_Ztmbcz = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Ztmbcz['在途目标差值'] = Wg_Ztmb.在途目标 - Wg_Essdzt1.二十四点在途

#------------------八月拍照未归档


Wg_Bypzwgd =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.是否9月前工单 == True)].groupby(['网格名称']).size().reset_index(name='八月拍照未归档')

Wg_Bypzwgd.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_Bypzwgd = pd.merge(Wangge,Wg_Bypzwgd , on=['所属网格'], how='left') #拼接
Wg_Bypzwgd.八月拍照未归档 = Wg_Bypzwgd.八月拍照未归档.fillna(0) #把nan填充为0

#------------------------ 大于48小时

Wg_dysbxs =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 48)].groupby(['网格名称']).size().reset_index(name='大于四十八小时')
Wg_dysbxs.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dysbxs = pd.merge(Wangge,Wg_dysbxs , on=['所属网格'], how='left') #拼接
Wg_dysbxs.大于四十八小时 =Wg_dysbxs.大于四十八小时.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#-----大于7天

Wg_dyqt =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 168)].groupby(['网格名称']).size().reset_index(name='大于7天')
Wg_dyqt.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyqt = pd.merge(Wangge,Wg_dyqt , on=['所属网格'], how='left') #拼接
Wg_dyqt.大于7天 =Wg_dyqt.大于7天.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#------大于15天

Wg_dyswt =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 360)].groupby(['网格名称']).size().reset_index(name='大于15天')
Wg_dyswt.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyswt = pd.merge(Wangge,Wg_dyswt , on=['所属网格'], how='left') #拼接
Wg_dyswt.大于15天 = Wg_dyswt.大于15天.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#-----大于30天

Wg_dysst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 720)].groupby(['网格名称']).size().reset_index(name='大于30天')
Wg_dysst.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dysst = pd.merge(Wangge,Wg_dysst , on=['所属网格'], how='left') #拼接
Wg_dysst.大于30天 = Wg_dysst.大于30天.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#------大于60天

Wg_dylst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 1440)].groupby(['网格名称']).size().reset_index(name='大于60天')
Wg_dylst.rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dylst = pd.merge(Wangge,Wg_dylst , on=['所属网格'], how='left') #拼接
Wg_dylst.大于60天= Wg_dylst.大于60天.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#------大于90天

Wg_dyjst =Gd_Sj[(Gd_Sj.是否要剔除 == False)&(Gd_Sj.时间差 > 2160)].groupby(['网格名称']).size().reset_index(name='大于90天')
Wg_dyjst .rename(columns={'网格名称': '所属网格'}, inplace=True)  # 改列名
Wg_dyjst  = pd.merge(Wangge,Wg_dyjst, on=['所属网格'], how='left') #拼接
Wg_dyjst.大于90天= Wg_dyjst.大于90天.fillna(0) #把nan填充为0
#Wg_dysbxs= dysbxs.drop(['区县分公司'],axis=1)#删除列

#--------------------超48小时占比

Wg_Csbxszb = Zd_Wg[['日期','地市','区县分公司','所属网格','昨日在途']]
Wg_Csbxszb['超48小时占比'] = (Wg_dysbxs.大于四十八小时 / Wg_Essdzt1.二十四点在途).apply(lambda x: '%.2f%%' % (x*100))

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

#--------------------修改列名

#--------------修改当日派单明细列名

Gd_Pb.rename(columns={'是否18点到达':'2021/4/11 18:00:00'}, inplace=True)  # 改列名

#-------------修改当日在途明细列名

Gd_Sj.rename(columns={'区域类型_x':'区域类型'}, inplace=True)  # 改列名
Gd_Sj.rename(columns={'时间差':'2021/4/11 24:00:00'}, inplace=True)  # 改列名
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

with pd.ExcelWriter('2021年广西移动家宽投诉在途工单.xlsx') as writer:  # 写入结果为当前路径生成Excle表格文件
   Gd_Zb.to_excel(writer, sheet_name='当日归档明细', startcol=0, index=False, header=True)
   Gd_Pb.to_excel(writer, sheet_name='当日派单明细', startcol=0, index=False, header=True)
   Gd_Sj.to_excel(writer, sheet_name='当日在途明细', startcol=0, index=False, header=True)
   Dishi_Rb.to_excel(writer, sheet_name='地市日报', startcol=0, index=False, header=True)
   Quxianbiao.to_excel(writer, sheet_name='表3-区县', startcol=0, index=False, header=True)
   Wanggebiao.to_excel(writer, sheet_name='表4-网格', startcol=0, index=False, header=True)












