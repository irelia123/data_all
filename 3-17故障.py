# -*- coding: utf-8 -*-
"""
Created on Tue Mar 23 16:48:43 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np

######################################################################## 初始化

Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','不详','广西']},
    pd.Index(range(16)))


start0 = time.clock()  # 开始计时
print('*******************开始计算数障表')

absolve =  400  ###减免系数   设置为 0 不减免
'''todaytime = datetime.datetime.strptime('2021-3-17','%Y-%m-%d') #手动输入日期
todaytime = datetime.datetime(todaytime.year, todaytime.month, todaytime.day, 00, 00, 00)  #调整至当天零点 方便后续计算

first_day = datetime.datetime(todaytime.year, todaytime.month, 17, 00, 00, 00)  # datetime类型 2019-09-01 00:00:00  取当月1日
today = str(todaytime.month) + '月' + todaytime.strftime('%d') + '日'  # 日期转换str：9月01日'''


if (absolve > 0):
    print("自动减免机制--启用")
    
'''HSx = pd.read_excel('./17G故障数据表.xlsx')
hxr = pd.read_excel('./中间数据.xlsx')
hxr = hxr.fillna('')

Xiangdan = HSx[['地市']]
Xiangdan['故障历时（小时）'] = HSx[['回复历时（小时）']]
Xiangdan.rename(columns={'故障历时（小时）': '故障小时'}, inplace=True)  # 改列名
Xiangdan['8:00']  = (pd.to_datetime(HSx.派单时间) - pd.to_datetime(HSx.T2最后回单时间)).dt.days*24 + (pd.to_datetime(HSx.派单时间) - pd.to_datetime(HSx.T2最后回单时间)).dt.seconds/3600 + 8

Xiangdan['是否超时'] = HSx[['是否超时']]
Xiangdan['是否延期'] = HSx[['是否延期']]
Xiangdan['影响用户数'] = HSx[['影响用户数']]
Xiangdan['22:00'] = Xiangdan.故障小时 * Xiangdan.影响用户数

#Xiangdan.ix[(HSx['质检情况'].map(len) == 1),'质检结果'] = '合格'  
Xiangdan['质检结果'] = (HSx.质检情况).map(len) < 1 #返回对应的类型len
Xiangdan['质检结果'][Xiangdan['质检结果'] ==False ] = '合格'   '''  

HSx1 = pd.read_excel('G故障工单统计-20210317.xlsx')
Temp1 = HSx1.groupby(['地市']).size().reset_index(name='工单总数')

#*****************新增一列新数据
Temp1['广西总数'] = Temp1.groupby('地市')['工单总数'].shift()
#***********************赋值=工单总数
Temp1['广西总数'] = Temp1.apply(lambda x: x.sum()).工单总数
#Temp1['活跃用户量'] = Temp1.groupby('地市')['工单总数'].shift()
#*********************************求和
Temp1 = Temp1.append([{'地市':'广西','工单总数':Temp1.apply(lambda x: x.sum()).工单总数, }], ignore_index=True)
###################000000000计算百分比
Temp1['占比'] = Temp1.工单总数 / Temp1.广西总数
Temp1=Temp1.fillna(0)
Temp1['占比']=Temp1['占比'].apply(lambda x: '%.2f%%' % (x*100))

######################删除广西总数那一列，删列
Temp1.drop(['广西总数'],axis=1,inplace=True) 
#######################及时工单数
Temp2 = HSx1[HSx1.是否超时=='否'].groupby(['地市']).size().reset_index(name='及时工单数')
Temp2 = Temp2.append([{'地市':'广西','及时工单数':Temp2.apply(lambda x: x.sum()).及时工单数, }], ignore_index=True)

#00000000000000000000000拼接列表
Temp1=pd.merge(Temp1,Temp2,on=['地市'],how='left')

Temp1['及时率'] = Temp1.及时工单数 / Temp1.工单总数  # 这里是计算达标率
Temp1['及时率']=Temp1['及时率'].apply(lambda x: '%.2f%%' % (x*100))
#########################延期工单数
Temp3 = HSx1[HSx1.是否延期=='是'].groupby(['地市']).size().reset_index(name='延期工单数')
Temp3 = Temp3.append([{'地市':'广西','延期工单数':Temp3.apply(lambda x: x.sum()).延期工单数, }], ignore_index=True)
Temp1=pd.merge(Temp1,Temp3,on=['地市'],how='left') 
#########################把nan转为0
Temp1=Temp1.fillna(0)

Temp1['挂起率'] = Temp1.延期工单数 / Temp1.工单总数  # 这里是计算达标率
Temp1['挂起率']=Temp1['挂起率'].apply(lambda x: '%.2f%%' % (x*100))

###########################筛选三列

Yonghushu= HSx1[['地市','影响用户数','是否列入统计']]
Yonghushu = Yonghushu[Yonghushu.是否列入统计==True]#.groupby(['地市']).size().reset_index(name='影响用户数')
#####################筛选符合条件所有地市
Yonghushu = Yonghushu.append([{'地市':'广西','影响用户数':Yonghushu.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
#######################同一地市累加
Yonghushu = pd.DataFrame(Yonghushu.groupby('地市')['影响用户数'].sum()).reset_index()
Yonghushu=pd.merge(Dishi,Yonghushu,on=['地市'],how='left') 
Temp1=pd.merge(Temp1,Yonghushu,on=['地市'],how='left') 

######################故障总时长
HSx1.rename(columns={'故障历时（小时）': '故障总时长'}, inplace=True)  # 改列名
Zongshichang= HSx1[['地市','故障总时长','是否列入统计']]
Zongshichang = Zongshichang[Zongshichang.是否列入统计 ==True]
Zongshichang = Zongshichang.append([{'地市':'广西','故障总时长':Zongshichang.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
Zongshichang = pd.DataFrame(Zongshichang.groupby('地市')['故障总时长'].sum()).reset_index()

#0000000000000000000000000故障平均时长
Pingjunshi= HSx1[['地市','故障总时长','是否列入统计']]
Pingjunshi = Pingjunshi[Pingjunshi.是否列入统计 ==True].groupby(['地市']).size().reset_index(name='是否列入统计')
Pingjunshi = Pingjunshi.append([{'地市':'广西','是否列入统计':Pingjunshi.apply(lambda x: x.sum()).是否列入统计}], ignore_index=True)
Zongshichang=pd.merge(Zongshichang,Pingjunshi,on=['地市'],how='left') 
Zongshichang['故障平均时长'] = Zongshichang.故障总时长 / Zongshichang.是否列入统计
Temp1=pd.merge(Temp1,Zongshichang,on=['地市'],how='left') 
Temp1['故障平均时长'] = Temp1['故障平均时长'].round(decimals=2)#保留两位小数
Temp1.drop(['是否列入统计'],axis=1,inplace=True) #删除多余列

#00000000000000000000000000质检合格工单数
Hegegong = HSx1[(HSx1.质检情况=='合格')&(HSx1.是否列入统计==True)].groupby(['地市']).size().reset_index(name='质检合格工单数')
Hegegong = Hegegong.append([{'地市':'广西','质检合格工单数':Hegegong.apply(lambda x: x.sum()).质检合格工单数}], ignore_index=True)
Temp1=pd.merge(Temp1,Hegegong,on=['地市'],how='left') 
#0000000000000000000000000000000000影响用户工单数
Yonghugong = HSx1[HSx1.是否列入统计==True].groupby(['地市']).size().reset_index(name='影响用户工单数')
Yonghugong = Yonghugong.append([{'地市':'广西','影响用户工单数':Yonghugong.apply(lambda x: x.sum()).影响用户工单数}], ignore_index=True)
Temp1=pd.merge(Temp1,Yonghugong,on=['地市'],how='left') 
#000000000000000000000000000000000000000000质检合格率
Temp1['质检合格率'] = Hegegong.质检合格工单数 / Yonghugong.影响用户工单数   
Temp1['质检合格率']=Temp1['质检合格率'].apply(lambda x: '%.2f%%' % (x*100))
#00000000000000000000000000000000000000用户数 * 历时
Yonghulishi= HSx1[['地市','故障总时长','影响用户数','是否列入统计']]
Yonghulishi = Yonghulishi[(Yonghulishi.是否列入统计==True)]
Yonghulishi['求和用户数乘历时'] = Yonghulishi.故障总时长 * Yonghulishi.影响用户数

Yonghulishi=Yonghulishi.fillna(0)
Yonghulishi = Yonghulishi.append([{'地市':'广西','求和用户数乘历时':Yonghulishi.apply(lambda x: x.sum()).求和用户数乘历时}], ignore_index=True)
Yonghulishi = pd.DataFrame(Yonghulishi.groupby('地市')['求和用户数乘历时'].sum()).reset_index()
Temp1=pd.merge(Temp1,Yonghulishi,on=['地市'],how='left')

#000000000000000000000000活跃用户量 

Zhongbiao = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','不详','广西']},
    pd.Index(range(16)))
Zhongbiao=pd.merge(Zhongbiao,Temp1,on=['地市'],how='left') 
Zhongbiao=Zhongbiao.fillna(0)
Huoyue = Dishi
Huoyue['活跃用户量'] = pd.Series([1238253,802596,754641,676885,448619,473272,455065,397020,362809,332263,262609,290167,233946,145052,2318])
Huoyue = Huoyue.append([{'地市':'广西','活跃用户量':Huoyue.apply(lambda x: x.sum()).活跃用户量}], ignore_index=True)
Huoyue = pd.DataFrame(Huoyue.groupby('地市')['活跃用户量'].sum()).reset_index()
Zhongbiao=pd.merge(Zhongbiao,Huoyue,on=['地市'],how='left') 
 
#000000000000000000000000000000000000用户/库存量
Zhongbiao['用户/库存量'] = Zhongbiao.影响用户数 / Zhongbiao.活跃用户量
Zhongbiao['用户/库存量'] = Zhongbiao['用户/库存量'].round(decimals=2)#保留两位小数
#00000000000000000家宽客户月均中断时长
Zhongbiao['家宽客户月均中断时长'] = Zhongbiao.求和用户数乘历时 / Zhongbiao.活跃用户量

#000000000000000000000000影响用户及时工单数
Yonghugongjishig = HSx1[(HSx1.是否超时=='否')&(HSx1.是否列入统计==True)].groupby(['地市']).size().reset_index(name='影响用户及时工单数')
Yonghugongjishig = Yonghugongjishig.append([{'地市':'广西','影响用户及时工单数':Yonghugongjishig.apply(lambda x: x.sum()).影响用户及时工单数}], ignore_index=True)
Zhongbiao=pd.merge(Zhongbiao,Yonghugongjishig,on=['地市'],how='left') 

# 000000000000000000000000000000影响用户及时率
Zhongbiao['影响用户及时率'] = Zhongbiao.影响用户及时工单数 / Zhongbiao.影响用户工单数
Zhongbiao=Zhongbiao.fillna(0)
Zhongbiao['影响用户及时率']=Zhongbiao['影响用户及时率'].apply(lambda x: '%.2f%%' % (x*100))

#0000000000000000000000新增一列空列
Zhongbiao['农村'] = ' '

#0000000000000000000000农村影响用户工单数
NCYonghugong = HSx1[(HSx1.是否列入统计==True)&(HSx1.区域属性=='农村')].groupby(['地市']).size().reset_index(name='影响用户工单数')
NCYonghugong = NCYonghugong.append([{'地市':'广西','影响用户工单数':NCYonghugong.apply(lambda x: x.sum()).影响用户工单数}], ignore_index=True)
Zhongbiao=pd.merge(Zhongbiao,NCYonghugong,on=['地市'],how='left') 
Zhongbiao.rename(columns={'影响用户工单数_y': '农村影响用户工单数'}, inplace=True)  # 改列名

#0000000000000000000000农村影响用户及时工单数
NCYonghugongjishig = HSx1[(HSx1.是否超时=='否')&(HSx1.是否列入统计==True)&(HSx1.区域属性=='农村')].groupby(['地市']).size().reset_index(name='影响用户及时工单数')
NCYonghugongjishig = NCYonghugongjishig.append([{'地市':'广西','影响用户及时工单数':NCYonghugongjishig.apply(lambda x: x.sum()).影响用户及时工单数}], ignore_index=True)
NCYonghugong=pd.merge(NCYonghugong,NCYonghugongjishig,on=['地市'],how='left') 

#00000000000000000000000农村影响用户及时率
NCYonghugong['影响用户及时率'] = NCYonghugong.影响用户及时工单数 / NCYonghugong.影响用户工单数
NCYonghugong['影响用户及时率']=NCYonghugong['影响用户及时率'].apply(lambda x: '%.2f%%' % (x*100))
Zhongbiao=pd.merge(Zhongbiao,NCYonghugong,on=['地市'],how='left') 
Zhongbiao.rename(columns={'影响用户及时率_y': '农村影响用户及时率'}, inplace=True)  # 改列名

Zhongbiao.drop(['影响用户工单数'],axis=1,inplace=True) 
Zhongbiao.drop(['影响用户及时工单数_y'],axis=1,inplace=True) 

#00000000000000000000000000农村影响用户量
NCYonghushu= HSx1[['地市','影响用户数','是否列入统计','区域属性']]
NCYonghushu= NCYonghushu[(NCYonghushu.区域属性=='农村')&(NCYonghushu.是否列入统计==True)]
NCYonghushu = NCYonghushu.append([{'地市':'广西','影响用户数':NCYonghushu.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
NCYonghushu = pd.DataFrame(NCYonghushu.groupby('地市')['影响用户数'].sum()).reset_index()
Zhongbiao=pd.merge(Zhongbiao,NCYonghushu,on=['地市'],how='left') 
Zhongbiao.rename(columns={'影响用户数_y': '农村影响用户数'}, inplace=True)  # 改列名

#000000000000000000000000000农村故障总时长
NCZongshichang= HSx1[['地市','故障总时长','是否列入统计','区域属性']]
NCZongshichang = NCZongshichang[(NCZongshichang.是否列入统计 ==True)&(NCZongshichang.区域属性 =='农村')]
NCZongshichang = NCZongshichang.append([{'地市':'广西','故障总时长':NCZongshichang.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
NCZongshichang = pd.DataFrame(NCZongshichang.groupby('地市')['故障总时长'].sum()).reset_index()

#0000000000000000000000000000农村故障平均时长
NCPingjunshi= HSx1[['地市','故障总时长','是否列入统计','区域属性']]
NCPingjunshi = NCPingjunshi[(NCPingjunshi.是否列入统计 ==True)&(NCPingjunshi.区域属性 =='农村')].groupby(['地市']).size().reset_index(name='是否列入统计')
NCPingjunshi = NCPingjunshi.append([{'地市':'广西','是否列入统计':NCPingjunshi.apply(lambda x: x.sum()).是否列入统计}], ignore_index=True)
NCZongshichang=pd.merge(NCZongshichang,NCPingjunshi,on=['地市'],how='left') 
NCZongshichang['农村故障平均时长'] = NCZongshichang.故障总时长 / NCZongshichang.是否列入统计
Zhongbiao=pd.merge(Zhongbiao,NCZongshichang,on=['地市'],how='left') 
Zhongbiao['农村故障平均时长'] = Zhongbiao['农村故障平均时长'].round(decimals=2)#保留两位小数
Zhongbiao=Zhongbiao.fillna(0)
Zhongbiao.drop(['是否列入统计'],axis=1,inplace=True) #删除多余列
Zhongbiao.drop(['故障总时长_y'],axis=1,inplace=True) #删除多余列

Zhongbiao[''] = ''

#00000000000000000000000000000筛选所有地市的故障总时长
Old= HSx1[['地市','故障总时长','中断点']]
Old = Old[(Old.中断点 =='OLT级别')]
Old = Old.append([{'地市':'广西','故障总时长':Old.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
Old = pd.DataFrame(Old.groupby('地市')['故障总时长'].sum()).reset_index()
 
#00000000000000000000000000筛选符合OLD级别的所有地市
Old1= HSx1[['地市','故障总时长','中断点']]
Old1 = Old1[Old1.中断点 =='OLT级别'].groupby(['地市']).size().reset_index(name='中断点')
Old1 = Old1.append([{'地市':'广西','中断点':Old1.apply(lambda x: x.sum()).中断点}], ignore_index=True)
Old1=pd.merge(Old1,Old,on=['地市'],how='left') 

#000000000000000000000000集合以上两项求OLD均值
Old1['OLT级别'] = Old1.故障总时长 / Old1.中断点
Zhongbiao=pd.merge(Zhongbiao,Old1,on=['地市'],how='left') 
Zhongbiao['OLT级别'] = Zhongbiao['OLT级别'].round(decimals=2)#保留两位小数

Zhongbiao.drop(['故障总时长'],axis=1,inplace=True) #删除多余列
Zhongbiao.drop(['中断点'],axis=1,inplace=True) #删除多余列

#000000000000000000000000筛选所有地市的故障总时长
Old2= HSx1[['地市','故障总时长','中断点']]
Old2 = Old2[(Old2.中断点 =='PON口级别')]
Old2 = Old2.append([{'地市':'广西','故障总时长':Old2.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
Old2 = pd.DataFrame(Old2.groupby('地市')['故障总时长'].sum()).reset_index()

#0000000000000000000000筛选符合PON级别的所有地市
Old3= HSx1[['地市','故障总时长','中断点']]
Old3 = Old3[Old3.中断点 =='PON口级别'].groupby(['地市']).size().reset_index(name='中断点')
Old3 = Old3.append([{'地市':'广西','中断点':Old3.apply(lambda x: x.sum()).中断点}], ignore_index=True)
Old3=pd.merge(Old3,Old2,on=['地市'],how='left') 

#000000000000000000000000000000集合以上两项求PON均值
Old3['PON口级别'] = Old3.故障总时长 / Old3.中断点
Zhongbiao=pd.merge(Zhongbiao,Old3,on=['地市'],how='left') 
Zhongbiao['PON口级别'] = Zhongbiao['PON口级别'].round(decimals=2)#保留两位小数
Zhongbiao.drop(['故障总时长'],axis=1,inplace=True) #删除多余列
Zhongbiao.drop(['中断点'],axis=1,inplace=True) #删除多余列

#0000000000000000000000筛选所有地市的故障总时长
Old3= HSx1[['地市','故障总时长','中断点']]
Old3 = Old3[(Old3.中断点 =='分光器级别')]
Old3 = Old3.append([{'地市':'广西','故障总时长':Old3.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
Old3 = pd.DataFrame(Old3.groupby('地市')['故障总时长'].sum()).reset_index()

#00000000000000000000000筛选符合分光器级别de所有地市
Old4= HSx1[['地市','故障总时长','中断点']]
Old4 = Old4[Old4.中断点 =='分光器级别'].groupby(['地市']).size().reset_index(name='中断点')
Old4 = Old4.append([{'地市':'广西','中断点':Old4.apply(lambda x: x.sum()).中断点}], ignore_index=True)
Old4=pd.merge(Old4,Old3,on=['地市'],how='left') 

#00000000000000000000000000集合以上两项求割接或其他的均值
Old4['分光器级别'] = Old4.故障总时长 / Old4.中断点
Zhongbiao=pd.merge(Zhongbiao,Old4,on=['地市'],how='left') 
Zhongbiao['分光器级别'] = Zhongbiao['分光器级别'].round(decimals=2)#保留两位小数
Zhongbiao.drop(['故障总时长'],axis=1,inplace=True) #删除多余列
Zhongbiao.drop(['中断点'],axis=1,inplace=True) #删除多余列
Zhongbiao=Zhongbiao.fillna(0)

Zhongbiao.rename(columns={'影响用户及时率_x': '影响用户及时率'}, inplace=True)  # 改列名
Zhongbiao.rename(columns={'影响用户及时工单数_x': '影响用户及时工单数'}, inplace=True)  # 改列名
Zhongbiao.rename(columns={'影响用户工单数_x': '影响用户工单数'}, inplace=True)  # 改列名
Zhongbiao.rename(columns={'故障总时长_x': '故障总时长'}, inplace=True)  # 改列名
Zhongbiao.rename(columns={'影响用户数_x': '影响用户数'}, inplace=True)  # 改列名

#00000000000000000000000重新编排表格顺序
Zhongbiao = Zhongbiao[['地市','工单总数','占比','及时工单数','及时率','延期工单数','挂起率','影响用户数','故障总时长','故障平均时长',
                       '质检合格工单数','质检合格率','求和用户数乘历时','用户/库存量','活跃用户量','家宽客户月均中断时长','影响用户工单数',
                       '影响用户及时工单数','影响用户及时率','农村','农村影响用户工单数','农村影响用户及时率','农村影响用户数',
                       '农村故障平均时长','','OLT级别','分光器级别','PON口级别']]

print('******************故障数据表  计算完毕')

#############################        故障原因类别     *************************************



#Dishi1 = pd.DataFrame(
  #  {'故障原因类别': ['告警自动清除', '客户侧原因', '设备故障', '数据配置问题', '网络割接升级', '网络拥塞', '线路故障', '其他','总计']},
  #  pd.Index(range(9)))
  
Guzhang = pd.read_excel('file:///D:/工程量/3-17G故障/故障原因类别.xlsx')  # 当日导出的在途清单 日期为昨日

print('*******************开始计算 故障原因类别')

#000000000000000000000数量1所有的条件值
Gaojinqingchu = HSx1[['故障原因类别','故障原因细分']]
Gaojinqingchu = Gaojinqingchu.groupby(['故障原因类别','故障原因细分']).size().reset_index(name='数量1')
Gaojinqingchu = Gaojinqingchu.append([{'故障原因类别':'总计','数量1':Gaojinqingchu.apply(lambda x: x.sum()).数量1}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Gaojinqingchu,on=['故障原因类别','故障原因细分'],how ='left')
Gaojinqingchu = Guzhang
Gaojinqingchu = Gaojinqingchu.append([{'故障原因类别':'总计','数量1':Gaojinqingchu.apply(lambda x: x.sum()).数量1}], ignore_index=True)

#000000000000000000000 取指定行 列的某个数值
z=Zhongbiao.iloc[15,1]
#0000000000000000000新增新的一行总计
Gaojinqingchu.loc[len(Gaojinqingchu)] = ' '
Gaojinqingchu.loc[33,'故障原因类别'] = '总计1'
Gaojinqingchu.loc[33,'数量1'] = z #赋值
Gaojinqingchu=Gaojinqingchu.fillna(0)

#0000000000000000000000相减得到其他的数量
Gaojinqingchu.loc[31,'数量1'] = z - Gaojinqingchu.iloc[32,2] 

Gaojinqingchu['占比1'] = Gaojinqingchu.数量1 / z
Gaojinqingchu['占比1']=Gaojinqingchu['占比1'].apply(lambda x: '%.2f%%' % (x*100))

#00000000000000000000000000耗时
Haoshi= HSx1[['故障原因类别','故障总时长','故障原因细分']]
Haoshi = Haoshi.groupby(['故障原因类别','故障原因细分']).size().reset_index(name='耗时1')

#000000000000000000000000000000故障总时长
Haoshi1= HSx1[['故障原因类别','故障原因细分','故障总时长']]
Haoshi1 = pd.DataFrame(Haoshi1.groupby('故障原因细分')['故障总时长'].sum()).reset_index()
Haoshi=pd.merge(Haoshi,Haoshi1,on=['故障原因细分'],how='left') 

#000000000000000000000000000000平均耗时
Haoshi['耗时'] = Haoshi.故障总时长 / Haoshi.耗时1
Haoshi['耗时'] = Haoshi['耗时'].round(decimals=2)#保留两位小数
Haoshi.drop(['耗时1'],axis=1,inplace=True) #删除多余列
Haoshi.drop(['故障总时长'],axis=1,inplace=True) #删除多余列
Gaojinqingchu=pd.merge(Gaojinqingchu,Haoshi,on=['故障原因类别','故障原因细分'],how='left') 

#0000000000000000000000平均影响用户数
Pingjunhu= HSx1[['故障原因类别','故障原因细分','影响用户数']]
Pingjunhu = Pingjunhu.groupby(['故障原因类别','故障原因细分']).size().reset_index(name='数量')
#Pingjunhu = Pingjunhu.append([{'故障原因类别':'总计','数量':Pingjunhu.apply(lambda x: x.sum()).数量}], ignore_index=True)

Pingjunhu1= HSx1[['故障原因类别','故障原因细分','影响用户数']]
#000000000000000000同一地市累加
Pingjunhu1 = pd.DataFrame(Pingjunhu1.groupby('故障原因细分')['影响用户数'].sum()).reset_index()
#Pingjunhu1 = Pingjunhu1.append([{'故障原因细分':'总计','影响用户数':Pingjunhu1.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
Pingjunhu=pd.merge(Pingjunhu,Pingjunhu1,on=['故障原因细分'],how='left')
Pingjunhu['平均影响用户数'] =Pingjunhu.影响用户数 /Pingjunhu.数量
Pingjunhu['平均影响用户数'] = Pingjunhu['平均影响用户数'].round(decimals=2)#保留两位小数
Gaojinqingchu=pd.merge(Gaojinqingchu,Pingjunhu,on=['故障原因类别','故障原因细分'],how='left')
Gaojinqingchu.drop(['数量'],axis=1,inplace=True) #删除多余列1
Gaojinqingchu.drop(['影响用户数'],axis=1,inplace=True) #删除多余列
Gaojinqingchu=Gaojinqingchu.fillna(0)

#00000000000000南宁占比
Nanning = HSx1[['故障原因类别','故障原因细分','地市']]
Nanning = Nanning[Nanning.地市=='南宁'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='南宁')
Nanning = Nanning.append([{'故障原因类别':'总计','南宁':Nanning.apply(lambda x: x.sum()).南宁}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Nanning,on=['故障原因类别','故障原因细分'],how ='left')
Nanning = Guzhang
Nanning = Nanning.append([{'故障原因类别':'总计','南宁':Nanning.apply(lambda x: x.sum()).南宁}], ignore_index=True)

# 0000000000000000000000000桂林占比
Guilin = HSx1[['故障原因类别','故障原因细分','地市']]
Guilin = Guilin[Guilin.地市=='桂林'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='桂林')
Guilin = Guilin.append([{'故障原因类别':'总计','桂林':Guilin.apply(lambda x: x.sum()).桂林}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Guilin,on=['故障原因类别','故障原因细分'],how ='left')
Guilin = Guzhang
Guilin = Guilin.append([{'故障原因类别':'总计','桂林':Guilin.apply(lambda x: x.sum()).桂林}], ignore_index=True)
test = pd.merge(Nanning,Guilin[['故障原因类别','故障原因细分','桂林']],on=['故障原因类别','故障原因细分'],how ='left')
test.drop(['数量1'],axis=1,inplace=True) #删除多余列

#000000000000000000000000柳州占比
Liuzhou = HSx1[['故障原因类别','故障原因细分','地市']]
Liuzhou = Liuzhou[Liuzhou.地市=='柳州'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='柳州')
Liuzhou = Liuzhou.append([{'故障原因类别':'总计','柳州':Liuzhou.apply(lambda x: x.sum()).柳州}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Liuzhou,on=['故障原因类别','故障原因细分'],how ='left')
Liuzhou = Guzhang
Liuzhou = Liuzhou.append([{'故障原因类别':'总计','柳州':Liuzhou.apply(lambda x: x.sum()).柳州}], ignore_index=True)
test = pd.merge(test,Liuzhou[['故障原因类别','故障原因细分','柳州']],on=['故障原因类别','故障原因细分'],how ='left')

#000000000000000000000玉林占比
Yulin = HSx1[['故障原因类别','故障原因细分','地市']]
Yulin = Yulin[Yulin.地市=='玉林'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='玉林')
Yulin = Yulin.append([{'故障原因类别':'总计','玉林':Yulin.apply(lambda x: x.sum()).玉林}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Yulin,on=['故障原因类别','故障原因细分'],how ='left')
Yulin = Guzhang
Yulin = Yulin.append([{'故障原因类别':'总计','玉林':Yulin.apply(lambda x: x.sum()).玉林}], ignore_index=True)
test = pd.merge(test,Yulin[['故障原因类别','故障原因细分','玉林']],on=['故障原因类别','故障原因细分'],how ='left')

#000000000000000000000百色占比
Basei = HSx1[['故障原因类别','故障原因细分','地市']]
Basei = Basei[Basei.地市=='百色'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='百色')
Basei = Basei.append([{'故障原因类别':'总计','百色':Basei.apply(lambda x: x.sum()).百色}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Basei,on=['故障原因类别','故障原因细分'],how ='left')
Basei = Guzhang
Basei = Basei.append([{'故障原因类别':'总计','百色':Basei.apply(lambda x: x.sum()).百色}], ignore_index=True)
test = pd.merge(test,Basei[['故障原因类别','故障原因细分','百色']],on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000000000河池占比
Hechi = HSx1[['故障原因类别','故障原因细分','地市']]
Hechi = Hechi[Hechi.地市=='河池'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='河池')
Hechi = Hechi.append([{'故障原因类别':'总计','河池':Hechi.apply(lambda x: x.sum()).河池}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Hechi,on=['故障原因类别','故障原因细分'],how ='left')
Hechi = Guzhang
Hechi = Hechi.append([{'故障原因类别':'总计','河池':Hechi.apply(lambda x: x.sum()).河池}], ignore_index=True)
test = pd.merge(test,Hechi[['故障原因类别','故障原因细分','河池']],on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000000000000000贵港占比
Guigang = HSx1[['故障原因类别','故障原因细分','地市']]
Guigang = Guigang[Guigang.地市=='贵港'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='贵港')
Guigang = Guigang.append([{'故障原因类别':'总计','贵港':Guigang.apply(lambda x: x.sum()).贵港}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Guigang,on=['故障原因类别','故障原因细分'],how ='left')
Guigang = Guzhang
Guigang = Guigang.append([{'故障原因类别':'总计','贵港':Guigang.apply(lambda x: x.sum()).贵港}], ignore_index=True)
test = pd.merge(test,Guigang[['故障原因类别','故障原因细分','贵港']],on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000000000000000钦州占比
Qinzhou = HSx1[['故障原因类别','故障原因细分','地市']]
Qinzhou = Qinzhou[Qinzhou.地市=='钦州'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='钦州')
Qinzhou = Qinzhou.append([{'故障原因类别':'总计','钦州':Qinzhou.apply(lambda x: x.sum()).钦州}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Qinzhou,on=['故障原因类别','故障原因细分'],how ='left')
Qinzhou = Guzhang
Qinzhou = Qinzhou.append([{'故障原因类别':'总计','钦州':Qinzhou.apply(lambda x: x.sum()).钦州}], ignore_index=True)
test = pd.merge(test,Qinzhou[['故障原因类别','故障原因细分','钦州']],on=['故障原因类别','故障原因细分'],how ='left')

#000000000000000000000000000梧州占比
Wuzhou = HSx1[['故障原因类别','故障原因细分','地市']]
Wuzhou = Wuzhou[Wuzhou.地市=='梧州'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='梧州')
Wuzhou = Wuzhou.append([{'故障原因类别':'总计','梧州':Wuzhou.apply(lambda x: x.sum()).梧州}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Wuzhou,on=['故障原因类别','故障原因细分'],how ='left')
Wuzhou = Guzhang
Wuzhou = Wuzhou.append([{'故障原因类别':'总计','梧州':Wuzhou.apply(lambda x: x.sum()).梧州}], ignore_index=True)
test = pd.merge(test,Wuzhou[['故障原因类别','故障原因细分','梧州']],on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000000000000北海占比
Beihai = HSx1[['故障原因类别','故障原因细分','地市']]
Beihai = Beihai[Beihai.地市=='北海'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='北海')
Beihai = Beihai.append([{'故障原因类别':'总计','北海':Beihai.apply(lambda x: x.sum()).北海}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Beihai,on=['故障原因类别','故障原因细分'],how ='left')
Beihai = Guzhang
Beihai = Beihai.append([{'故障原因类别':'总计','北海':Beihai.apply(lambda x: x.sum()).北海}], ignore_index=True)
test = pd.merge(test,Beihai[['故障原因类别','故障原因细分','北海']],on=['故障原因类别','故障原因细分'],how ='left')

#0000000000000000000000000崇左占比
Chongzuo = HSx1[['故障原因类别','故障原因细分','地市']]
Chongzuo = Chongzuo[Chongzuo.地市=='崇左'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='崇左')
Chongzuo = Chongzuo.append([{'故障原因类别':'总计','崇左':Chongzuo.apply(lambda x: x.sum()).崇左}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Chongzuo,on=['故障原因类别','故障原因细分'],how ='left')
Chongzuo = Guzhang
Chongzuo = Chongzuo.append([{'故障原因类别':'总计','崇左':Chongzuo.apply(lambda x: x.sum()).崇左}], ignore_index=True)
test = pd.merge(test,Chongzuo[['故障原因类别','故障原因细分','崇左']],on=['故障原因类别','故障原因细分'],how ='left')

#000000000000000000000000来宾占比
Laibai = HSx1[['故障原因类别','故障原因细分','地市']]
Laibai = Laibai[Laibai.地市=='来宾'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='来宾')
Laibai = Laibai.append([{'故障原因类别':'总计','来宾':Laibai.apply(lambda x: x.sum()).来宾}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Laibai,on=['故障原因类别','故障原因细分'],how ='left')
Laibai = Guzhang
Laibai = Laibai.append([{'故障原因类别':'总计','来宾':Laibai.apply(lambda x: x.sum()).来宾}], ignore_index=True)
test = pd.merge(test,Laibai[['故障原因类别','故障原因细分','来宾']],on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000000000000贺州占比
Hezhou = HSx1[['故障原因类别','故障原因细分','地市']]
Hezhou = Hezhou[Hezhou.地市=='贺州'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='贺州')
Hezhou = Hezhou.append([{'故障原因类别':'总计','贺州':Hezhou.apply(lambda x: x.sum()).贺州}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Hezhou,on=['故障原因类别','故障原因细分'],how ='left')
Hezhou = Guzhang
Hezhou = Hezhou.append([{'故障原因类别':'总计','贺州':Hezhou.apply(lambda x: x.sum()).贺州}], ignore_index=True)
test = pd.merge(test,Hezhou[['故障原因类别','故障原因细分','贺州']],on=['故障原因类别','故障原因细分'],how ='left')

#000000000000000000000000防城港占比
Fangchenggang = HSx1[['故障原因类别','故障原因细分','地市']]
Fangchenggang = Fangchenggang[Fangchenggang.地市=='防城港'].groupby(['故障原因类别','故障原因细分']).size().reset_index(name='防城港')
Fangchenggang = Fangchenggang.append([{'故障原因类别':'总计','防城港':Fangchenggang.apply(lambda x: x.sum()).防城港}], ignore_index=True)
Guzhang = pd.merge(Guzhang,Fangchenggang,on=['故障原因类别','故障原因细分'],how ='left')
Fangchenggang = Guzhang
Fangchenggang = Fangchenggang.append([{'故障原因类别':'总计','防城港':Fangchenggang.apply(lambda x: x.sum()).防城港}], ignore_index=True)
test = pd.merge(test,Fangchenggang[['故障原因类别','故障原因细分','防城港']],on=['故障原因类别','故障原因细分'],how ='left')
test = test.fillna(0)

#0000000000000000000000提取指定行 列的某个数值
n=Zhongbiao.iloc[0,1]
g=Zhongbiao.iloc[1,1]
l=Zhongbiao.iloc[2,1]
y=Zhongbiao.iloc[3,1]
b=Zhongbiao.iloc[4,1]
h=Zhongbiao.iloc[5,1]
gg=Zhongbiao.iloc[6,1]
q=Zhongbiao.iloc[7,1]
w=Zhongbiao.iloc[8,1]
bh=Zhongbiao.iloc[9,1]
c=Zhongbiao.iloc[10,1]
lb=Zhongbiao.iloc[11,1]
hz=Zhongbiao.iloc[12,1]
fcg=Zhongbiao.iloc[13,1]

#000000000000000000000新增新的一行总计
test.loc[len(test)] = ' '
test.loc[33,'故障原因类别'] = '总计1'
test.loc[33,'南宁'] = n 
test.loc[33,'桂林'] = g 
test.loc[33,'柳州'] = l 
test.loc[33,'玉林'] = y 
test.loc[33,'百色'] = b 
test.loc[33,'河池'] = h 
test.loc[33,'贵港'] = gg 
test.loc[33,'钦州'] = q
test.loc[33,'梧州'] = w 
test.loc[33,'北海'] = bh 
test.loc[33,'崇左'] = c
test.loc[33,'来宾'] = lb 
test.loc[33,'贺州'] = hz
test.loc[33,'防城港'] = fcg 

test.loc[31,'南宁'] = n - test.iloc[32,2] 
test.loc[31,'桂林'] = g - test.iloc[32,3]
test.loc[31,'柳州'] = l - test.iloc[32,4]
test.loc[31,'玉林'] = y - test.iloc[32,5]
test.loc[31,'百色'] = b - test.iloc[32,6]
test.loc[31,'河池'] = h - test.iloc[32,7]
test.loc[31,'贵港'] = gg - test.iloc[32,8]
test.loc[31,'钦州'] = q - test.iloc[32,9]
test.loc[31,'梧州'] = w - test.iloc[32,10]
test.loc[31,'北海'] = bh - test.iloc[32,11]
test.loc[31,'崇左'] = c - test.iloc[32,12]
test.loc[31,'来宾'] = lb - test.iloc[32,13]
test.loc[31,'贺州'] = hz - test.iloc[32,14]
test.loc[31,'防城港'] = fcg - test.iloc[32,15]
Gaojinqingchu = pd.merge(Gaojinqingchu,test,on=['故障原因类别','故障原因细分'],how ='left')

#00000000000000000计算占比
Gaojinqingchu['南宁占比'] = Gaojinqingchu.南宁 / n
Gaojinqingchu['桂林占比'] = Gaojinqingchu.桂林 / g
Gaojinqingchu['柳州占比'] = Gaojinqingchu.柳州 / l
Gaojinqingchu['玉林占比'] = Gaojinqingchu.玉林 / y
Gaojinqingchu['百色占比'] = Gaojinqingchu.百色 / b
Gaojinqingchu['河池占比'] = Gaojinqingchu.河池 / h
Gaojinqingchu['贵港占比'] = Gaojinqingchu.贵港/ gg
Gaojinqingchu['钦州占比'] = Gaojinqingchu.钦州 / q
Gaojinqingchu['梧州占比'] = Gaojinqingchu.梧州 / w
Gaojinqingchu['北海占比'] = Gaojinqingchu.北海 / bh
Gaojinqingchu['崇左占比'] = Gaojinqingchu.崇左 / c
Gaojinqingchu['来宾占比'] = Gaojinqingchu.来宾 / lb
Gaojinqingchu['贺州占比'] = Gaojinqingchu.贺州 / hz
Gaojinqingchu['防城港占比'] = Gaojinqingchu.防城港 / fcg

Gaojinqingchu['南宁占比']=Gaojinqingchu['南宁占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['桂林占比']=Gaojinqingchu['桂林占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['柳州占比']=Gaojinqingchu['柳州占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['玉林占比']=Gaojinqingchu['玉林占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['百色占比']=Gaojinqingchu['百色占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['河池占比']=Gaojinqingchu['河池占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['贵港占比']=Gaojinqingchu['贵港占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['钦州占比']=Gaojinqingchu['钦州占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['梧州占比']=Gaojinqingchu['梧州占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['北海占比']=Gaojinqingchu['北海占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['崇左占比']=Gaojinqingchu['崇左占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['来宾占比']=Gaojinqingchu['来宾占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['贺州占比']=Gaojinqingchu['贺州占比'].apply(lambda x: '%.2f%%' % (x*100))
Gaojinqingchu['防城港占比']=Gaojinqingchu['防城港占比'].apply(lambda x: '%.2f%%' % (x*100))

Gaojinqingchu=Gaojinqingchu.fillna(0)
Gaojinqingchu = Gaojinqingchu.drop(Gaojinqingchu.index[32]) #删除多余行 总计
#Tingdian = Tingdian.drop_duplicates(keep="last") #drop_duplicates函数删除重复数据，keep="last"保留最后出现的数值  




Gaojinqingchu = Gaojinqingchu[['故障原因类别','故障原因细分','数量1','占比1','耗时','平均影响用户数','南宁','南宁占比','桂林','桂林占比',
                       '柳州','柳州占比','玉林','玉林占比','百色','百色占比','河池','河池占比','贵港','贵港占比','钦州','钦州占比','梧州',
                       '梧州占比','北海','北海占比','崇左','崇左占比','来宾','来宾占比','贺州','贺州占比','防城港','防城港占比']]





########**************重新赋值的百分比 变成空值
Gaojinqingchu.loc[31,'故障原因类别'] = ' '
Gaojinqingchu.loc[31,'故障原因细分'] = ' '
Gaojinqingchu.loc[33,'耗时'] = ' '
Gaojinqingchu.loc[33,'平均影响用户数'] = ' '
Gaojinqingchu.loc[33,'南宁占比'] = ' '
Gaojinqingchu.loc[33,'桂林占比'] = ' '
Gaojinqingchu.loc[33,'柳州占比'] = ' '
Gaojinqingchu.loc[33,'玉林占比'] = ' '
Gaojinqingchu.loc[33,'百色占比'] = ' '
Gaojinqingchu.loc[33,'河池占比'] = ' '
Gaojinqingchu.loc[33,'贵港占比'] = ' '
Gaojinqingchu.loc[33,'钦州占比'] = ' '
Gaojinqingchu.loc[33,'梧州占比'] = ' '
Gaojinqingchu.loc[33,'北海占比'] = ' '
Gaojinqingchu.loc[33,'崇左占比'] = ' '
Gaojinqingchu.loc[33,'来宾占比'] = ' '
Gaojinqingchu.loc[33,'贺州占比'] = ' '
Gaojinqingchu.loc[33,'防城港占比'] = ' '

##########################******          数量2 占比2        *******0
HSx1=HSx1[(HSx1.故障原因类别!='设备被盗/被破坏')]  #删减数据
wa = HSx1.groupby(['故障原因类别']).size().reset_index(name='数量2')
wa = wa.append([{'故障原因类别':'总计','数量2':wa.apply(lambda x: x.sum()).数量2}], ignore_index=True)
waa = wa.iloc[7,1]
Gaojinqingchu = pd.merge(Gaojinqingchu, wa, on=['故障原因类别'], how='left')  # 拼接 
Gaojinqingchu.loc[31,'数量2'] = waa
Gaojinqingchu['占比2'] = Gaojinqingchu.数量2 / z
Gaojinqingchu['占比2']=Gaojinqingchu['占比2'].apply(lambda x: '%.2f%%' % (x*100))

#把数量2 占比2放到指定列 ，调整位置
Cols4 = Gaojinqingchu.columns.tolist()
Cols4.insert(3,Cols4.pop(Cols4.index('数量2')))
Gaojinqingchu = Gaojinqingchu[Cols4]

Cols5 = Gaojinqingchu.columns.tolist()
Cols5.insert(5,Cols5.pop(Cols5.index('占比2')))
Gaojinqingchu = Gaojinqingchu[Cols5]
Gaojinqingchu =Gaojinqingchu.fillna(' ')
Gaojinqingchu.loc[32,'占比2'] = ' '

print('**********************故障类别  计算完毕')

#********************************把指定行列值取出来 100% ***************************
#                                                      *

#      ZA=Gaojinqingchu.iloc[32,7]                     *
#      ZB=Gaojinqingchu.iloc[32,9]                     * 
#      ZC=Gaojinqingchu.iloc[32,11]                    *             
#      ZD=Gaojinqingchu.iloc[32,13]                    *
#      ZE=Gaojinqingchu.iloc[32,15]                    *
#      ZF=Gaojinqingchu.iloc[32,17]                    *
#      ZG=Gaojinqingchu.iloc[32,19]                    *
#      ZH=Gaojinqingchu.iloc[32,21]                    * 
#      ZI=Gaojinqingchu.iloc[32,23]                    *
#      ZJ=Gaojinqingchu.iloc[32,25]                    * 
#      ZK=Gaojinqingchu.iloc[32,27]                    * 
#      ZL=Gaojinqingchu.iloc[32,27]                    *
#      ZM=Gaojinqingchu.iloc[32,31]                    *
#      ZN=Gaojinqingchu.iloc[32,33]                    *
#                                                      *
###############################把取出来的值重新赋值，为空


#######################       中间数据    告警标题     **************************************

Shuju = pd.read_excel('file:///D:/工程量/3-17G故障/告警故障.xlsx')  

print('***************开始计算 告警标题')
#00000000000000 数量1所有的条件值
Shujubiao = HSx1[['中断点','告警标题']]
Shujubiao = Shujubiao.groupby(['中断点','告警标题']).size().reset_index(name='数量1')
Shujubiao = Shujubiao.append([{'告警标题':'总计','数量1':Shujubiao.apply(lambda x: x.sum()).数量1}], ignore_index=True)
Shujubiao= pd.merge(Shuju,Shujubiao,on=['中断点','告警标题'],how ='left')
Shujubiao=Shujubiao.fillna(0)

#0000000000000提取某行某列单个值
Z=Shujubiao.iloc[44,2]
#00000000000000占比1
Shujubiao['占比1'] = Shujubiao.数量1 / Z
Shujubiao['占比1']=Shujubiao['占比1'].apply(lambda x: '%.2f%%' % (x*100))
#Shujubiao['数量2'] = Zaaa + Zaa

###############################耗时
Shuhaoshi = HSx1[['中断点','告警标题']]
Shuhaoshi = Shuhaoshi.groupby(['中断点','告警标题']).size().reset_index(name='耗时1')
Shuhaoshi = Shuhaoshi.append([{'告警标题':'总计','耗时1':Shuhaoshi.apply(lambda x: x.sum()).耗时1}], ignore_index=True)

Shuju=pd.merge(Shuju,Shuhaoshi,on=['中断点','告警标题'],how='left')
Shuhaoshi = Shuju
Shuhaoshi = Shuhaoshi.fillna(0)
#################故障总时长
Shuhaoshi1= HSx1[['中断点','告警标题','故障总时长']]
Shuhaoshi1 = pd.DataFrame(Shuhaoshi1.groupby('告警标题')['故障总时长'].sum()).reset_index()
Shuhaoshi1 = Shuhaoshi1.append([{'告警标题':'总计','故障总时长':Shuhaoshi1.apply(lambda x: x.sum()).故障总时长}], ignore_index=True)
Shuju=pd.merge(Shuju,Shuhaoshi1,on=['告警标题'],how='left')

Shuhaoshi1 = Shuju
Shuhaoshi1['耗时'] = Shuhaoshi1.故障总时长 / Shuhaoshi1.耗时1
Shuhaoshi1 =Shuhaoshi1.fillna(0)
Shujubiao=pd.merge(Shujubiao,Shuhaoshi1,on=['中断点','告警标题'],how='left') 
Shujubiao.drop(['耗时1'],axis=1,inplace=True) #删除多余列
Shujubiao.drop(['故障总时长'],axis=1,inplace=True) #删除多
Shujubiao['耗时'] = Shujubiao['耗时'].round(decimals=2)#保留两位小数

#################平均影响用户数
Shupingjunhu= HSx1[['中断点','告警标题','影响用户数']]

###################同一地市累加
Shupingjunhu = pd.DataFrame(Shupingjunhu.groupby('告警标题')['影响用户数'].sum()).reset_index()
Shupingjunhu = Shupingjunhu.append([{'告警标题':'总计','影响用户数':Shupingjunhu.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
Shuju=pd.merge(Shuju,Shupingjunhu,on=['告警标题'],how='left')
Shupingjunhu = Shuju
Shupingjunhu1= HSx1[['中断点','告警标题','影响用户数']]
Shupingjunhu1 = Shupingjunhu1.groupby(['中断点','告警标题']).size().reset_index(name='数量')
Shupingjunhu1 = Shupingjunhu1.append([{'告警标题':'总计','数量':Shupingjunhu1.apply(lambda x: x.sum()).数量}], ignore_index=True)
Shuju=pd.merge(Shuju,Shupingjunhu1,on=['中断点','告警标题'],how='left')
Shupingjunhu1 =Shuju
Shupingjunhu1['平均影响用户数'] =Shupingjunhu1.影响用户数 /Shupingjunhu1.数量
Shupingjunhu1.drop(['故障总时长'],axis=1,inplace=True) #删除多余列
Shupingjunhu1.drop(['耗时1'],axis=1,inplace=True) #删除多余列
Shupingjunhu1.drop(['耗时'],axis=1,inplace=True) #删除多余列
Shuju.drop(['影响用户数'],axis=1,inplace=True) #删除多余列
Shuju.drop(['数量'],axis=1,inplace=True) #删除多余列

Shupingjunhu1 = Shupingjunhu1.fillna(0)
Shujubiao=pd.merge(Shujubiao,Shupingjunhu1,on=['中断点','告警标题'],how='left')
Shujubiao['平均影响用户数'] = Shujubiao['平均影响用户数'].round(decimals=2)#保留两位小数

#=Shujubiao['耗时'].sum() 求一列的总和

#00000000000000000000000南宁占比
Nanning1= HSx1[['中断点','告警标题','地市']]
Nanning1 = Nanning1[Nanning1.地市=='南宁'].groupby(['中断点','告警标题']).size().reset_index(name='南宁')
Nanning1 = Nanning1.append([{'告警标题':'总计','南宁':Nanning1.apply(lambda x: x.sum()).南宁}], ignore_index=True)
Shuju=pd.merge(Shuju,Nanning1,on=['中断点','告警标题'],how='left')
Nanning1 = Shuju

#0000000000000000000桂林占比
Guilin1 = HSx1[['中断点','告警标题','地市']]
Guilin1 = Guilin1[Guilin1.地市=='桂林'].groupby(['中断点','告警标题']).size().reset_index(name='桂林')
Guilin1 = Guilin1.append([{'告警标题':'总计','桂林':Guilin1.apply(lambda x: x.sum()).桂林}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Guilin1,on=['中断点','告警标题'],how ='left')
Guilin=Guilin.fillna(0)

#000000000000000000000柳州占比
Liuzhou1= HSx1[['中断点','告警标题','地市']]
Liuzhou1 = Liuzhou1[Liuzhou1.地市=='柳州'].groupby(['中断点','告警标题']).size().reset_index(name='柳州')
Liuzhou1 = Liuzhou1.append([{'告警标题':'总计','柳州':Liuzhou1.apply(lambda x: x.sum()).柳州}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Liuzhou1,on=['中断点','告警标题'],how ='left')

#000000000000000000000000玉林占比
Yulin1 = HSx1[['中断点','告警标题','地市']]
Yulin1 = Yulin1[Yulin1.地市=='玉林'].groupby(['中断点','告警标题']).size().reset_index(name='玉林')
Yulin1 = Yulin1.append([{'告警标题':'总计','玉林':Yulin1.apply(lambda x: x.sum()).玉林}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Yulin1,on=['中断点','告警标题'],how ='left')

#0000000000000000000000000百色占比
Baise1 = HSx1[['中断点','告警标题','地市']]
Baise1 = Baise1[Baise1.地市=='百色'].groupby(['中断点','告警标题']).size().reset_index(name='百色')
Baise1 = Baise1.append([{'告警标题':'总计','百色':Baise1.apply(lambda x: x.sum()).百色}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Baise1,on=['中断点','告警标题'],how ='left')

#00000000000000000000000000河池占比
Hechi1 = HSx1[['中断点','告警标题','地市']]
Hechi1 = Hechi1[Hechi1.地市=='河池'].groupby(['中断点','告警标题']).size().reset_index(name='河池')
Hechi1 = Hechi1.append([{'告警标题':'总计','河池':Hechi1.apply(lambda x: x.sum()).河池}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Hechi1,on=['中断点','告警标题'],how ='left')

#00000000000000000000000000000贵港占比
Guigang1 = HSx1[['中断点','告警标题','地市']]
Guigang1 = Guigang1[Guigang1.地市=='贵港'].groupby(['中断点','告警标题']).size().reset_index(name='贵港')
Guigang1 = Guigang1.append([{'告警标题':'总计','贵港':Guigang1.apply(lambda x: x.sum()).贵港}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Guigang1,on=['中断点','告警标题'],how ='left')

#00000000000000000000000000钦州占比
Qinzhou1 = HSx1[['中断点','告警标题','地市']]
Qinzhou1 = Qinzhou1[Qinzhou1.地市=='钦州'].groupby(['中断点','告警标题']).size().reset_index(name='钦州')
Qinzhou1 = Qinzhou1.append([{'告警标题':'总计','钦州':Qinzhou1.apply(lambda x: x.sum()).钦州}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Qinzhou1,on=['中断点','告警标题'],how ='left')

#000000000000000000000000000梧州占比
Wuzhou1 = HSx1[['中断点','告警标题','地市']]
Wuzhou1 = Wuzhou1[Wuzhou1.地市=='梧州'].groupby(['中断点','告警标题']).size().reset_index(name='梧州')
Wuzhou1 = Wuzhou1.append([{'告警标题':'总计','梧州':Wuzhou1.apply(lambda x: x.sum()).梧州}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Wuzhou1,on=['中断点','告警标题'],how ='left')

#00000000000000000000000000北海占比
Beihai1 = HSx1[['中断点','告警标题','地市']]
Beihai1 = Beihai1[Beihai1.地市=='北海'].groupby(['中断点','告警标题']).size().reset_index(name='北海')
Beihai1 = Beihai1.append([{'告警标题':'总计','北海':Beihai1.apply(lambda x: x.sum()).北海}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Beihai1,on=['中断点','告警标题'],how ='left')

#000000000000000000000000000崇左占比
Chongzuo1 = HSx1[['中断点','告警标题','地市']]
Chongzuo1 = Chongzuo1[Chongzuo1.地市=='崇左'].groupby(['中断点','告警标题']).size().reset_index(name='崇左')
Chongzuo1 = Chongzuo1.append([{'告警标题':'总计','崇左':Chongzuo1.apply(lambda x: x.sum()).崇左}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Chongzuo1,on=['中断点','告警标题'],how ='left')

#000000000000000000000000000000来宾占比
Laibin1 = HSx1[['中断点','告警标题','地市']]
Laibin1 = Laibin1[Laibin1.地市=='来宾'].groupby(['中断点','告警标题']).size().reset_index(name='来宾')
Laibin1 = Laibin1.append([{'告警标题':'总计','来宾':Laibin1.apply(lambda x: x.sum()).来宾}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Laibin1,on=['中断点','告警标题'],how ='left')

#00000000000000000000000000贺州占比
Hezhou1 = HSx1[['中断点','告警标题','地市']]
Hezhou1 = Hezhou1[Hezhou1.地市=='贺州'].groupby(['中断点','告警标题']).size().reset_index(name='贺州')
Hezhou1 = Hezhou1.append([{'告警标题':'总计','贺州':Hezhou1.apply(lambda x: x.sum()).贺州}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Hezhou1,on=['中断点','告警标题'],how ='left')

#000000000000000000000000000000防城港占比
Fangchenggang1 = HSx1[['中断点','告警标题','地市']]
Fangchenggang1 = Fangchenggang1[Fangchenggang1.地市=='防城港'].groupby(['中断点','告警标题']).size().reset_index(name='防城港')
Fangchenggang1 = Fangchenggang1.append([{'告警标题':'总计','防城港':Fangchenggang1.apply(lambda x: x.sum()).防城港}], ignore_index=True)
Nanning1 = pd.merge(Nanning1,Fangchenggang1,on=['中断点','告警标题'],how ='left')
Nanning1.drop(['平均影响用户数'],axis=1,inplace=True) #删除多余列

#0000000000000提取指定行列的值
n1=Nanning1.iloc[44,2]
g1=Nanning1.iloc[44,3]
l1=Nanning1.iloc[44,4]
y1=Nanning1.iloc[44,5]
b1=Nanning1.iloc[44,6]
h1=Nanning1.iloc[44,7]
gg1=Nanning1.iloc[44,8]
q1=Nanning1.iloc[44,9]
w1=Nanning1.iloc[44,10]
bh1=Nanning1.iloc[44,11]
c1=Nanning1.iloc[44,12]
lb1=Nanning1.iloc[44,13]
hz1=Nanning1.iloc[44,14]
fcg1=Nanning1.iloc[44,15]

#0000000000000000000000000000000000计算占比
Nanning1=Nanning1.fillna(0)
Nanning1['南宁占比'] = Nanning1.南宁 / n1
Nanning1['桂林占比'] = Nanning1.桂林 / g1
Nanning1['柳州占比'] = Nanning1.柳州 / l1
Nanning1['玉林占比'] = Nanning1.玉林 / y1
Nanning1['百色占比'] = Nanning1.百色 / b1
Nanning1['河池占比'] = Nanning1.河池 / h1
Nanning1['贵港占比'] = Nanning1.贵港/ gg1
Nanning1['钦州占比'] = Nanning1.钦州 / q1
Nanning1['梧州占比'] = Nanning1.梧州 / w1
Nanning1['北海占比'] = Nanning1.北海 / bh1
Nanning1['崇左占比'] = Nanning1.崇左 / c1
Nanning1['来宾占比'] = Nanning1.来宾 / lb1
Nanning1['贺州占比'] = Nanning1.贺州 / hz1
Nanning1['防城港占比'] = Nanning1.防城港 / fcg1

Nanning1['南宁占比']=Nanning1['南宁占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['桂林占比']=Nanning1['桂林占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['柳州占比']=Nanning1['柳州占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['玉林占比']=Nanning1['玉林占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['百色占比']=Nanning1['百色占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['河池占比']=Nanning1['河池占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['贵港占比']=Nanning1['贵港占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['钦州占比']=Nanning1['钦州占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['梧州占比']=Nanning1['梧州占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['北海占比']=Nanning1['北海占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['崇左占比']=Nanning1['崇左占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['来宾占比']=Nanning1['来宾占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['贺州占比']=Nanning1['贺州占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1['防城港占比']=Nanning1['防城港占比'].apply(lambda x: '%.2f%%' % (x*100))
Nanning1=Nanning1.fillna(0)
Shujubiao = pd.merge(Shujubiao,Nanning1,on=['中断点','告警标题'],how ='left')
Shujubiao.loc[44,'中断点'] = ' '

Zhanbi=pd.read_excel('file:///D:/工程量/3-17G故障/G故障工单统计-20210317.xlsx',sheet_name ='中断点数量2')  #读取 前台退单原因分类
Zhanbi = Zhanbi.fillna(0)
Shujubiao = pd.merge(Shujubiao,Zhanbi,on=['中断点','告警标题'],how ='left')
Shujubiao = Shujubiao.fillna(0)
Shujubiao['占比2']=Shujubiao['占比2'].apply(lambda x: '%.2f%%' % (x*100))


Cols = Shujubiao.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
Cols.insert(3, Cols.pop(Cols.index('数量2')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
Shujubiao = Shujubiao[Cols]  

Cols1 = Shujubiao.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
Cols1.insert(5, Cols1.pop(Cols1.index('占比2')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
Shujubiao = Shujubiao[Cols1]  


Shujubiao = Shujubiao[['中断点','告警标题','数量1','数量2','占比1','占比2','耗时','平均影响用户数','南宁','南宁占比','桂林','桂林占比',
                       '柳州','柳州占比','玉林','玉林占比','百色','百色占比','河池','河池占比','贵港','贵港占比','钦州','钦州占比','梧州',
                       '梧州占比','北海','北海占比','崇左','崇左占比','来宾','来宾占比','贺州','贺州占比','防城港','防城港占比']]

print('***********************告警标题 计算完毕')

with pd.ExcelWriter('3-17故障完成表' + '.xlsx') as writer:  # 写入结果为当前路径
         Zhongbiao.to_excel(writer, sheet_name='故障表', startcol=0, index=False, header=True)
         Gaojinqingchu.to_excel(writer, sheet_name='故障原因类别', startcol=0, index=False, header=True)
         Shujubiao.to_excel(writer, sheet_name='告警标题', startcol=0, index=False, header=True)

elapsed0 = (time.clock() - start0)  # 结束计时
print("读取数据用时:", elapsed0, '秒')











