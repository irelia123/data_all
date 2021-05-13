# -*- coding: utf-8 -*-
"""
Created on Thu Apr 29 17:06:37 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np

Dishi = pd.DataFrame(
        {'区县': ['江州区','扶绥县','宁明县','龙州县','大新县','凭祥市','天等县','合计' ]},       
           pd.Index(range(8)))

Dishi1 = pd.DataFrame(
        {'区县': ['江州区','扶绥县','宁明县','龙州县','大新县','凭祥市','天等县','合计' ]},       
           pd.Index(range(8)))

Dishi2 = pd.DataFrame(
        {'区县': ['江州区','扶绥县','宁明县','龙州县','大新县','凭祥市','天等县','合计' ]},       
           pd.Index(range(8)))

Dishi3 = pd.DataFrame(
        {'区县': ['江州区','扶绥县','宁明县','龙州县','大新县','凭祥市','天等县','合计' ]},       
           pd.Index(range(8)))

Hsx = pd.read_excel('./截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx',sheet_name='已修复')
Jiyaguzhang = pd.read_excel('./截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx',sheet_name='批量故障跟进表')

Hsx = Hsx.fillna('')

print('*******开始运行')
#todaytime1 = datetime.datetime.strptime('2021-5-12','%Y-%m-%d') #手动输入日期
#todaytime1 = str(todaytime1.year)+'/'+str(todaytime1.month) +'/'+ todaytime1.strftime('%d')


todaytime=datetime.datetime.now()  #系统取当天日期
first_day = datetime.datetime(todaytime.year, todaytime.month, 12, 23, 59, 59)  # datetime类型 2019-09-01 00:00:00
todaytime1 = str(todaytime.month) + '月' + todaytime.strftime('%d') + '日'  # 日期转换str：4月17日

#--------------     故障次数   *******************

# applymap函数：应用于整个文件类型转换 
#Hsx = Hsx.applymap(int)
#转换某列的数据类型
#Hsx['影响用户数'].astype(str)
#Hsx['影响用户数'] = [int(i) for i in range(len(Hsx['影响用户数']))]
#Hsx = Hsx[(Hsx.影响用户数!='')]  #删减空白数据
Guzhangcishu_cs = Hsx[(Hsx.责任部门 == '客响') & (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='城镇')
Guzhangcishu_cs = Guzhangcishu_cs.append([{'区县':'合计','城镇':Guzhangcishu_cs.apply(lambda x: x.sum()).城镇}], ignore_index=True)

Guzhangcishu_nc = Hsx[(Hsx.责任部门 == '客响') & (Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='农村')
Guzhangcishu_nc = pd.DataFrame(Guzhangcishu_nc.groupby('区县')['农村'].sum()).reset_index()

Guzhangcishu_nc = Guzhangcishu_nc.append([{'区县':'合计','农村':Guzhangcishu_nc.apply(lambda x: x.sum()).农村}], ignore_index=True)
Guzhangcishu_cs = pd.merge(Guzhangcishu_cs,Guzhangcishu_nc,on = ['区县'], how = 'left')
Guzhang = Guzhangcishu_cs

Guzhang = Guzhang.fillna(0)
Guzhang['线路故障总数'] = Guzhang.城镇 + Guzhang.农村

#******************     处理时长 ******************

Shichang = Hsx[ (Hsx.责任部门 == '客响') & (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='城镇时长')
Shichang = pd.DataFrame(Shichang.groupby('区县')['总时长'].sum()).reset_index()
Shichang = Shichang.append([{'区县':'合计','总时长':Shichang.apply(lambda x: x.sum()).总时长}], ignore_index=True)
#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值
Shichang['城镇时长'] = Shichang.总时长 / Guzhang.城镇
Shichang = Shichang.round({'城镇时长': 2})  # 四舍五入
Shichang.rename(columns={'总时长': '时长'}, inplace=True)  # 改列名


Shichang_nc = Hsx[ (Hsx.责任部门 == '客响') & (Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='农村时长')
Shichang_nc = pd.DataFrame(Shichang_nc.groupby('区县')['总时长'].sum()).reset_index()
Shichang_nc = Shichang_nc.append([{'区县':'合计','总时长':Shichang_nc.apply(lambda x: x.sum()).总时长}], ignore_index=True)
Shichang = pd.merge(Shichang,Shichang_nc,on = ['区县'], how = 'left')
Shichang = Shichang.fillna(0)

#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值
Shichang['农村时长'] = Shichang.总时长 / Guzhang.农村
Shichang = Shichang.round({'农村时长': 2})  # 四舍五入
Shichang.drop(['总时长'],axis=1,inplace=True) #删除多余列
Shichang.drop(['时长'],axis=1,inplace=True) #删除多余列
Guzhang = pd.merge(Guzhang,Shichang,on = ['区县'], how = 'left')
Guzhang = Guzhang.fillna(0)


#线路故障平均时长
Shichang_gz = Hsx[ (Hsx.责任部门 == '客响') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='故障时长')
quxian = Shichang_gz.groupby(['区县']).size().reset_index(name='数量')
quxian = quxian.append([{'区县':'合计','数量':quxian.apply(lambda x: x.sum()).数量}], ignore_index=True)

Shichang_gz = pd.DataFrame(Shichang_gz.groupby('区县')['总时长'].sum()).reset_index()
Shichang_gz = Shichang_gz.append([{'区县':'合计','总时长':Shichang_gz.apply(lambda x: x.sum()).总时长}], ignore_index=True)
Guzhang['线路故障平均时长'] = Shichang_gz.总时长 / quxian.数量

Guzhang = Guzhang.round({'线路故障平均时长': 2})  # 四舍五入
#Shichang_gz1 = Hsx[['区县', '总时长']].groupby(['区县']).mean().reset_index() #计算平均时长

#******   及时率 *******
Jishilv = Hsx[ (Hsx.责任部门 == '客响') & (Hsx.所属区域 =='农村') & (Hsx.总时长 <=36)].groupby(['区县']).size().reset_index(name='数量')
Jishilv = pd.DataFrame(Jishilv.groupby('区县')['数量'].sum()).reset_index()
Jishilv = Jishilv.append([{'区县':'合计','数量':Jishilv.apply(lambda x: x.sum()).数量}], ignore_index=True)


Jishilv1 = Hsx[ (Hsx.责任部门 == '客响') & (Hsx.所属区域 =='城镇') & (Hsx.总时长 <=18)].groupby(['区县']).size().reset_index(name='数量')
Jishilv1 = pd.DataFrame(Jishilv1.groupby('区县')['数量'].sum()).reset_index()
Jishilv1 = Jishilv1.append([{'区县':'合计','数量':Jishilv1.apply(lambda x: x.sum()).数量}], ignore_index=True)

Jishilv2 = Hsx[ (Hsx.责任部门 == '客响')].groupby(['区县']).size().reset_index(name='数量')
Jishilv2 = pd.DataFrame(Jishilv2.groupby('区县')['数量'].sum()).reset_index()
Jishilv2 = Jishilv2.append([{'区县':'合计','数量':Jishilv2.apply(lambda x: x.sum()).数量}], ignore_index=True)


Jishilv['数量1'] = Jishilv.数量 + Jishilv1.数量
Guzhang['整体处理及时率'] = Jishilv.数量1 / Jishilv2.数量
#Guzhang = Guzhang.round({'整体处理及时率': 2})  # 四舍五入
Guzhang['整体处理及时率']=Guzhang['整体处理及时率'].apply(lambda x: '%.2f%%' % (x*100))

#-***************影响用户数
Yonghushu = Hsx[(Hsx.责任部门 == '客响') & (Hsx.影响用户数 >=10) ]
Yonghushu = pd.DataFrame(Yonghushu.groupby('区县')['影响用户数'].sum()).reset_index()
Yonghushu = Yonghushu.append([{'区县':'合计','影响用户数':Yonghushu.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)



#Xinzeng = pd.read_excel('file:///D:/工程量/群障/截至2021年4月28日有线专业线路故障修复完成情况（影响用户数超10户）(新版)0429.xlsx',sheet_name = '今日新增')
#Huifui= pd.read_excel('file:///D:/工程量/群障/截至2021年4月28日有线专业线路故障修复完成情况（影响用户数超10户）(新版)0429.xlsx',sheet_name = '今日恢复')
#Jiyaguzhang= pd.read_excel('file:///D:/工程量/群障/截至2021年4月28日有线专业线路故障修复完成情况（影响用户数超10户）(新版)0429.xlsx',sheet_name = '积压故障')
#Jiyaguzhang.恢复时间 = todaytime
#Jiyaguzhang = pd.read_excel('file:///D:/工程量/群障/批量故障跟进表 (1).xlsx')
#Jiyaguzhang.恢复时间 = todaytime
Jiyaguzhang.rename(columns={'通知/移交时间': '今日新增时间'}, inplace=True)  # 改列名
Jiyaguzhang.rename(columns={'县份': '区县'}, inplace=True)  # 改列名
#Jiyaguzhang['现在时'] = (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.days*24 + (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Jiyaguzhang['现在时'] = (todaytime-pd.to_datetime(Jiyaguzhang.今日新增时间)).dt.days*24 + (todaytime-pd.to_datetime(Jiyaguzhang.今日新增时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改

#**********今日新增
Jiyaguzhang['今日新增时间'] = Jiyaguzhang['今日新增时间'].dt.strftime('%Y/%m/%d')
Jinrizeng = Jiyaguzhang[Jiyaguzhang.今日新增时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除
Jinrizeng = Jinrizeng[ (Jinrizeng.责任部门 == '客响') & (Jinrizeng.影响用户 >=10) ].groupby(['区县','影响用户']).size().reset_index(name='今日新增')
Jinrizeng = pd.DataFrame(Jinrizeng.groupby('区县')['今日新增'].sum()).reset_index()
Jinrizeng = Jinrizeng.append([{'区县':'合计','今日新增':Jinrizeng.apply(lambda x: x.sum()).今日新增}], ignore_index=True)

#**********今日恢复
Jinrifu=Jiyaguzhang[(Jiyaguzhang.恢复时间!='未恢复')]  #删减多余数据
Jinrifu['恢复时间'] = Jinrifu['恢复时间'].dt.strftime('%Y/%m/%d')
Jinrifu = Jinrifu[Jinrifu.恢复时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除
Jinrifu = Jinrifu[(Jiyaguzhang.责任部门 == '客响') & (Jinrifu.影响用户 >=10)].groupby(['区县']).size().reset_index(name='今日恢复')
Jinrifu = pd.DataFrame(Jinrifu.groupby('区县')['今日恢复'].sum()).reset_index()
Jinrifu = Jinrifu.append([{'区县':'合计','今日恢复':Jinrifu.apply(lambda x: x.sum()).今日恢复}], ignore_index=True)
#**********积压故障
Jiya = Jiyaguzhang[Jiyaguzhang.恢复时间.str.contains('未恢复') ==True]  #筛选4/28的数据，，，，，false为剔除

Jiya = Jiya[ (Jiya.责任部门 == '客响') & (Jiya.影响用户 >=10)].groupby(['区县','影响用户']).size().reset_index(name='积压故障')
Jiya = pd.DataFrame(Jiya.groupby('区县')['积压故障'].sum()).reset_index()
Jiya = Jiya.append([{'区县':'合计','积压故障':Jiya.apply(lambda x: x.sum()).积压故障}], ignore_index=True)


#********超48小时工单数
#Jiyaguzhang['现在时'] = (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.days*24 + (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Chao48 = Jiyaguzhang[ (Jiyaguzhang.责任部门 == '客响') & (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48)].groupby(['区县']).size().reset_index(name='超48小时工单数')
Chao48 = pd.DataFrame(Chao48.groupby('区县')['超48小时工单数'].sum()).reset_index()
Chao48 = Chao48.append([{'区县':'合计','超48小时工单数':Chao48.apply(lambda x: x.sum()).超48小时工单数}], ignore_index=True)

#********超48小时用户数

Chao48y = Jiyaguzhang[ (Jiyaguzhang.责任部门 == '客响') & (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48)]
Chao48y = pd.DataFrame(Chao48y.groupby('区县')['影响用户'].sum()).reset_index()
Chao48y = Chao48y.append([{'区县':'合计','影响用户':Chao48y.apply(lambda x: x.sum()).影响用户}], ignore_index=True)
Chao48y.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名


Dishi = pd.merge(Dishi,Guzhang,on = ['区县'],how = 'left')
Guzhang = Dishi
Guzhang = pd.merge(Guzhang,Yonghushu,on = ['区县'], how = 'left')
Guzhang = pd.merge(Guzhang,Jinrizeng,on = ['区县'], how = 'left')
Guzhang = pd.merge(Guzhang,Jinrifu,on = ['区县'], how = 'left')
Guzhang = pd.merge(Guzhang,Jiya,on = ['区县'], how = 'left')
Guzhang = pd.merge(Guzhang,Chao48,on = ['区县'], how = 'left')
Guzhang = pd.merge(Guzhang,Chao48y,on = ['区县'], how = 'left')
Guzhang = Guzhang.fillna(0)

############******************        影响pon口10户以上            *******************



Guzhangcishu_cs1 = Hsx[(Hsx.责任部门 == '传输') & (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='城镇')
Guzhangcishu_cs1 = Guzhangcishu_cs1.append([{'区县':'合计','城镇':Guzhangcishu_cs1.apply(lambda x: x.sum()).城镇}], ignore_index=True)
Dishi1 = pd.merge(Dishi1,Guzhangcishu_cs1,on = ['区县'], how = 'left')
Guzhangcishu_cs1 = Dishi1

Guzhangcishu_nc1 = Hsx[(Hsx.责任部门 == '传输') & (Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='农村')
Guzhangcishu_nc1 = Guzhangcishu_nc1.append([{'区县':'合计','农村':Guzhangcishu_nc1.apply(lambda x: x.sum()).农村}], ignore_index=True)
Guzhang1 = Dishi1
Guzhang1 = pd.merge(Guzhangcishu_cs1,Guzhangcishu_nc1,on = ['区县'], how = 'left')
Guzhang1 = Guzhang1.fillna(0)
Guzhang1['线路故障总数'] = Guzhang1.城镇 + Guzhang1.农村
Guzhang1 = Guzhang1.fillna(0)

#******************     处理时长 ******************

Shichang1 = Hsx[ (Hsx.责任部门 == '传输') & (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='城镇时长')
Shichang1 = pd.DataFrame(Shichang1.groupby('区县')['总时长'].sum()).reset_index()
Shichang1 = Shichang1.append([{'区县':'合计','总时长':Shichang1.apply(lambda x: x.sum()).总时长}], ignore_index=True)
#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值
Zheyou = pd.merge(Dishi1,Shichang1,on = ['区县'], how = 'left')
Zheyou['城镇时长'] = Zheyou.总时长 / Zheyou.城镇
Zheyou = Zheyou.round({'城镇时长': 2})  # 四舍五入
Zheyou.drop(['城镇'],axis=1,inplace=True) #删除多余列
Zheyou.drop(['总时长'],axis=1,inplace=True) #删除多余列
Zheyou = Zheyou.fillna(0)


####*****农村平均
Shichang_nc1 = Hsx[ (Hsx.责任部门 == '传输') & (Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='农村时长')
Shichang_nc1 = pd.DataFrame(Shichang_nc1.groupby('区县')['总时长'].sum()).reset_index()
Shichang_nc1 = Shichang_nc1.append([{'区县':'合计','总时长':Shichang_nc1.apply(lambda x: x.sum()).总时长}], ignore_index=True)
#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值
Zheyou = pd.merge(Zheyou,Shichang_nc1,on = ['区县'], how = 'left')
Zheyou['农村时长'] = Zheyou.总时长 / Guzhang1.农村
Zheyou = Zheyou.round({'农村时长': 2})  # 四舍五入
Zheyou.drop(['总时长'],axis=1,inplace=True) #删除多余列


#线路故障平均时长
Shichang_gz1 = Hsx[ (Hsx.责任部门 == '传输') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='故障时长')
quxian1 = Shichang_gz1.groupby(['区县']).size().reset_index(name='数量')
quxian1 = quxian1.append([{'区县':'合计','数量':quxian1.apply(lambda x: x.sum()).数量}], ignore_index=True)

Shichang_gz1 = pd.DataFrame(Shichang_gz1.groupby('区县')['总时长'].sum()).reset_index()
Shichang_gz1 = Shichang_gz1.append([{'区县':'合计','总时长':Shichang_gz1.apply(lambda x: x.sum()).总时长}], ignore_index=True)
Dishi1 = pd.merge(Dishi1,Shichang_gz1,on = ['区县'], how = 'left')
Shichang_gz1 = Dishi1
Zheyou = pd.merge(Zheyou,quxian1,on = ['区县'], how = 'left')

Zheyou['线路故障平均时长'] = Shichang_gz1.总时长 / Zheyou.数量
Zheyou = Zheyou.round({'线路故障平均时长': 2})  # 四舍五入
#Zheyou.drop(['总时长'],axis=1,inplace=True) #删除多余列
Zheyou.drop(['数量'],axis=1,inplace=True) #删除多余列

############## 及时率****************

Jishilv4 = Hsx[ (Hsx.责任部门 == '传输') & (Hsx.所属区域 =='农村') & (Hsx.总时长 <=36)].groupby(['区县']).size().reset_index(name='数量')
Jishilv4 = pd.DataFrame(Jishilv4.groupby('区县')['数量'].sum()).reset_index()
Jishilv4 = Jishilv4.append([{'区县':'合计','数量':Jishilv4.apply(lambda x: x.sum()).数量}], ignore_index=True)
Dishi1 = pd.merge(Dishi1,Jishilv4,on = ['区县'], how = 'left')

Jishilv41 = Hsx[ (Hsx.责任部门 == '传输') & (Hsx.所属区域 =='城镇') & (Hsx.总时长 <=18)].groupby(['区县']).size().reset_index(name='数量')
Jishilv41 = pd.DataFrame(Jishilv41.groupby('区县')['数量'].sum()).reset_index()
Jishilv41 = Jishilv41.append([{'区县':'合计','数量':Jishilv41.apply(lambda x: x.sum()).数量}], ignore_index=True)
Dishi1 = pd.merge(Dishi1,Jishilv41,on = ['区县'], how = 'left')
Dishi1 = Dishi1.fillna(0)
Dishi1['数量1'] = Dishi1.数量_x + Dishi1.数量_y
Dishi1.drop(['数量_x'],axis=1,inplace=True) #删除多余列
Dishi1.drop(['数量_y'],axis=1,inplace=True) #删除多余列


Jishilv5 = Hsx[ (Hsx.责任部门 == '传输')].groupby(['区县']).size().reset_index(name='数量')
Jishilv5 = pd.DataFrame(Jishilv5.groupby('区县')['数量'].sum()).reset_index()
Jishilv5 = Jishilv5.append([{'区县':'合计','数量':Jishilv5.apply(lambda x: x.sum()).数量}], ignore_index=True)
Dishi1 = pd.merge(Dishi1,Jishilv5,on = ['区县'], how = 'left')


Zheyou['整体处理及时率'] = Dishi1.数量1 / Dishi1.数量
#Zheyou = Zheyou.round({'整体处理及时率': 2})  # 四舍五入
Zheyou['整体处理及时率']=Zheyou['整体处理及时率'].apply(lambda x: '%.2f%%' % (x*100))


#-***************影响用户数
Yonghushu1 = Hsx[(Hsx.责任部门 == '传输') & (Hsx.影响用户数 >=10)]
Yonghushu1 = pd.DataFrame(Yonghushu1.groupby('区县')['影响用户数'].sum()).reset_index()
Yonghushu1 = Yonghushu1.append([{'区县':'合计','影响用户数':Yonghushu1.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
Zheyou = pd.merge(Zheyou,Yonghushu1,on = ['区县'], how = 'left')

#**********今日新增
Jinrizeng1 = Jiyaguzhang[Jiyaguzhang.今日新增时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除
Jinrizeng1 = Jinrizeng1[ (Jinrizeng1.责任部门 == '传输') & (Jinrizeng1.影响用户 >=10) ].groupby(['区县','影响用户']).size().reset_index(name='今日新增')
Jinrizeng1 = pd.DataFrame(Jinrizeng1.groupby('区县')['今日新增'].sum()).reset_index()
Jinrizeng1 = Jinrizeng1.append([{'区县':'合计','今日新增':Jinrizeng1.apply(lambda x: x.sum()).今日新增}], ignore_index=True)
Zheyou = pd.merge(Zheyou,Jinrizeng1,on = ['区县'], how = 'left')

#**********今日恢复
Jinrifu1=Jiyaguzhang[(Jiyaguzhang.恢复时间!='未恢复')]  #删减多余数据
Jinrifu1['恢复时间'] = Jinrifu1['恢复时间'].dt.strftime('%Y/%m/%d')
Jinrifu1 = Jinrifu1[Jinrifu1.恢复时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除

Jinrifu1 = Jinrifu1[ (Jinrifu1.责任部门 == '传输') & (Jinrifu1.影响用户 >=10)].groupby(['区县']).size().reset_index(name='今日恢复')
Jinrifu1 = pd.DataFrame(Jinrifu1.groupby('区县')['今日恢复'].sum()).reset_index()
Jinrifu1 = Jinrifu1.append([{'区县':'合计','今日恢复':Jinrifu1.apply(lambda x: x.sum()).今日恢复}], ignore_index=True)
Zheyou = pd.merge(Zheyou,Jinrifu1,on = ['区县'], how = 'left')


#**********积压故障
#Jiyaguzhang = Jiyaguzhang[(Jiyaguzhang.影响用户!='不可删除')]  #删减空白数据

Jiya1 = Jiyaguzhang[Jiyaguzhang.恢复时间.str.contains('未恢复') ==True]  #筛选4/28的数据，，，，，false为剔除
Jiya1 = Jiya1[ (Jiya1.责任部门 == '传输') & (Jiya1.影响用户 >=10)].groupby(['区县','影响用户']).size().reset_index(name='积压故障')
Jiya1 = pd.DataFrame(Jiya1.groupby('区县')['积压故障'].sum()).reset_index()
Jiya1 = Jiya1.append([{'区县':'合计','积压故障':Jiya1.apply(lambda x: x.sum()).积压故障}], ignore_index=True)
Zheyou = pd.merge(Zheyou,Jiya1,on = ['区县'], how = 'left')

#********超48小时工单数
#Jiyaguzhang['现在时'] = (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.days*24 + (pd.to_datetime(Jiyaguzhang.恢复时间)-pd.to_datetime(Jiyaguzhang.通知时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Chao481 = Jiyaguzhang[ (Jiyaguzhang.责任部门 == '传输') & (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48) & (Jiyaguzhang.恢复时间 =='未恢复')].groupby(['区县']).size().reset_index(name='超48小时工单数')
Chao481 = pd.DataFrame(Chao481.groupby('区县')['超48小时工单数'].sum()).reset_index()
Chao481 = Chao481.append([{'区县':'合计','超48小时工单数':Chao481.apply(lambda x: x.sum()).超48小时工单数}], ignore_index=True)
Zheyou = pd.merge(Zheyou,Chao481,on = ['区县'], how = 'left')

#********超48小时用户数
Chao48y1 = Jiyaguzhang[ (Jiyaguzhang.责任部门 == '传输') & (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48) & (Jiyaguzhang.恢复时间 =='未恢复')]#.groupby(['区县','影响用户']).size().reset_index(name='用户数')
Chao48y1 = pd.DataFrame(Chao48y1.groupby('区县')['影响用户'].sum()).reset_index()
Chao48y1= Chao48y1.append([{'区县':'合计','影响用户':Chao48y1.apply(lambda x: x.sum()).影响用户}], ignore_index=True)
Chao48y1.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名

Zheyou = pd.merge(Zheyou,Chao48y1,on = ['区县'], how = 'left')
Zheyou = Zheyou.fillna(0)
Guzhang1 = pd.merge(Guzhang1,Zheyou,on = ['区县'], how = 'left')



###################################           超10户群障 (不分部门)         *****************************


Guzhangcishu_cs2 = Hsx[ (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='城镇')
Guzhangcishu_cs2 = Guzhangcishu_cs2.append([{'区县':'合计','城镇':Guzhangcishu_cs2.apply(lambda x: x.sum()).城镇}], ignore_index=True)


Guzhangcishu_nc2 = Hsx[(Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10)].groupby(['区县']).size().reset_index(name='农村')
Guzhangcishu_nc2 = Guzhangcishu_nc2.append([{'区县':'合计','农村':Guzhangcishu_nc2.apply(lambda x: x.sum()).农村}], ignore_index=True)
Guzhang2 = pd.merge(Guzhangcishu_cs2,Guzhangcishu_nc2,on = ['区县'], how = 'left')
Guzhang2['线路故障总数'] = Guzhang2.城镇 + Guzhang2.农村

#Dishi = pd.merge(Dishi,Guzhang2,on = ['区县'], how = 'left')
#Guzhang2 = Dishi

#******************     处理时长 ******************
Shichang2 = Hsx[ (Hsx.所属区域 =='城镇') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='城镇时长')
Shichang2 = pd.DataFrame(Shichang2.groupby('区县')['总时长'].sum()).reset_index()
Shichang2 = Shichang2.append([{'区县':'合计','总时长':Shichang2.apply(lambda x: x.sum()).总时长}], ignore_index=True)
#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值
Shichang2['城镇时长'] = Shichang2.总时长 / Guzhang2.城镇
Shichang2 = Shichang2.round({'城镇时长': 2})  # 四舍五入


Shichang_nc2 = Hsx[(Hsx.所属区域 =='农村') & (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='农村时长')
Shichang_nc2 = pd.DataFrame(Shichang_nc2.groupby('区县')['总时长'].sum()).reset_index()
Shichang_nc2 = Shichang_nc2.append([{'区县':'合计','总时长':Shichang_nc2.apply(lambda x: x.sum()).总时长}], ignore_index=True)
#Shichang = Shichang[['区县', '总时长']].groupby(['区县']).mean().reset_index()  # 计算平均值

Shichang2['农村时长'] = Shichang_nc2.总时长 / Guzhang2.农村
Shichang2 = Shichang2.round({'农村时长': 2})  # 四舍五入
Shichang2.drop(['总时长'],axis=1,inplace=True) #删除多余列
Guzhang2 = pd.merge(Guzhang2,Shichang2,on = ['区县'], how = 'left')


#线路故障平均时长
Shichang_gz2 = Hsx[ (Hsx.影响用户数 >=10) ].groupby(['区县','总时长','地点']).size().reset_index(name='故障时长')
quxian2 = Shichang_gz2.groupby(['区县']).size().reset_index(name='数量')
quxian2 = quxian2.append([{'区县':'合计','数量':quxian2.apply(lambda x: x.sum()).数量}], ignore_index=True)

Shichang_gz2 = pd.DataFrame(Shichang_gz2.groupby('区县')['总时长'].sum()).reset_index()
Shichang_gz2 = Shichang_gz2.append([{'区县':'合计','总时长':Shichang_gz2.apply(lambda x: x.sum()).总时长}], ignore_index=True)

Guzhang2['线路故障平均时长'] = Shichang_gz2.总时长 / quxian2.数量
Guzhang2 = Guzhang2.round({'线路故障平均时长': 2})  # 四舍五入
#Shichang_gz1 = Hsx[['区县', '总时长']].groupby(['区县']).mean().reset_index() 计算平均时长

Jishilv6 = Hsx[ (Hsx.所属区域=='农村') & (Hsx.总时长 <=36)].groupby(['区县']).size().reset_index(name='数量')
Jishilv6 = pd.DataFrame(Jishilv6.groupby('区县')['数量'].sum()).reset_index()
Jishilv6 = Jishilv6.append([{'区县':'合计','数量':Jishilv6.apply(lambda x: x.sum()).数量}], ignore_index=True)

Jishilv61 = Hsx[ (Hsx.所属区域=='城镇') & (Hsx.总时长 <=18)].groupby(['区县']).size().reset_index(name='数量')
Jishilv61 = pd.DataFrame(Jishilv61.groupby('区县')['数量'].sum()).reset_index()
Jishilv61 = Jishilv61.append([{'区县':'合计','数量':Jishilv61.apply(lambda x: x.sum()).数量}], ignore_index=True)
Jishilv6['数量1'] = Jishilv6.数量 + Jishilv61.数量


Jishilv7 = Hsx.groupby(['区县']).size().reset_index(name='数量')
Jishilv7 = pd.DataFrame(Jishilv7.groupby('区县')['数量'].sum()).reset_index()
Jishilv7 = Jishilv7.append([{'区县':'合计','数量':Jishilv7.apply(lambda x: x.sum()).数量}], ignore_index=True)
#Jishilv71 = Jishilv7.数量 - 1
#Jishilv71 = pd.DataFrame(Jishilv71)


Guzhang2['整体处理及时率'] = Jishilv6.数量1 / Jishilv7.数量
Guzhang2['整体处理及时率']=Guzhang2['整体处理及时率'].apply(lambda x: '%.2f%%' % (x*100))

#-***************影响用户数
Yonghushu2 = Hsx[(Hsx.影响用户数 >=10) ]
Yonghushu2 = pd.DataFrame(Yonghushu2.groupby('区县')['影响用户数'].sum()).reset_index()
Yonghushu2 = Yonghushu2.append([{'区县':'合计','影响用户数':Yonghushu2.apply(lambda x: x.sum()).影响用户数}], ignore_index=True)
Guzhang2 = pd.merge(Guzhang2,Yonghushu2,on = ['区县'], how = 'left')


#**********今日新增
Jinrizeng2 = Jiyaguzhang[Jiyaguzhang.今日新增时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除

Jinrizeng2 = Jinrizeng2[ (Jinrizeng2.影响用户 >=10) ].groupby(['区县','影响用户']).size().reset_index(name='今日新增')
Jinrizeng2 = pd.DataFrame(Jinrizeng2.groupby('区县')['今日新增'].sum()).reset_index()
Jinrizeng2 = Jinrizeng2.append([{'区县':'合计','今日新增':Jinrizeng2.apply(lambda x: x.sum()).今日新增}], ignore_index=True)

#**********今日恢复
Jinrifu2=Jiyaguzhang[(Jiyaguzhang.恢复时间!='未恢复')]  #删减多余数据
Jinrifu2['恢复时间'] = Jinrifu2['恢复时间'].dt.strftime('%Y/%m/%d')
Jinrifu2 = Jinrifu2[Jinrifu2.恢复时间.str.contains('05/12') ==True]  #筛选4/28的数据，，，，，false为剔除
Jinrifu2 = Jinrifu2[ (Jinrifu2.影响用户 >=10)].groupby(['区县']).size().reset_index(name='今日恢复')
Jinrifu2 = pd.DataFrame(Jinrifu2.groupby('区县')['今日恢复'].sum()).reset_index()
Jinrifu2 = Jinrifu2.append([{'区县':'合计','今日恢复':Jinrifu2.apply(lambda x: x.sum()).今日恢复}], ignore_index=True)

#**********积压故障
Jiya2 = Jiyaguzhang[Jiyaguzhang.恢复时间.str.contains('未恢复') ==True]  #筛选4/28的数据，，，，，false为剔除

Jiya2 = Jiya2[(Jiya2.影响用户 >=10)].groupby(['区县','影响用户']).size().reset_index(name='积压故障')
Jiya2 = pd.DataFrame(Jiya2.groupby('区县')['积压故障'].sum()).reset_index()
Jiya2 = Jiya2.append([{'区县':'合计','积压故障':Jiya2.apply(lambda x: x.sum()).积压故障}], ignore_index=True)

#********超48小时工单数
Chao482 = Jiyaguzhang[ (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48) & (Jiyaguzhang.恢复时间 =='未恢复')].groupby(['区县']).size().reset_index(name='超48小时工单数')
Chao482 = pd.DataFrame(Chao482.groupby('区县')['超48小时工单数'].sum()).reset_index()
Chao482 = Chao482.append([{'区县':'合计','超48小时工单数':Chao482.apply(lambda x: x.sum()).超48小时工单数}], ignore_index=True)

#********超48小时用户数

Chao48y2 = Jiyaguzhang[ (Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.现在时 >=48) & (Jiyaguzhang.恢复时间 =='未恢复')]
Chao48y2 = pd.DataFrame(Chao48y2.groupby('区县')['影响用户'].sum()).reset_index()
Chao48y2 = Chao48y2.append([{'区县':'合计','影响用户':Chao48y2.apply(lambda x: x.sum()).影响用户}], ignore_index=True)
Chao48y2.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名


Guzhang2 = pd.merge(Guzhang2,Jinrizeng2,on = ['区县'], how = 'left')
Guzhang2 = pd.merge(Guzhang2,Jinrifu2,on = ['区县'], how = 'left')
Guzhang2 = pd.merge(Guzhang2,Jiya2,on = ['区县'], how = 'left')
Guzhang2 = pd.merge(Guzhang2,Chao482,on = ['区县'], how = 'left')
Guzhang2 = pd.merge(Guzhang2,Chao48y2,on = ['区县'], how = 'left')
Dishi2 = pd.merge(Dishi2,Guzhang2,on = ['区县'], how = 'left')
Guzhang2 = Dishi2
Guzhang2 = Guzhang2.fillna(0)


################           小表           **********************


Guzhang3 = Dishi3
Guzhang3['昨日新增'] = Guzhang2.今日新增
Guzhang3['昨日遗留'] = Guzhang2.积压故障
Guzhang3['超48小时工单数'] = Guzhang2.超48小时工单数

#********超48小时用户数
Chao48y3 = Jiyaguzhang[(Jiyaguzhang.影响用户 >=10) & (Jiyaguzhang.恢复时间=='未恢复')]
Chao48y3 = pd.DataFrame(Chao48y3.groupby('区县')['影响用户'].sum()).reset_index()
Chao48y3 = Chao48y3.append([{'区县':'合计','影响用户':Chao48y3.apply(lambda x: x.sum()).影响用户}], ignore_index=True)
Chao48y3.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名
Guzhang3 = pd.merge(Guzhang3,Chao48y3,on = ['区县'], how = 'left').fillna(0)




title = pd.DataFrame({'截止6月5日影响用户数超10户群障完成情况（铁通）': ['']})
title1 = pd.DataFrame({'截止6月5日网运（浙邮）线路故障完成情况（影响PON口10户以上）': ['']})
title2 = pd.DataFrame({'截止6月5日影响用户数超10户群障完成情况（整体不分部门）': ['']})
title3 = pd.DataFrame({'每日通报数据': ['']})

with pd.ExcelWriter( todaytime1 + '-崇左群障' + '.xlsx') as writer:  # 写入结果为当前路径
     title.to_excel(writer, sheet_name='群障', startcol=0, index=False, header=True)
     Guzhang.to_excel(writer, sheet_name='群障', startcol=0, startrow=1,index=False, header=True)
     
     title1.to_excel(writer, sheet_name='群障', startcol=0, startrow=15,index=False, header=True)
     Guzhang1.to_excel(writer, sheet_name='群障', startcol=0, startrow=16, index=False, header=True)
     
     title2.to_excel(writer, sheet_name='群障', startcol=0, startrow=30, index=False, header=True)
     Guzhang2.to_excel(writer, sheet_name='群障', startcol=0,startrow=31, index=False, header=True)
     
     title3.to_excel(writer, sheet_name='群障', startcol=0, startrow=46, index=False, header=True)
     Guzhang3.to_excel(writer, sheet_name='群障', startcol=0,startrow=47, index=False, header=True)

# 列距startcol=0, 行距startrow=23
print('******运行结束')




















