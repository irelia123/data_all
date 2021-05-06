# -*- coding: utf-8 -*-
"""
Created on Mon Mar 29 09:29:14 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np
#from prettytable import PrettyTable
######################################################################## 初始化
Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']},
    pd.Index(range(14)))
absolve =  400  ###减免系数   设置为 0 不减免

#######***************************************读取当日在途清单
#读取表单，sheet_name 读取详单里面的第二个表的名称
HSx = pd.read_excel('10086回访模板.xlsx', sheet_name='回访详表（前）')  # 当日导出的在途清单 日期为昨日

#修改列名
HSx.rename(columns={'5.Q4:【无重新报装】好的，我们想了解下是什么原因让您要求工作人员进行取消呢？ （比对异动指标）': '异动指标'}, inplace=True)  # 改列名
HSx.rename(columns={'6.Q5:【有撤单重录】好的，我们想了解下是什么原因让您要求工作人员进行取消呢？': '撤单重录'}, inplace=True)  # 改列名
HSx.rename(columns={'7.Q6:好的，那请问是什么原因未能给您成功办理呢？（比对异动指标原因是否一致）': '比对异动指标'}, inplace=True)  # 改列名
#buyizhi = HSx[HSx.移动指标=='不一致'].groupby(['地市']).size().reset_index(name='不一致')
#筛选三列不一致的数量
buyizhi=HSx[HSx['异动指标'].str.contains('不一致')|HSx['撤单重录'].str.contains('不一致')|HSx['比对异动指标'].str.contains('不一致')]
buyizhi1=pd.merge(Dishi,buyizhi[['地市','异动指标','撤单重录','比对异动指标']],on = ['地市'],how='left')

buyizhi1.rename(columns={'异动指标': '不一致'}, inplace=True)  # 改列名
#所有城市不一致的数量
buyizhi1 = buyizhi1.groupby(['地市']).size().reset_index(name='不一致')
#汇总所有城市不一致的数量
buyizhi1 = buyizhi1.append([{'地市':'全区','不一致':buyizhi1.apply(lambda x: x.sum()).不一致, }], ignore_index=True)   #计算不一致总和
#筛选一致的数量 |代表非
yizhi = HSx[(HSx.异动指标=='1、一致，转Q7')|(HSx.撤单重录=='1、一致，转Q7')|(HSx.比对异动指标=='1、一致，转Q7')]
#yizhi的表从Dishil里面只要三列
yizhi=pd.merge(Dishi,yizhi[['地市','异动指标','撤单重录','比对异动指标']],on = ['地市'],how='left')
#所有城市一致的数量
yizhi.rename(columns={'异动指标': '一致'}, inplace=True)  # 改列名
yizhi = yizhi.groupby(['地市']).size().reset_index(name='一致')
yizhi = yizhi.append([{'地市':'全区','一致':yizhi.apply(lambda x: x.sum()).一致, }], ignore_index=True)   #计算一致总和
zongji = Dishi
zongji=pd.merge(zongji,buyizhi1,on=['地市'],how='left')
zongji=pd.merge(zongji,yizhi,on=['地市'],how='left')
zongji['合计'] = zongji.不一致 + zongji.一致  # 这里是对应城市的总量，新增一列
zongji['退单原因选择正确率'] = zongji.一致 / zongji.合计  # 这里是正确率
zongji['排名']= zongji['退单原因选择正确率'].rank(axis=0, ascending=False)#排名
zongji= zongji.round({'排名': 0})  #排名四舍五入

  #计算全区总和
zongji = zongji.append([{
                        '地市':'全区','不一致':zongji.apply(lambda x: x.sum()).不一致, 
                        '地市':'全区','一致':zongji.apply(lambda x: x.sum()).一致, 
                        '地市':'全区','合计':zongji.apply(lambda x: x.sum()).合计                      
                         }], ignore_index=True)   
zongji['退单原因选择正确率'] = zongji.一致 / zongji.合计  # 这里是计算达标率
zongji['退单原因选择正确率']=zongji['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
zongji = zongji.fillna(' ')  # 批量替换nan，化为空白值
#重新编排表格顺序

 
#zuihou = pd.DataFrame([' '])
#zongji=zongji.append(zuihou,ignore_index=True)  # ignore_index=True,表示不按原来的索引，从0开始自动递增
#zongji['前端']= None
#zongji['前端'].replace('None',' ')
zongji = zongji[['排名','地市','不一致','一致','合计','退单原因选择正确率']]


#qianduan = np.array(['前端','zongji'])



################################   前端2   ##########################################

buyizhi2=HSx[HSx['异动指标'].str.contains('不一致')|HSx['撤单重录'].str.contains('不一致')|HSx['比对异动指标'].str.contains('不一致')]
buyizhi5=pd.merge(Dishi,buyizhi2[['地市','异动指标','撤单重录','比对异动指标']],on = ['地市'],how='left')

buyizhi5.rename(columns={'异动指标': '不一致'}, inplace=True)  # 改列名
#所有城市不一致的数量
buyizhi5 = buyizhi5.groupby(['地市']).size().reset_index(name='不一致')
#汇总所有城市不一致的数量
buyizhi5 = buyizhi5.append([{'地市':'全区','不一致':buyizhi5.apply(lambda x: x.sum()).不一致, }], ignore_index=True)   #计算不一致总和
#筛选一致的数量 |代表非
yizhi2 = HSx[(HSx.异动指标=='1、一致，转Q7')|(HSx.撤单重录=='1、一致，转Q7')|(HSx.比对异动指标=='1、一致，转Q7')]
#yizhi的表从Dishil里面只要三列
yizhi2=pd.merge(Dishi,yizhi2[['地市','异动指标','撤单重录','比对异动指标']],on = ['地市'],how='left')
#所有城市一致的数量
yizhi2.rename(columns={'异动指标': '一致'}, inplace=True)  # 改列名
yizhi2 = yizhi2.groupby(['地市']).size().reset_index(name='一致')
yizhi2 = yizhi2.append([{'地市':'全区','一致':yizhi2.apply(lambda x: x.sum()).一致, }], ignore_index=True)   #计算一致总和

zongji2 = Dishi
zongji2=pd.merge(zongji2,buyizhi5,on=['地市'],how='left')
zongji2=pd.merge(zongji2,yizhi2,on=['地市'],how='left')
zongji2['合计'] = zongji2.不一致 + zongji2.一致  # 这里是对应城市的总量，新增一列
zongji2['退单原因选择正确率'] = zongji2.一致 / zongji2.合计  # 这里是正确率
zongji2['排名']= zongji2['退单原因选择正确率'].rank(axis=0, ascending=False)#排名
zongji2= zongji2.round({'排名': 0})  #排名四舍五入
zongji2 = zongji2.sort_values(by='退单原因选择正确率', ascending=False).reset_index(drop=True)  # 正确率 降序 排序


  #计算全区总和
zongji2 = zongji2.append([{
                        '地市':'全区','不一致':zongji2.apply(lambda x: x.sum()).不一致, 
                        '地市':'全区','一致':zongji2.apply(lambda x: x.sum()).一致, 
                        '地市':'全区','合计':zongji2.apply(lambda x: x.sum()).合计                      
                         }], ignore_index=True)   
    
zongji2['退单原因选择正确率'] = zongji2.一致 / zongji2.合计  # 这里是计算达标率
zongji2['退单原因选择正确率']=zongji2['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
zongji2 = zongji2.fillna(' ')  # 批量替换nan，化为空白值
#重新编排表格顺序
zongji2 = zongji2[['排名','地市','不一致','一致','合计','退单原因选择正确率']]
zongji2 = zongji2.sort_values(by='退单原因选择正确率', ascending=False).reset_index(drop=True)  # 正确率 降序 排序



##################        后端      **********************8*************


Dishi1 = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']},
    pd.Index(range(14)))
HSx1 = pd.read_excel('10086回访模板.xlsx', sheet_name='回访详表（后）')  # 当日导出的在途清单 日期为昨日
#修改列名
HSx1.rename(columns={'5.单选题Q5:XX先生/女士,是这样的，根据我们这边的记录，您的宽带未能成功安装，请问具体是什么原因未能给您成功安装呢？根据驳回原因异动指标显示及客户反馈的原因做比较后选择，如不一致则需选择不一致原因': '异动指标'}, inplace=True)  # 改列名'
#同时满足不一致和调研成功两个条件
houbuyizhi=HSx1[HSx1['异动指标'].str.contains('不一致')&HSx1['结果'].str.contains('调研成功')]
#计算不一致的数量
houbuyizhi = houbuyizhi.groupby(['地市']).size().reset_index(name='不一致')
#汇总全区不一致的总数
houbuyizhi = houbuyizhi.append([{'地市':'全区','不一致':houbuyizhi.apply(lambda x: x.sum()).不一致, }], ignore_index=True)   #计算不一致总和

#计算yizhi
houyizhi = HSx1[(HSx1.异动指标=='1、一致，转Q8')&(HSx1.结果=='调研成功')]
houyizhi = houyizhi.groupby(['地市']).size().reset_index(name='一致')
houyizhi = houyizhi.append([{'地市':'全区','一致':houyizhi.apply(lambda x: x.sum()).一致, }], ignore_index=True)   #计算一致总和

#算总计
houzongji = Dishi1
houzongji=pd.merge(houzongji,houbuyizhi,on=['地市'],how='left')
houzongji=pd.merge(houzongji,houyizhi,on=['地市'],how='left')
houzongji['合计'] = houzongji.不一致 + houzongji.一致
houzongji['退单原因选择正确率'] = houzongji.一致 / houzongji.合计
houzongji['退单原因选择正确率']=houzongji['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
houzongji['排名'] = houzongji['退单原因选择正确率'].rank(axis = 0,ascending = False)
houzongji= houzongji.round({'排名': 0})  #排名四舍五入
#计算全区总和
houzongji = houzongji.append([{
                        '地市':'全区','不一致':houzongji.apply(lambda x: x.sum()).不一致, 
                        '地市':'全区','一致':houzongji.apply(lambda x: x.sum()).一致, 
                        '地市':'全区','合计':houzongji.apply(lambda x: x.sum()).合计                      
                         }], ignore_index=True)   
    
houzongji['退单原因选择正确率'] = houzongji.一致 / houzongji.合计  
houzongji['退单原因选择正确率']=houzongji['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
houzongji = houzongji.fillna(' ')  # 批量替换nan，化为空白值

HSx1.rename(columns={'7.单选题Q7【无需询问，根据客户回答填写】不一致，请选择不一致原因,转Q8': '怎么没有上门'}, inplace=True)  # 改列名'
anzhuang =HSx1[(HSx1.怎么没有上门=='3、装机人员未联系客户，也未上门安装，直接后端退单')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='装机人员没有和客户解释无法安装的原因1')
anzhuang1 =HSx1[(HSx1.怎么没有上门=='2、装机人员没有和客户做好沟通，客户不了解真正安装不了的原因')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='装机人员没有和客户解释无法安装的原因2')
anzhuang=pd.merge(Dishi1,anzhuang,on=['地市'],how='left')
anzhuang=pd.merge(anzhuang,anzhuang1,on=['地市'],how='left')
anzhuang = anzhuang.fillna(0)  # 批量替换nan，化为空白值
anzhuang['装机人员没有和客户解释无法安装的原因'] = anzhuang.装机人员没有和客户解释无法安装的原因1 + anzhuang.装机人员没有和客户解释无法安装的原因2

anzhuang3= anzhuang[['地市','装机人员没有和客户解释无法安装的原因']]

anzhuang3= anzhuang3.append([{'地市':'全区','装机人员没有和客户解释无法安装的原因':anzhuang3.apply(lambda x: x.sum()).装机人员没有和客户解释无法安装的原因}], ignore_index=True)
houzongji =pd.merge(houzongji,anzhuang3,on=['地市'],how='left')


Weizhuang =HSx1[(HSx1.异动指标=='2、不一致，转Q7')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='未减装机人员未按要求选择与实际情况对应的退单原因2')
Weizhuang = Weizhuang.append([{'地市':'全区','未减装机人员未按要求选择与实际情况对应的退单原因2':Weizhuang.apply(lambda x: x.sum()).未减装机人员未按要求选择与实际情况对应的退单原因2, }], ignore_index=True)   #计算总和
anzhuang3 =pd.merge(Weizhuang,anzhuang3,on=['地市'],how='left')

tuidan =HSx1[(HSx1.怎么没有上门=='7、其他，根据客户描述由话务员自行填写_____电信已经装好了，所以不装了')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='未联系客户直接退单或退单理由不成立虚假退单')

tuidan=pd.merge(Dishi1,tuidan,on=['地市'],how='left')

tuidan = tuidan.fillna(0)  # 批量替换nan，化为空白值
tuidan= tuidan.append([{'地市':'全区','未联系客户直接退单或退单理由不成立虚假退单':tuidan.apply(lambda x: x.sum()).未联系客户直接退单或退单理由不成立虚假退单}], ignore_index=True)
anzhuang3 =pd.merge(anzhuang3,tuidan,on=['地市'],how='left')

Weizhuang['未减装机人员未按要求选择与实际情况对应的退单原因'] =Weizhuang.未减装机人员未按要求选择与实际情况对应的退单原因2 - anzhuang3.装机人员没有和客户解释无法安装的原因 - tuidan.未联系客户直接退单或退单理由不成立虚假退单  

haoweizhuang= Weizhuang[['地市','未减装机人员未按要求选择与实际情况对应的退单原因']]
houzongji=pd.merge(houzongji,haoweizhuang,on=['地市'],how='left')
houzongji =pd.merge(houzongji,tuidan,on=['地市'],how='left')
#
houzongji = houzongji[['排名','地市','不一致','一致','合计','退单原因选择正确率','装机人员没有和客户解释无法安装的原因','未减装机人员未按要求选择与实际情况对应的退单原因','未联系客户直接退单或退单理由不成立虚假退单']]



#####################                后端2           *******************************



zuihoubuyizhi=HSx1[HSx1['异动指标'].str.contains('不一致')&HSx1['结果'].str.contains('调研成功')]
#计算不一致的数量
zuihoubuyizhi = zuihoubuyizhi.groupby(['地市']).size().reset_index(name='不一致')
#汇总全区不一致的总数
zuihoubuyizhi = zuihoubuyizhi.append([{'地市':'全区','不一致':zuihoubuyizhi.apply(lambda x: x.sum()).不一致, }], ignore_index=True)   #计算不一致总和

#计算yizhi
zuihouyizhi = HSx1[(HSx1.异动指标=='1、一致，转Q8')&(HSx1.结果=='调研成功')]
zuihouyizhi = zuihouyizhi.groupby(['地市']).size().reset_index(name='一致')
zuihouyizhi = zuihouyizhi.append([{'地市':'全区','一致':zuihouyizhi.apply(lambda x: x.sum()).一致, }], ignore_index=True)   #计算一致总和

#算总计
zuihouzongji = Dishi1
zuihouzongji=pd.merge(zuihouzongji,zuihoubuyizhi,on=['地市'],how='left')
zuihouzongji=pd.merge(zuihouzongji,zuihouyizhi,on=['地市'],how='left')
zuihouzongji['合计'] = zuihouzongji.不一致 + zuihouzongji.一致
zuihouzongji['退单原因选择正确率'] = zuihouzongji.一致 / zuihouzongji.合计
zuihouzongji['退单原因选择正确率']=zuihouzongji['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
zuihouzongji['排名'] = zuihouzongji['退单原因选择正确率'].rank(axis = 0,ascending = False)
zuihouzongji=zuihouzongji.round({'排名': 0})  #排名四舍五入
zuihouzongji = zuihouzongji.sort_values(by='退单原因选择正确率', ascending=False).reset_index(drop=True)  # 正确率 降序 排序

#计算全区总和
zuihouzongji = zuihouzongji.append([{
                        '地市':'全区','不一致':zuihouzongji.apply(lambda x: x.sum()).不一致, 
                        '地市':'全区','一致':zuihouzongji.apply(lambda x: x.sum()).一致, 
                        '地市':'全区','合计':zuihouzongji.apply(lambda x: x.sum()).合计                      
                         }], ignore_index=True)   

zuihouzongji['退单原因选择正确率'] = zuihouzongji.一致 / zuihouzongji.合计  

zuihouzongji['退单原因选择正确率']=zuihouzongji['退单原因选择正确率'].apply(lambda x: '%.2f%%' % (x*100))
zuihouzongji = zuihouzongji.fillna(' ')  # 批量替换nan，化为空白值
HSx1.rename(columns={'7.单选题Q7【无需询问，根据客户回答填写】不一致，请选择不一致原因,转Q8': '怎么没有上门'}, inplace=True)  # 改列名'
zuihouanzhuang =HSx1[(HSx1.怎么没有上门=='3、装机人员未联系客户，也未上门安装，直接后端退单')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='装机人员没有和客户解释无法安装的原因1')
zuihouanzhuang1 =HSx1[(HSx1.怎么没有上门=='2、装机人员没有和客户做好沟通，客户不了解真正安装不了的原因')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='装机人员没有和客户解释无法安装的原因2')
zuihouanzhuang=pd.merge(Dishi1,zuihouanzhuang,on=['地市'],how='left')
zuihouanzhuang=pd.merge(zuihouanzhuang,zuihouanzhuang1,on=['地市'],how='left')
zuihouanzhuang = zuihouanzhuang.fillna(0)  # 批量替换nan，化为空白值
zuihouanzhuang['装机人员没有和客户解释无法安装的原因'] = zuihouanzhuang.装机人员没有和客户解释无法安装的原因1 + zuihouanzhuang.装机人员没有和客户解释无法安装的原因2
zuihouanzhuang3= zuihouanzhuang[['地市','装机人员没有和客户解释无法安装的原因']]
zuihouanzhuang3= zuihouanzhuang3.append([{'地市':'全区','装机人员没有和客户解释无法安装的原因':zuihouanzhuang3.apply(lambda x: x.sum()).装机人员没有和客户解释无法安装的原因}], ignore_index=True)
zuihouzongji =pd.merge(zuihouzongji,zuihouanzhuang3,on=['地市'],how='left')



zuihouWeizhuang =HSx1[(HSx1.异动指标=='2、不一致，转Q7')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='未减装机人员未按要求选择与实际情况对应的退单原因2')
zuihouWeizhuang = zuihouWeizhuang.append([{'地市':'全区','未减装机人员未按要求选择与实际情况对应的退单原因2':zuihouWeizhuang.apply(lambda x: x.sum()).未减装机人员未按要求选择与实际情况对应的退单原因2, }], ignore_index=True)   #计算总和
zuihouanzhuang3 =pd.merge(zuihouWeizhuang,zuihouanzhuang3,on=['地市'],how='left')

zuihoutuidan =HSx1[(HSx1.怎么没有上门=='7、其他，根据客户描述由话务员自行填写_____电信已经装好了，所以不装了')&(HSx1.结果=='调研成功')].groupby(['地市']).size().reset_index(name='未联系客户直接退单或退单理由不成立虚假退单')
zuihoutuidan=pd.merge(Dishi1,zuihoutuidan,on=['地市'],how='left')
zuihoutuidan = zuihoutuidan.fillna(0)  # 批量替换nan，化为空白值
zuihoutuidan= zuihoutuidan.append([{'地市':'全区','未联系客户直接退单或退单理由不成立虚假退单':zuihoutuidan.apply(lambda x: x.sum()).未联系客户直接退单或退单理由不成立虚假退单}], ignore_index=True)
zuihouanzhuang3 =pd.merge(zuihouanzhuang3,zuihoutuidan,on=['地市'],how='left')

zuihouWeizhuang['未减装机人员未按要求选择与实际情况对应的退单原因'] =zuihouWeizhuang.未减装机人员未按要求选择与实际情况对应的退单原因2 - zuihouanzhuang3.装机人员没有和客户解释无法安装的原因 - zuihoutuidan.未联系客户直接退单或退单理由不成立虚假退单  
zuihouhaoweizhuang= zuihouWeizhuang[['地市','未减装机人员未按要求选择与实际情况对应的退单原因']]

zuihouzongji=pd.merge(zuihouzongji,zuihouhaoweizhuang,on=['地市'],how='left')
zuihouzongji =pd.merge(zuihouzongji,zuihoutuidan,on=['地市'],how='left')
#
zuihouzongji = zuihouzongji[['排名','地市','不一致','一致','合计','退单原因选择正确率','装机人员没有和客户解释无法安装的原因','未减装机人员未按要求选择与实际情况对应的退单原因','未联系客户直接退单或退单理由不成立虚假退单']]

zuihouzongji = zuihouzongji.sort_values(by='退单原因选择正确率', ascending=False).reset_index(drop=True)  # 正确率 降序 排序


with pd.ExcelWriter('10086回访统计' + '.xlsx') as writer:  # 写入结果为当前路径
    zongji.to_excel(writer, sheet_name='前端', startcol=0, index=False, header=True)
    zongji2.to_excel(writer, sheet_name='前端2', startcol=0, index=False, header=True)
    houzongji.to_excel(writer, sheet_name='后端', startcol=0, index=False, header=True)
    zuihouzongji.to_excel(writer, sheet_name='后端2', startcol=0, index=False, header=True)
