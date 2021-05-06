# -*- coding: utf-8 -*-
"""
Created on Fri Mar 26 11:47:43 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np



#模糊查询9和7开头的小微宽带以及开头投诉分类为'集团业务->网络质量->小微宽带'的宽带，拼接。
Xw_Zb = pd.read_excel('file:///D:/工程量/小微/小微宽带03+04(1).xlsx',dtype=object)
Xw_Qd = Xw_Zb[['工单类型','工单流水号','客户电话','宽带账号','用户品质','区域类型','重派次数','工单状态','催办次数','客服催办次数','首次到单时间','首次响应时间','首次联系客户时间','广西时限','集团时限','客服流水号','责任地市','责任区县','责任代维公司','责任代维班组','投诉分类','账号所属公司']]
#aaaa = Xw_Qd['客户电话']

Xw_data1=Xw_Qd[(Xw_Qd.首次联系客户时间!='nan')]  #删减多余数据
Xw_data1=Xw_data1.fillna('') #  批量替换nan


Xw_data1 = Xw_Qd[(Xw_Qd.工单类型=='家庭宽带')|(Xw_Qd.工单类型=='企业宽带')]
#Xw_data1['客户电话'] = pd.to_numeric(Xw_data1['客户电话']).round(0).astype(str)       #转换客户电话为str类型a
Mask =Xw_data1[(Xw_data1['客户电话'].str.match('9.*'))|(Xw_data1['客户电话'].str.match('7.*'))|(Xw_data1['投诉分类'].str.match('集团业务->网络质量->小微宽带.*'))]  #数值的要转换类型为str才能模糊查询！！
Mask=Mask[(Mask.账号所属公司!='铁通割接')]  #删减多余数据

Mask=Mask.fillna('') #  批量替换nan


Xiaowei = Mask
Xiaowei = Xiaowei.applymap(str)
#Xiaowei = pd.concat([Zb, Mask], axis=1)     #拼接单列合并到多列不改列名
Xiaowei['地市'] = Xiaowei['责任地市']
Xiaowei['区域类型1'] = Mask['区域类型']
Xiaowei['工单开始时间'] = Xiaowei['首次到单时间']
todaytime = datetime.datetime.strptime('2021-4-9','%Y-%m-%d')      #手动输入日期
#time20 = datetime.datetime.strptime('20:00:00','%H:%M:%S')      #手动输入日期
#time16 = datetime.datetime.strptime('16:00:00','%H:%M:%S')      #手动输入日期
#Xiaowei['广西时限'] = Xiaowei['广西时限'].apply(pd.to_datetime)
#Xiaowei['广西时限'] = Xiaowei['广西时限'].dt.strftime('%H:%M:%S') #实现全列修改时间格式
#time20 = time20.strftime('%H') + ':' + time20.strftime('%M') + ':' + time20.strftime('%S')  # 日期转换str：20:00:00
#time16 = time16.strftime('%H') + ':' + time16.strftime('%M') + ':' + time16.strftime('%S')  # 日期转换str：20:00:00
Xiaowei['当前日期'] = (pd.to_datetime(todaytime)-pd.to_datetime(Xiaowei.首次到单时间)).dt.days*24+(pd.to_datetime(todaytime)-pd.to_datetime(Xiaowei.首次到单时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改


Xiaowei['是否超七天'] = Xiaowei.当前日期 /24 >= 168   #时间小于7天    备注：公式是时间算的不能按元Excel数据修改
Xiaowei['是否超48小时'] = Xiaowei.当前日期 >= 48   #时间小于48小时    备注：公式是时间算的不能按元Excel数据修改
Xiaowei['催单次数'] = (Xiaowei.催办次数 + Xiaowei.客服催办次数).map(int) 

todaytime = datetime.datetime.strptime('2021-4-9','%Y-%m-%d')      #手动输入日期
#todaytime当天时间 - todaytime（格式）转化为小时
Xiaowei['是否超时'] = (todaytime - pd.to_datetime(Xiaowei.广西时限)).dt.days*24 + (todaytime - pd.to_datetime(Xiaowei.广西时限)).dt.seconds/3600 +7

#Xiaowei['是否超时'] = (pd.to_datetime(todaytime)-pd.to_datetime(Xiaowei.广西时限)).dt.days*24+(pd.to_datetime(todaytime)-pd.to_datetime(Xiaowei.广西时限)).dt.seconds/3600+7     #计算时间公式    备注：公式是时间算的不能按元Excel数据修改

Xiaowei['tag1'] = ' ' #新增tag1列
Xiaowei['tag1'][Xiaowei['当前日期'] >= 168] = '超7日未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Xiaowei['tag1'][(Xiaowei['当前日期'] < 168)&(Xiaowei['当前日期'] >= 96)] = '超4-7日未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Xiaowei['tag1'][(Xiaowei['当前日期'] < 96)&(Xiaowei['当前日期'] >= 48)] = '超2-4日未竣工'     #时间小于四天=超2-4日未竣工    备注：公式是时间算的不能按元Excel数据修改
Xiaowei['tag1'][Xiaowei['当前日期'] < 48] = '超2-4日未竣工'     #时间小于四天=超2-4日未竣工

Xiaowei['是否待质检'] = Xiaowei.工单状态 == '质检' #判断是否等于质检，返回true或false

#可以使用DF.Units = DF.Units.map(int)或DF.Units = DF.Units.astype(int)
Xiaowei['被客服催单'] = (Xiaowei.客服催办次数).map(int) - 1 >= 0
Xiaowei['是否多次驳回'] = (Xiaowei.重派次数).map(int) - 2 >= 0
Xiaowei['是否签单'] = (Xiaowei.首次响应时间).map(len) > 1 #返回对应的类型len



#重新编排表格顺序
Xiaowei = Xiaowei[['地市','是否超七天','是否超48小时','是否多次驳回','工单开始时间','区域类型','当前日期','tag1','是否签单',
                       '是否待质检','催单次数','被客服催单','是否超时','工单类型','工单流水号','客户电话',
                       '宽带账号','用户品质','区域类型1','重派次数','工单状态','催办次数','客服催办次数','首次到单时间','首次响应时间',
                       '首次联系客户时间','广西时限','集团时限','客服流水号','责任地市','责任区县','责任代维公司','责任代维班组']]



with pd.ExcelWriter('1小微投诉在途清单' + '.xlsx') as writer:  # 写入结果为当前路径
     Xiaowei.to_excel(writer, sheet_name='小微投诉在途清单', startcol=0, index=False, header=True)






























