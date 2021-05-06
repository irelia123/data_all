# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 16:05:02 2021

@author: Administrator
"""

import os, glob
import time, datetime
import pandas as pd
import numpy as np


HSx = pd.read_excel('./0427在途.xlsx')  # 当日导出的在途清单 日期为昨日
#frist_day = datetime.datetime.strptime('2021-4-17','%Y-%m-%d')      #手动输入日期
todaytime=datetime.datetime.now()  #系统取当天日期
first_day = datetime.datetime(todaytime.year, todaytime.month, 27, 23, 59, 59)  # datetime类型 2019-09-01 00:00:00

HSx['安装地址'] = ''
HSx['投诉内容'] = ''
HSx['全球通身份等级'] = ''
HSx['客户星级'] = ''
HSx = HSx[(HSx.工单类型 !='和家亲')&(HSx.工单类型 !='问题解决单')]
HSx = HSx.fillna('')

#HSx['重派次数'][HSx['重派次数'] ==' '] =0
#HSx['客服催办次数'][HSx['客服催办次数'] ==' '] =0
#HSx['催办次数'][HSx['催办次数'] ==' '] =0


Charu = HSx[['责任地市']]
Charu['工单流水号'] = HSx[['工单流水号']]
Charu.rename(columns={'责任地市':'地市'}, inplace = True)   #责任地市改名为地市
Charu['是否多次驳回'] = (HSx.重派次数).map(int) - 2 >= 0
Charu['工单开始时间'] = HSx.首次到单时间
Charu[first_day] = (first_day-pd.to_datetime(Charu.工单开始时间)).dt.days*24 + (first_day-pd.to_datetime(Charu.工单开始时间)).dt.seconds/3600     #计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
Charu.ix[(HSx.工单状态 !='待归档'),'是否要剔除'] = 'False' #新增判断的一列
#Charu.ix[(HSx.工单类型 ==0)|(HSx.工单类型 =='家庭宽带'),'是否要剔除'] = 'False' #新增判断的一列
Charu = Charu[Charu.是否要剔除=='False']
Charu['催单次数'] = (HSx.催办次数 + HSx.客服催办次数).map(int) 
Charu['被客服催单'] = (HSx.客服催办次数 - 1).map(int) >=0
#*****************新增判断的一列
Charu.ix[(Charu[first_day] >120),'tag1'] = '超5天未竣工'
Charu.ix[(Charu[first_day] >168),'tag1'] = '超7天未竣工'
Charu.ix[(Charu[first_day] >360),'tag1'] = '超15天未竣工'
Charu.ix[(Charu[first_day] >720),'tag1'] = '超30天未竣工'
Charu = Charu.fillna('')
Charu=Charu[(Charu.tag1!='')]  #删减空白数据
Charu = Charu.sort_values(by='地市', ascending=True).reset_index(drop=True)  # 按地市 排序
HSx = pd.merge(Charu, HSx, on=['工单流水号'], how='left')  # 拼接 超72小时数量

Cols = HSx.columns.tolist()                     # 把Cols1的列名称，取出来放到一个list里边。即返回['a', 'b', 'c', 'd', 'e', '责任地市']
Cols.insert(9, Cols.pop(Cols.index('工单流水号')))        # pop()把工单开始时间从cols列表里挖出来，通过位置参数“0”，然后放到第一列。
HSx = HSx[Cols]      #排列后数据


with pd.ExcelWriter('全流程时长管控' + '.xlsx') as writer:  # 写入结果为当前路径
     HSx.to_excel(writer, sheet_name='全流程时长管控清单', startcol=0, index=False, header=True)





