import os, glob
import time, datetime
import pandas as pd
import numpy as np
##初始化
today=datetime.datetime.now()
today1=str(today.month) + '月' + today.strftime('%d') + '日'
Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))
HD1 = pd.read_excel('16回单质量检测指标地市报表详情.xlsx')
if list(HD1)[0] == '查询结果':  ###修改表头
    list0 = list(HD1.iloc[0])
    ###print('******请检查新的表头是否正确：',list0)
    HD1.columns = list0  ###重命名表头
    HD1.dropna(subset=['归档时间'], inplace=True)  # ,inplace=True  删除空行
    HD1 = HD1.drop(HD1[HD1.归档时间 == '归档时间'].index)  # 删除多余行
data1=HD1[(HD1.工单主题!='家庭宽带-1000M-FTTB-移机-正装机')]  #删减多余数据

c1 = data1[data1.区域类型=='城镇'].groupby(['地市']).size().reset_index(name='城镇总工单数')
c1 = c1.append([{'地市': '全区', '城镇总工单数': c1.apply(lambda x: x.sum()).城镇总工单数}], ignore_index=True)

c2 = data1[(data1.测速=='达标')&(data1.区域类型=='城镇')].groupby(['地市']).size().reset_index(name='整改前测速达标数')
c2 = c2.append([{'地市': '全区', '整改前测速达标数': c2.apply(lambda x: x.sum()).整改前测速达标数}], ignore_index=True)
c1 = pd.merge(c1, c2, on=['地市'], how='left') #拼接

c3 = data1[(data1.最终测速结果=='达标') & (data1.区域类型=='城镇')].groupby(['地市']).size().reset_index(name='整改后测速达标数')
c3 = c3.append([{'地市': '全区', '整改后测速达标数': c3.apply(lambda x: x.sum()).整改后测速达标数}], ignore_index=True)
c1 = pd.merge(c1, c3, on=['地市'], how='left') #拼接

c1['整改条数']=c1.整改后测速达标数-c1.整改前测速达标数

c1['整改后测速达标率']=c1.整改后测速达标数 / c1.城镇总工单数


c1['整改前测速达标率']=c1.整改前测速达标数/c1.城镇总工单数
c1['整改后测速达标率']=c1.整改后测速达标数/c1.城镇总工单数

c1['排名']=c1['整改后测速达标率'].rank(axis=0,ascending=False,method='dense')   # 输出排名

c1 = c1.round({'排名': 0})  # 四舍五入



c1['整改前测速达标率']=c1['整改前测速达标率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
c1['整改后测速达标率']=c1['整改后测速达标率'].apply(lambda x: '%.2f%%' % (x * 100))

c1['整改条数']=c1.整改后测速达标数-c1.整改前测速达标数
c1['整改提升百分比']=(c1.整改后测速达标数/c1.城镇总工单数)-(c1.整改前测速达标数/c1.城镇总工单数)
c1['整改提升百分比'] = c1['整改提升百分比'].apply(lambda x: '%.2f%%' % (x * 100))

c1=c1[['地市','城镇总工单数','整改前测速达标数','整改后测速达标数','整改条数','整改前测速达标率','整改后测速达标率','排名','整改提升百分比']]

with pd.ExcelWriter('程序生成报表2021年'+today1+'千兆速率不达标通报' +'.xlsx') as writer:
    c1.to_excel(writer, sheet_name='总工单数', startcol=0, index=False, header=True)
