#导入模块.
import os, glob
import time, datetime
import pandas as pd
import numpy as np
#################################
today=datetime.datetime.now()#实时
today2= today.strftime('%Y') + '年' + str(today.month) + '月' + today.strftime('%d') + '日'  # 日期转换str：2019年9月10日
#计算日期完毕!
Bymbh = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

MBH=pd.read_excel('魔百和装机及时率报表.xls')
######################################
if list(MBH)[0] == '查询结果' :  ###修改表头
    list0 =list(MBH.iloc[0])
    # print('请检查新的表头是否正确：',list0)
    MBH.columns = list0  ###重命名表头
    MBH = MBH.drop(MBH[MBH.魔百和装机及时率报表 == '魔百和装机及时率报表'].index)  # 删除多余行
#######################
if list(MBH)[0] == '魔百和装机及时率报表' :  ###修改表头
    list0 =list(MBH.iloc[0])
#     # print('请检查新的表头是否正确：',list0)
    MBH.columns = list0  ###重命名表头
    MBH = MBH.drop(MBH[MBH.地市 == '地市'].index)  # 删除多余行
MBH=MBH.reset_index(drop=True)
###############
MBH.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
MBH.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
MBH.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
MBH.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
MBH.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
MBH.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
MBH1=MBH.drop(index=(MBH.loc[(MBH['地市']=='总计')].index))
####准备完毕
MBH_CGtime=MBH1[['地市','城镇高品质装机时长']]
MBH_NGtime=MBH1[['地市','农村高品质装机时长']]
MBH_Pcitytime=MBH1[['地市','城镇普通品质装机时长']]
MBH_PNCtime=MBH1[['地市','农村普通品质装机时长']]
MBH_citytime=MBH1[['地市','城镇装机时长']]
MBH_NCtime=MBH1[['地市','农村装机时长']]
MBH_time=MBH1[['地市','整体装机时长']]
MBH_GJS=MBH1[['地市','高品质装机及时率']]
MBH_PJS=MBH1[['地市','普通品质装机及时率']]
MBH_JS=MBH1[['地市','整体装机及时率']]

MBHdata=pd.merge(MBH_CGtime, MBH_NGtime, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_Pcitytime, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_PNCtime, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_citytime, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_NCtime, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_time, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_GJS, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_PJS, on=['地市'], how='left')##拼接
MBHdata=pd.merge(MBHdata,MBH_JS, on=['地市'], how='left')##拼接
     #获取需要数据
###########################################################
IMS=pd.read_excel('IMS装机及时率报表.xls')
######################################
if list(IMS)[0] == '查询结果' :  ###修改表头
    list0 =list(IMS.iloc[0])
    # print('请检查新的表头是否正确：',list0)
    IMS.columns = list0  ###重命名表头
    IMS = IMS.drop(IMS[IMS.IMS装机及时率报表 == 'IMS装机及时率报表'].index)  # 删除多余行
#######################
if list(IMS)[0] == 'IMS装机及时率报表' :  ###修改表头
    list0 =list(IMS.iloc[0])
#     # print('请检查新的表头是否正确：',list0)
    IMS.columns = list0  ###重命名表头
    IMS = IMS.drop(IMS[IMS.地市 == '地市'].index)  # 删除多余行
IMS=IMS.reset_index(drop=True)
###############
IMS.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
IMS.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
IMS.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
IMS.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
IMS.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
IMS.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
IMS1=IMS.drop(index=(IMS.loc[(IMS['地市']=='总计')].index))
####准备完毕
IMS_CGtime=IMS1[['地市','城镇高品质装机时长']]
IMS_NGtime=IMS1[['地市','农村高品质装机时长']]
IMS_Pcitytime=IMS1[['地市','城镇普通品质装机时长']]
IMS_PNCtime=IMS1[['地市','农村普通品质装机时长']]
IMS_citytime=IMS1[['地市','城镇装机时长']]
IMS_NCtime=IMS1[['地市','农村装机时长']]
IMS_time=IMS1[['地市','整体装机时长']]
IMS_GJS=IMS1[['地市','高品质装机及时率']]
IMS_PJS=IMS1[['地市','普通品质装机及时率']]
IMS_JS=IMS1[['地市','整体装机及时率']]

IMSdata=pd.merge(IMS_CGtime, IMS_NGtime, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_Pcitytime, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_PNCtime, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_citytime, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_NCtime, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_time, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_GJS, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_PJS, on=['地市'], how='left')##拼接
IMSdata=pd.merge(IMSdata,IMS_JS, on=['地市'], how='left')##拼接
     #获取需要数据
###########################################################
HM=pd.read_excel('和目装机及时率报表.xls')
######################################
if list(HM)[0] == '查询结果' :  ###修改表头
    list0 =list(HM.iloc[0])
    # print('请检查新的表头是否正确：',list0)
    HM.columns = list0  ###重命名表头
    HM = HM.drop(HM[HM.和目装机及时率报表 == '和目装机及时率报表'].index)  # 删除多余行
#######################
if list(HM)[0] == '和目装机及时率报表' :  ###修改表头
    list0 =list(HM.iloc[0])
#     # print('请检查新的表头是否正确：',list0)
    HM.columns = list0  ###重命名表头
    HM = HM.drop(HM[HM.地市 == '地市'].index)  # 删除多余行
HM=HM.reset_index(drop=True)
###############
HM.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
HM.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
HM.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
HM.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
HM.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
HM.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
HM1=HM.drop(index=(HM.loc[(HM['地市']=='总计')].index))
####准备完毕
HM_CGtime=HM1[['地市','城镇高品质装机时长']]
HM_NGtime=HM1[['地市','农村高品质装机时长']]
HM_Pcitytime=HM1[['地市','城镇普通品质装机时长']]
HM_PNCtime=HM1[['地市','农村普通品质装机时长']]
HM_citytime=HM1[['地市','城镇装机时长']]
HM_NCtime=HM1[['地市','农村装机时长']]
HM_time=HM1[['地市','整体装机时长']]
HM_GJS=HM1[['地市','高品质装机及时率']]
HM_PJS=HM1[['地市','普通品质装机及时率']]
HM_JS=HM1[['地市','整体装机及时率']]

HMdata=pd.merge(HM_CGtime, HM_NGtime, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_Pcitytime, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_PNCtime, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_citytime, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_NCtime, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_time, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_GJS, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_PJS, on=['地市'], how='left')##拼接
HMdata=pd.merge(HMdata,HM_JS, on=['地市'], how='left')##拼接
     #获取需要数据
###########################################################
PA=pd.read_excel('平安乡村及时率报表.xls')
######################################
if list(PA)[0] == '查询结果' :  ###修改表头
    list0 =list(PA.iloc[0])
    # print('请检查新的表头是否正确：',list0)
    PA.columns = list0  ###重命名表头
    PA = PA.drop(PA[PA.平安乡村及时率报表 == '平安乡村及时率报表'].index)  # 删除多余行
#######################
if list(PA)[0] == '平安乡村及时率报表' :  ###修改表头
    list0 =list(PA.iloc[0])
#     # print('请检查新的表头是否正确：',list0)
    PA.columns = list0  ###重命名表头
    PA = PA.drop(PA[PA.地市 == '地市'].index)  # 删除多余行
PA=PA.reset_index(drop=True)
###############
PA.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
PA.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
PA.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
PA.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
PA.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
PA.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
PA1=PA.drop(index=(PA.loc[(PA['地市']=='总计')].index))
####准备完毕
PA_CGtime=PA1[['地市','城镇高品质装机时长']]
PA_NGtime=PA1[['地市','农村高品质装机时长']]
PA_Pcitytime=PA1[['地市','城镇普通品质装机时长']]
PA_PNCtime=PA1[['地市','农村普通品质装机时长']]
PA_citytime=PA1[['地市','城镇装机时长']]
PA_NCtime=PA1[['地市','农村装机时长']]
PA_time=PA1[['地市','整体装机时长']]
PA_GJS=PA1[['地市','高品质装机及时率']]
PA_PJS=PA1[['地市','普通品质装机及时率']]
PA_JS=PA1[['地市','整体装机及时率']]

PAdata=pd.merge(PA_CGtime, PA_NGtime, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_Pcitytime, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_PNCtime, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_citytime, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_NCtime, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_time, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_GJS, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_PJS, on=['地市'], how='left')##拼接
PAdata=pd.merge(PAdata,PA_JS, on=['地市'], how='left')##拼接
     #获取需要数据
###########################################################
ZW=pd.read_excel('智能组网装机及时率报表.xls')
######################################
if list(ZW)[0] == '查询结果' :  ###修改表头
    list0 =list(ZW.iloc[0])
    # print('请检查新的表头是否正确：',list0)
    ZW.columns = list0  ###重命名表头
    ZW = ZW.drop(ZW[ZW.智能组网装机及时率报表 == '智能组网装机及时率报表'].index)  # 删除多余行
#######################
if list(ZW)[0] == '智能组网装机及时率报表' :  ###修改表头
    list0 =list(ZW.iloc[0])
#     # print('请检查新的表头是否正确：',list0)
    ZW.columns = list0  ###重命名表头
    ZW = ZW.drop(ZW[ZW.地市 == '地市'].index)  # 删除多余行
ZW=ZW.reset_index(drop=True)
###############
ZW.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
ZW.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
ZW.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
ZW.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
ZW.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
ZW.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
ZW1=ZW.drop(index=(ZW.loc[(ZW['地市']=='总计')].index))
####准备完毕
ZW_CGtime=ZW1[['地市','城镇高品质装机时长']]
ZW_NGtime=ZW1[['地市','农村高品质装机时长']]
ZW_Pcitytime=ZW1[['地市','城镇普通品质装机时长']]
ZW_PNCtime=ZW1[['地市','农村普通品质装机时长']]
ZW_citytime=ZW1[['地市','城镇装机时长']]
ZW_NCtime=ZW1[['地市','农村装机时长']]
ZW_time=ZW1[['地市','整体装机时长']]
ZW_GJS=ZW1[['地市','高品质装机及时率']]
ZW_PJS=ZW1[['地市','普通品质装机及时率']]
ZW_JS=ZW1[['地市','整体装机及时率']]

ZWdata=pd.merge(ZW_CGtime, ZW_NGtime, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_Pcitytime, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_PNCtime, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_citytime, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_NCtime, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_time, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_GJS, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_PJS, on=['地市'], how='left')##拼接
ZWdata=pd.merge(ZWdata,ZW_JS, on=['地市'], how='left')##拼接
     #获取需要数据
###########################################################
with pd.ExcelWriter('筛选数据'+'.xlsx') as writer:
    MBHdata.to_excel(writer, sheet_name='魔百盒数据', startcol=0, index=False, header=True)
    IMSdata.to_excel(writer, sheet_name='IMS数据', startcol=0, index=False, header=True)
    HMdata.to_excel(writer, sheet_name='和目数据', startcol=0, index=False, header=True)
    PAdata.to_excel(writer, sheet_name='平安乡村数据', startcol=0, index=False, header=True)
    ZWdata.to_excel(writer, sheet_name='智能组网数据', startcol=0, index=False, header=True)
################################
################################
################################     开始制作表格
shuju= pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

MBHda=pd.read_excel('筛选数据.xlsx',sheet_name='魔百盒数据')
MBHZJ= MBH.query('地市 == "总计"')
MBHZJ=MBHZJ.reset_index(drop=True)
MBHZJ1=MBHZJ[['城镇高品质装机时长']]
MBHZJ2=MBHZJ[['农村高品质装机时长']]
MBHZJ3=MBHZJ[['城镇普通品质装机时长']]
MBHZJ4=MBHZJ[['农村普通品质装机时长']]
MBHZJ5=MBHZJ[['城镇装机时长']]
MBHZJ6=MBHZJ[['农村装机时长']]
MBHZJ7=MBHZJ[['整体装机时长']]
MBHZJ8=MBHZJ[['高品质装机及时率']]
MBHZJ9=MBHZJ[['普通品质装机及时率']]
MBHZJ10=MBHZJ[['整体装机及时率']]
                #####数据整合完毕
######################################################
MBH_CGtimeda=MBHda[['地市','城镇高品质装机时长']]
MBH_CGtimeda['排名']=MBHda['城镇高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_CGtimeda=pd.concat([MBH_CGtimeda,MBHZJ1]) #拼接


######################################################
MBH_NGtimeda=MBHda[['地市','农村高品质装机时长']]
MBH_NGtimeda['排名']=MBHda['农村高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_NGtimeda=pd.concat([MBH_NGtimeda,MBHZJ2]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_NGtimeda, on=['地市'], how='left')  #拼接


#################################################
MBH_Pcitytimeda=MBHda[['地市','城镇普通品质装机时长']]
MBH_Pcitytimeda['排名']=MBHda['城镇普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_Pcitytimeda=pd.concat([MBH_Pcitytimeda,MBHZJ3]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_Pcitytimeda, on=['地市'], how='left')  #拼接

################################################
MBH_PNCtimeda=MBHda[['地市','农村普通品质装机时长']]
MBH_PNCtimeda['排名']=MBHda['农村普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_PNCtimeda=pd.concat([MBH_PNCtimeda,MBHZJ4]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_PNCtimeda, on=['地市'], how='left')  #拼接

####################################################
MBH_citytimeda=MBHda[['地市','城镇装机时长']]
MBH_citytimeda['排名']=MBHda['城镇装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_citytimeda=pd.concat([MBH_citytimeda,MBHZJ5]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_citytimeda, on=['地市'], how='left')  #拼接

#####################################################
MBH_NCtimeda=MBHda[['地市','农村装机时长']]
MBH_NCtimeda['排名']=MBHda['农村装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_NCtimeda=pd.concat([MBH_NCtimeda,MBHZJ6]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_NCtimeda, on=['地市'], how='left')  #拼接

#####################################################
MBH_timeda=MBHda[['地市','整体装机时长']]
MBH_timeda['排名']=MBHda['整体装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
MBH_timeda=pd.concat([MBH_timeda,MBHZJ7]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_timeda, on=['地市'], how='left')  #拼接

###################################################
MBH_GJSda=MBHda[['地市','高品质装机及时率']]
MBH_GJSda['排名']=MBHda['高品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
MBH_GJSda=pd.concat([MBH_GJSda,MBHZJ8]) #拼接
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_GJSda, on=['地市'], how='left')  #拼接

#################################################
MBH_PJSda=MBHda[['地市','普通品质装机及时率']]
MBH_PJSda['排名']=MBHda['普通品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
MBH_PJSda=pd.concat([MBH_PJSda,MBHZJ9])
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_PJSda, on=['地市'], how='left')  #拼接

#################################################
MBH_JSda=MBHda[['地市','整体装机及时率']]
MBH_JSda['排名']=MBHda['整体装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
MBH_JSda=pd.concat([MBH_JSda,MBHZJ10])
MBH_CGtimeda=pd.merge(MBH_CGtimeda, MBH_JSda, on=['地市'], how='left')  #拼接

MBH_CGtimeda.rename(columns={'排名_y': '排名'}, inplace=True)
MBH_CGtimeda.rename(columns={'排名_x': '排名'}, inplace=True)
##########################################
IMSda=pd.read_excel('筛选数据.xlsx',sheet_name='IMS数据')
IMSZJ= IMS.query('地市 == "总计"')
IMSZJ=IMSZJ.reset_index(drop=True)
IMSZJ1=IMSZJ[['城镇高品质装机时长']]
IMSZJ2=IMSZJ[['农村高品质装机时长']]
IMSZJ3=IMSZJ[['城镇普通品质装机时长']]
IMSZJ4=IMSZJ[['农村普通品质装机时长']]
IMSZJ5=IMSZJ[['城镇装机时长']]
IMSZJ6=IMSZJ[['农村装机时长']]
IMSZJ7=IMSZJ[['整体装机时长']]
IMSZJ8=IMSZJ[['高品质装机及时率']]
IMSZJ9=IMSZJ[['普通品质装机及时率']]
IMSZJ10=IMSZJ[['整体装机及时率']]
                #####数据整合完毕
######################################################
IMS_CGtimeda=IMSda[['地市','城镇高品质装机时长']]
IMS_CGtimeda['排名']=IMSda['城镇高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_CGtimeda=pd.concat([IMS_CGtimeda,IMSZJ1]) #拼接


######################################################
IMS_NGtimeda=IMSda[['地市','农村高品质装机时长']]
IMS_NGtimeda['排名']=IMSda['农村高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_NGtimeda=pd.concat([IMS_NGtimeda,IMSZJ2]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_NGtimeda, on=['地市'], how='left')  #拼接


#################################################
IMS_Pcitytimeda=IMSda[['地市','城镇普通品质装机时长']]
IMS_Pcitytimeda['排名']=IMSda['城镇普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_Pcitytimeda=pd.concat([IMS_Pcitytimeda,IMSZJ3]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_Pcitytimeda, on=['地市'], how='left')  #拼接

################################################
IMS_PNCtimeda=IMSda[['地市','农村普通品质装机时长']]
IMS_PNCtimeda['排名']=IMSda['农村普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_PNCtimeda=pd.concat([IMS_PNCtimeda,IMSZJ4]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_PNCtimeda, on=['地市'], how='left')  #拼接

####################################################
IMS_citytimeda=IMSda[['地市','城镇装机时长']]
IMS_citytimeda['排名']=IMSda['城镇装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_citytimeda=pd.concat([IMS_citytimeda,IMSZJ5]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_citytimeda, on=['地市'], how='left')  #拼接

#####################################################
IMS_NCtimeda=IMSda[['地市','农村装机时长']]
IMS_NCtimeda['排名']=IMSda['农村装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_NCtimeda=pd.concat([IMS_NCtimeda,IMSZJ6]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_NCtimeda, on=['地市'], how='left')  #拼接

#####################################################
IMS_timeda=IMSda[['地市','整体装机时长']]
IMS_timeda['排名']=IMSda['整体装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
IMS_timeda=pd.concat([IMS_timeda,IMSZJ7]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_timeda, on=['地市'], how='left')  #拼接

###################################################
IMS_GJSda=IMSda[['地市','高品质装机及时率']]
IMS_GJSda['排名']=IMSda['高品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
IMS_GJSda=pd.concat([IMS_GJSda,IMSZJ8]) #拼接
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_GJSda, on=['地市'], how='left')  #拼接

#################################################
IMS_PJSda=IMSda[['地市','普通品质装机及时率']]
IMS_PJSda['排名']=IMSda['普通品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
IMS_PJSda=pd.concat([IMS_PJSda,IMSZJ9])
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_PJSda, on=['地市'], how='left')  #拼接

#################################################
IMS_JSda=IMSda[['地市','整体装机及时率']]
IMS_JSda['排名']=IMSda['整体装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
IMS_JSda=pd.concat([IMS_JSda,IMSZJ10])
IMS_CGtimeda=pd.merge(IMS_CGtimeda, IMS_JSda, on=['地市'], how='left')  #拼接

IMS_CGtimeda.rename(columns={'排名_y': '排名'}, inplace=True)
IMS_CGtimeda.rename(columns={'排名_x': '排名'}, inplace=True)
##########################################
HMda=pd.read_excel('筛选数据.xlsx',sheet_name='和目数据')
HMZJ= HM.query('地市 == "总计"')
HMZJ=HMZJ.reset_index(drop=True)
HMZJ1=HMZJ[['城镇高品质装机时长']]
HMZJ2=HMZJ[['农村高品质装机时长']]
HMZJ3=HMZJ[['城镇普通品质装机时长']]
HMZJ4=HMZJ[['农村普通品质装机时长']]
HMZJ5=HMZJ[['城镇装机时长']]
HMZJ6=HMZJ[['农村装机时长']]
HMZJ7=HMZJ[['整体装机时长']]
HMZJ8=HMZJ[['高品质装机及时率']]
HMZJ9=HMZJ[['普通品质装机及时率']]
HMZJ10=HMZJ[['整体装机及时率']]
                #####数据整合完毕
######################################################
HM_CGtimeda=HMda[['地市','城镇高品质装机时长']]
HM_CGtimeda['排名']=HMda['城镇高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_CGtimeda=pd.concat([HM_CGtimeda,HMZJ1]) #拼接


######################################################
HM_NGtimeda=HMda[['地市','农村高品质装机时长']]
HM_NGtimeda['排名']=HMda['农村高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_NGtimeda=pd.concat([HM_NGtimeda,HMZJ2]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_NGtimeda, on=['地市'], how='left')  #拼接


#################################################
HM_Pcitytimeda=HMda[['地市','城镇普通品质装机时长']]
HM_Pcitytimeda['排名']=HMda['城镇普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_Pcitytimeda=pd.concat([HM_Pcitytimeda,HMZJ3]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_Pcitytimeda, on=['地市'], how='left')  #拼接

################################################
HM_PNCtimeda=HMda[['地市','农村普通品质装机时长']]
HM_PNCtimeda['排名']=HMda['农村普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_PNCtimeda=pd.concat([HM_PNCtimeda,HMZJ4]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_PNCtimeda, on=['地市'], how='left')  #拼接

####################################################
HM_citytimeda=HMda[['地市','城镇装机时长']]
HM_citytimeda['排名']=HMda['城镇装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_citytimeda=pd.concat([HM_citytimeda,HMZJ5]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_citytimeda, on=['地市'], how='left')  #拼接

#####################################################
HM_NCtimeda=HMda[['地市','农村装机时长']]
HM_NCtimeda['排名']=HMda['农村装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_NCtimeda=pd.concat([HM_NCtimeda,HMZJ6]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_NCtimeda, on=['地市'], how='left')  #拼接

#####################################################
HM_timeda=HMda[['地市','整体装机时长']]
HM_timeda['排名']=HMda['整体装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
HM_timeda=pd.concat([HM_timeda,HMZJ7]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_timeda, on=['地市'], how='left')  #拼接

###################################################
HM_GJSda=HMda[['地市','高品质装机及时率']]
HM_GJSda['排名']=HMda['高品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
HM_GJSda=pd.concat([HM_GJSda,HMZJ8]) #拼接
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_GJSda, on=['地市'], how='left')  #拼接

#################################################
HM_PJSda=HMda[['地市','普通品质装机及时率']]
HM_PJSda['排名']=HMda['普通品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
HM_PJSda=pd.concat([HM_PJSda,HMZJ9])
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_PJSda, on=['地市'], how='left')  #拼接

#################################################
HM_JSda=HMda[['地市','整体装机及时率']]
HM_JSda['排名']=HMda['整体装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
HM_JSda=pd.concat([HM_JSda,HMZJ10])
HM_CGtimeda=pd.merge(HM_CGtimeda, HM_JSda, on=['地市'], how='left')  #拼接

HM_CGtimeda.rename(columns={'排名_y': '排名'}, inplace=True)
HM_CGtimeda.rename(columns={'排名_x': '排名'}, inplace=True)
##########################################
PAda=pd.read_excel('筛选数据.xlsx',sheet_name='平安乡村数据')
PAZJ= PA.query('地市 == "总计"')
PAZJ=PAZJ.reset_index(drop=True)
PAZJ1=PAZJ[['城镇高品质装机时长']]
PAZJ2=PAZJ[['农村高品质装机时长']]
PAZJ3=PAZJ[['城镇普通品质装机时长']]
PAZJ4=PAZJ[['农村普通品质装机时长']]
PAZJ5=PAZJ[['城镇装机时长']]
PAZJ6=PAZJ[['农村装机时长']]
PAZJ7=PAZJ[['整体装机时长']]
PAZJ8=PAZJ[['高品质装机及时率']]
PAZJ9=PAZJ[['普通品质装机及时率']]
PAZJ10=PAZJ[['整体装机及时率']]
                #####数据整合完毕
######################################################
PA_CGtimeda=PAda[['地市','城镇高品质装机时长']]
PA_CGtimeda['排名']=PAda['城镇高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_CGtimeda=pd.concat([PA_CGtimeda,PAZJ1]) #拼接


######################################################
PA_NGtimeda=PAda[['地市','农村高品质装机时长']]
PA_NGtimeda['排名']=PAda['农村高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_NGtimeda=pd.concat([PA_NGtimeda,PAZJ2]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_NGtimeda, on=['地市'], how='left')  #拼接


#################################################
PA_Pcitytimeda=PAda[['地市','城镇普通品质装机时长']]
PA_Pcitytimeda['排名']=PAda['城镇普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_Pcitytimeda=pd.concat([PA_Pcitytimeda,PAZJ3]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_Pcitytimeda, on=['地市'], how='left')  #拼接

################################################
PA_PNCtimeda=PAda[['地市','农村普通品质装机时长']]
PA_PNCtimeda['排名']=PAda['农村普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_PNCtimeda=pd.concat([PA_PNCtimeda,PAZJ4]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_PNCtimeda, on=['地市'], how='left')  #拼接

####################################################
PA_citytimeda=PAda[['地市','城镇装机时长']]
PA_citytimeda['排名']=PAda['城镇装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_citytimeda=pd.concat([PA_citytimeda,PAZJ5]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_citytimeda, on=['地市'], how='left')  #拼接

#####################################################
PA_NCtimeda=PAda[['地市','农村装机时长']]
PA_NCtimeda['排名']=PAda['农村装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_NCtimeda=pd.concat([PA_NCtimeda,PAZJ6]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_NCtimeda, on=['地市'], how='left')  #拼接

#####################################################
PA_timeda=PAda[['地市','整体装机时长']]
PA_timeda['排名']=PAda['整体装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
PA_timeda=pd.concat([PA_timeda,PAZJ7]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_timeda, on=['地市'], how='left')  #拼接

###################################################
PA_GJSda=PAda[['地市','高品质装机及时率']]
PA_GJSda['排名']=PAda['高品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
PA_GJSda=pd.concat([PA_GJSda,PAZJ8]) #拼接
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_GJSda, on=['地市'], how='left')  #拼接

#################################################
PA_PJSda=PAda[['地市','普通品质装机及时率']]
PA_PJSda['排名']=PAda['普通品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
PA_PJSda=pd.concat([PA_PJSda,PAZJ9])
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_PJSda, on=['地市'], how='left')  #拼接

#################################################
PA_JSda=PAda[['地市','整体装机及时率']]
PA_JSda['排名']=PAda['整体装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
PA_JSda=pd.concat([PA_JSda,PAZJ10])
PA_CGtimeda=pd.merge(PA_CGtimeda, PA_JSda, on=['地市'], how='left')  #拼接

PA_CGtimeda.rename(columns={'排名_y': '排名'}, inplace=True)
PA_CGtimeda.rename(columns={'排名_x': '排名'}, inplace=True)
##########################################
ZWda=pd.read_excel('筛选数据.xlsx',sheet_name='智能组网数据')
ZWZJ= ZW.query('地市 == "总计"')
ZWZJ=ZWZJ.reset_index(drop=True)
ZWZJ1=ZWZJ[['城镇高品质装机时长']]
ZWZJ2=ZWZJ[['农村高品质装机时长']]
ZWZJ3=ZWZJ[['城镇普通品质装机时长']]
ZWZJ4=ZWZJ[['农村普通品质装机时长']]
ZWZJ5=ZWZJ[['城镇装机时长']]
ZWZJ6=ZWZJ[['农村装机时长']]
ZWZJ7=ZWZJ[['整体装机时长']]
ZWZJ8=ZWZJ[['高品质装机及时率']]
ZWZJ9=ZWZJ[['普通品质装机及时率']]
ZWZJ10=ZWZJ[['整体装机及时率']]
                #####数据整合完毕
######################################################
ZW_CGtimeda=ZWda[['地市','城镇高品质装机时长']]
ZW_CGtimeda['排名']=ZWda['城镇高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_CGtimeda=pd.concat([ZW_CGtimeda,ZWZJ1]) #拼接



######################################################
ZW_NGtimeda=ZWda[['地市','农村高品质装机时长']]
ZW_NGtimeda['排名']=ZWda['农村高品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_NGtimeda=pd.concat([ZW_NGtimeda,ZWZJ2]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_NGtimeda, on=['地市'], how='left')  #拼接


#################################################
ZW_Pcitytimeda=ZWda[['地市','城镇普通品质装机时长']]
ZW_Pcitytimeda['排名']=ZWda['城镇普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_Pcitytimeda=pd.concat([ZW_Pcitytimeda,ZWZJ3]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_Pcitytimeda, on=['地市'], how='left')  #拼接

################################################
ZW_PNCtimeda=ZWda[['地市','农村普通品质装机时长']]
ZW_PNCtimeda['排名']=ZWda['农村普通品质装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_PNCtimeda=pd.concat([ZW_PNCtimeda,ZWZJ4]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_PNCtimeda, on=['地市'], how='left')  #拼接

####################################################
ZW_citytimeda=ZWda[['地市','城镇装机时长']]
ZW_citytimeda['排名']=ZWda['城镇装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_citytimeda=pd.concat([ZW_citytimeda,ZWZJ5]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_citytimeda, on=['地市'], how='left')  #拼接

#####################################################
ZW_NCtimeda=ZWda[['地市','农村装机时长']]
ZW_NCtimeda['排名']=ZWda['农村装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_NCtimeda=pd.concat([ZW_NCtimeda,ZWZJ6]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_NCtimeda, on=['地市'], how='left')  #拼接

#####################################################
ZW_timeda=ZWda[['地市','整体装机时长']]
ZW_timeda['排名']=ZWda['整体装机时长'].rank(axis= 0,ascending=True,method='dense')   # 输出排名
ZW_timeda=pd.concat([ZW_timeda,ZWZJ7]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_timeda, on=['地市'], how='left')  #拼接

###################################################
ZW_GJSda=ZWda[['地市','高品质装机及时率']]
ZW_GJSda['排名']=ZWda['高品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
ZW_GJSda=pd.concat([ZW_GJSda,ZWZJ8]) #拼接
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_GJSda, on=['地市'], how='left')  #拼接

#################################################
ZW_PJSda=ZWda[['地市','普通品质装机及时率']]
ZW_PJSda['排名']=ZWda['普通品质装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
ZW_PJSda=pd.concat([ZW_PJSda,ZWZJ9])
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_PJSda, on=['地市'], how='left')  #拼接

#################################################
ZW_JSda=ZWda[['地市','整体装机及时率']]
ZW_JSda['排名']=ZWda['整体装机及时率'].rank(axis= 0,ascending=False,method='dense')   # 输出排名
ZW_JSda=pd.concat([ZW_JSda,ZWZJ10])
ZW_CGtimeda=pd.merge(ZW_CGtimeda, ZW_JSda, on=['地市'], how='left')  #拼接


ZW_CGtimeda.rename(columns={'排名_y': '排名'}, inplace=True)
ZW_CGtimeda.rename(columns={'排名_x': '排名'}, inplace=True)

##########################################
MBH_CGtimeda.loc[MBH_CGtimeda.shape[0]-1,'地市'] = ['全区']
IMS_CGtimeda.loc[IMS_CGtimeda.shape[0]-1,'地市'] = ['全区']
HM_CGtimeda.loc[HM_CGtimeda.shape[0]-1,'地市'] = ['全区']
PA_CGtimeda.loc[PA_CGtimeda.shape[0]-1,'地市'] = ['全区']
ZW_CGtimeda.loc[ZW_CGtimeda.shape[0]-1,'地市'] = ['全区']







with pd.ExcelWriter('魔百盒月报'+'.xlsx') as writer:
    MBH_CGtimeda.to_excel(writer, sheet_name='魔百和装机及时率', startcol=0, index=False, header=True)
    IMS_CGtimeda.to_excel(writer, sheet_name='IMS装机及时率', startcol=0, index=False, header=True)
    HM_CGtimeda.to_excel(writer, sheet_name='和目装机及时率', startcol=0, index=False, header=True)
    PA_CGtimeda.to_excel(writer, sheet_name='平安乡村装机及时率', startcol=0, index=False, header=True)
    ZW_CGtimeda.to_excel(writer, sheet_name='智能组网装机及时率', startcol=0, index=False, header=True)

