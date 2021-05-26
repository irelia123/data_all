import os, glob
import time
import pandas as pd
import numpy as np


Df = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','全区']},
    pd.Index(range(15)))

Gxjk = pd.read_excel('新家宽装机及时率报表.xlsx')
Qttd = pd.read_excel('（前）退单一览表.xlsx')
Httd = pd.read_excel('（后）退单一览表.xlsx')
Qtfl = pd.read_excel('退单分类.xlsx',sheet_name = '前台分类')
Htfl = pd.read_excel('退单分类.xlsx',sheet_name = '后台分类')
Qttd.BOSS撤单原因.fillna('空白',inplace=True) #  批量替换nan
Qttd.rename(columns={'BOSS撤单原因': '整体撤单原因'}, inplace=True)
Httd.rename(columns={'后端驳回原因': '整体撤单原因'}, inplace=True)
##################################################################
###这里比对前台宽带号码
Qttd1=pd.merge(Qttd,Gxjk [['宽带号码','序号']], on=['宽带号码'], how='left')
Qttd1.序号_y.fillna('保留',inplace=True) #  批量替换nan
Qttd1=Qttd1[Qttd1.序号_y=='保留']
Qttd_data=pd.merge(Qttd,Qttd1 [['宽带号码','序号_y']], on=['宽带号码'], how='left')
Qttd_data=Qttd_data[Qttd_data.序号_y !='保留']
Qttd_data = Qttd_data.drop(['后端驳回原因'],axis=1) ###删除不需要的数据，方便之后对齐
Qttd_data = Qttd_data.drop(['后端驳回备注'],axis=1)
Qttd_data = Qttd_data.drop(['BOSS撤单时间'],axis=1)
Qttd_data.rename(columns={'BOSS撤单原因': '整体撤单原因'}, inplace=True)

########################################################
###这里比对后台台宽带号码
Httd1=  pd.merge(Httd,Gxjk [['宽带号码','序号']], on=['宽带号码'], how='left')
Httd1.序号_y.fillna('保留',inplace=True) #  批量替换nan
Httd1 = Httd1[Httd1.序号_y=='保留']
Httd_data = pd.merge(Httd,Httd1 [['宽带号码','序号_y']], on=['宽带号码'], how='left')
Httd_data = Httd_data[Httd_data.序号_y !='保留']
Httd_data = Httd_data.drop(['BOSS撤单时间'],axis=1) ###删除不需要的数据，方便之后对齐
Httd_data = Httd_data.drop(['BOSS撤单原因'],axis=1)
Httd_data = Httd_data.drop(['后端驳回备注'],axis=1)
Httd_data.rename(columns={'后端驳回原因': '整体撤单原因'}, inplace=True)


####################################################
Zcl = pd.concat([Qttd_data,Httd_data], ignore_index=True)  #前后台退单重录拼接
Ztd = pd.concat([Qttd,Httd], ignore_index=True) #前后台退单拼接
Tdzs = Ztd.groupby(['地市']).size().reset_index(name='退单总数')
Tdzs=Tdzs.append([{'地市': '全区', '退单总数': Tdzs.apply(lambda x: x.sum()).退单总数}], ignore_index=True)


Qtcl = Qttd_data.groupby(['地市']).size().reset_index(name='前台退单重录')
Qtcl = Qtcl.append([{'地市': '全区', '前台退单重录': Qtcl.apply(lambda x: x.sum()).前台退单重录}], ignore_index=True)
Tdzs = pd.merge(Tdzs, Qtcl, on=['地市'], how='left')  #拼接
Htcl = Httd_data.groupby(['地市']).size().reset_index(name='后台退单重录')
Htcl = Htcl.append([{'地市': '全区', '后台退单重录': Htcl.apply(lambda x: x.sum()).后台退单重录}], ignore_index=True)
Tdzs = pd.merge(Tdzs, Htcl, on=['地市'], how='left')  #拼接
####################################
Tdcl = Zcl.groupby(['地市']).size().reset_index(name='撤单重录总数')
Tdcl = Tdcl.append([{'地市': '全区', '撤单重录总数': Tdcl.apply(lambda x: x.sum()).撤单重录总数}], ignore_index=True)
Tdzs = pd.merge(Tdzs,Tdcl, on=['地市'], how='left')
##########################################

Tdzs['撤单重录占比']=Tdzs.撤单重录总数/Tdzs.退单总数
Tdzs = Tdzs.fillna(0)  # 批量替换nan 为数字 0
Tdzs = Tdzs.round({'撤单重录占比': 4})  # 四舍五入
Tdzs['撤单重录占比']=Tdzs['撤单重录占比'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
####################################################
#####这里比对前台原因
QT = pd.merge(Qttd_data, Qtfl, on=['整体撤单原因'], how='left') #拼接
QT.前台退单原因.fillna('前台原因',inplace=True) #  批量替换nan
QT.前后台.fillna('前台',inplace=True) #  批量替换nan

QT_data = QT[QT.前台退单原因=='其他原因'].groupby(['地市']).size().reset_index(name='前台其他原因')
QT_data = QT_data.append([{'地市': '全区', '前台其他原因': QT_data.apply(lambda x: x.sum()).前台其他原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, QT_data, on=['地市'], how='left')  #拼接
##
QT_data1 = QT[QT.前台退单原因=='前台原因'].groupby(['地市']).size().reset_index(name='前台前台原因')
QT_data1 = QT_data1.append([{'地市': '全区', '前台前台原因': QT_data1.apply(lambda x: x.sum()).前台前台原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, QT_data1, on=['地市'], how='left')  #拼接
##
QT_data2 = QT[QT.前台退单原因=='用户原因'].groupby(['地市']).size().reset_index(name='前台用户原因')
QT_data2 = QT_data2.append([{'地市': '全区', '前台用户原因': QT_data2.apply(lambda x: x.sum()).前台用户原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, QT_data2, on=['地市'], how='left')  #拼接
####################################################
#####这里比对后台原因
HT = pd.merge(Httd_data, Htfl, on=['整体撤单原因'], how='left')
##
HT_data = HT[HT.后台退单原因=='前台原因'].groupby(['地市']).size().reset_index(name='后台前台原因')
HT_data = HT_data.append([{'地市': '全区', '后台前台原因': HT_data.apply(lambda x: x.sum()).后台前台原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, HT_data, on=['地市'], how='left')  #拼接

##
HT_data1 = HT[HT.后台退单原因=='其他原因'].groupby(['地市']).size().reset_index(name='后台其他原因')
HT_data1 = HT_data1.append([{'地市': '全区', '后台其他原因': HT_data1.apply(lambda x: x.sum()).后台其他原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, HT_data1, on=['地市'], how='left')  #拼接
##
HT_data2 = HT[HT.后台退单原因=='用户原因'].groupby(['地市']).size().reset_index(name='后台用户原因')
HT_data2 = HT_data2.append([{'地市': '全区', '后台用户原因': HT_data2.apply(lambda x: x.sum()).后台用户原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, HT_data2, on=['地市'], how='left')  #拼接
##
HT_data3 = HT[HT.后台退单原因=='建设原因'].groupby(['地市']).size().reset_index(name='后台建设原因')
HT_data3 = HT_data3.append([{'地市': '全区', '后台建设原因': HT_data3.apply(lambda x: x.sum()).后台建设原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, HT_data3, on=['地市'], how='left')  #拼接
##
HT_data4 = HT[HT.后台退单原因=='网络原因'].groupby(['地市']).size().reset_index(name='后台网络原因')
HT_data4 = HT_data4.append([{'地市': '全区', '后台网络原因': HT_data4.apply(lambda x: x.sum()).后台网络原因}], ignore_index=True)
Tdzs = pd.merge(Tdzs, HT_data4, on=['地市'], how='left')  #拼接
Tdzs = Tdzs.fillna(0)  # 批量替换nan 为数字 0
##
#计算前后台数据
Tdzs['前后台前台原因'] = Tdzs.前台前台原因+Tdzs.后台前台原因
Tdzs['前后台用户原因'] = Tdzs.前台用户原因+Tdzs.后台用户原因
Tdzs['前后台其他原因'] = Tdzs.前台其他原因+Tdzs.后台其他原因


Tdzs = Tdzs[['地市','前台退单重录','后台退单重录','撤单重录总数','退单总数','撤单重录占比','前台其他原因','前台前台原因','前台用户原因',
             '后台前台原因','后台用户原因','后台其他原因','后台建设原因','后台网络原因','前后台前台原因','前后台用户原因','前后台其他原因']]
Tdzs.rename(columns={'前台其他原因': '其他原因(前台)'}, inplace=True)
Tdzs.rename(columns={'前台前台原因': '前台原因(前台)'}, inplace=True)
Tdzs.rename(columns={'前台用户原因': '用户原因(前台)'}, inplace=True)
Tdzs.rename(columns={'后台前台原因': '前台原因(后台)'}, inplace=True)
Tdzs.rename(columns={'后台用户原因': '用户原因(后台)'}, inplace=True)
Tdzs.rename(columns={'后台其他原因': '其他原因(后台)'}, inplace=True)
Tdzs.rename(columns={'后台建设原因': '建设原因(后台)'}, inplace=True)
Tdzs.rename(columns={'后台网络原因': '网络原因(后台)'}, inplace=True)
Tdzs.rename(columns={'前后台前台原因': '前台原因(前后台)'}, inplace=True)
Tdzs.rename(columns={'前后台用户原因': '用户原因(前后台)'}, inplace=True)
Tdzs.rename(columns={'前后台其他原因': '其他原因(前后台)'}, inplace=True)
#######################################################
###第二页报表
#这里比对前台
Df_data1=pd.DataFrame({'前后台': ['前台','后台','总计']})##创建一个新的DataFrame
QT1 = pd.merge(Qttd, Qtfl, on=['整体撤单原因'], how='left') #拼接
QT1.前台退单原因.fillna('前台原因',inplace=True) #  批量替换nan
QT2 = pd.merge(Qttd_data, Qtfl, on=['整体撤单原因'], how='left') #拼接
QT_DATA = QT1.groupby(['前后台','整体撤单原因','前台退单原因']).size().reset_index(name='退单数')
Df2 = QT2.groupby(['前后台','整体撤单原因','前台退单原因']).size().reset_index(name='退单重录数')
QT_DATA = pd.merge(QT_DATA, Df2, on=['前后台','整体撤单原因','前台退单原因'], how='left') #拼接
QT_DATA = QT_DATA.fillna(0)  # 批量替换nan 为数字 0
#######################
#这里比对后台
HT1 = pd.merge(Httd, Htfl, on=['整体撤单原因'], how='left') #拼接
HT1.后台退单原因.fillna('后台退单原因',inplace=True) #  批量替换nan
HT2 = pd.merge(Httd_data, Htfl, on=['整体撤单原因'], how='left') #拼接
HT_DATA = HT1.groupby(['前后台','整体撤单原因','后台退单原因']).size().reset_index(name='退单数')
Df3 = HT2.groupby(['前后台','整体撤单原因','后台退单原因']).size().reset_index(name='退单重录数')
HT_DATA = pd.merge(HT_DATA, Df3, on=['前后台','整体撤单原因','后台退单原因'], how='left') #拼接
HT_DATA = HT_DATA.fillna(0)  # 批量替换nan 为数字 0
##################
#修改列名整合
QT_DATA.rename(columns={'前台退单原因': '退单原因'}, inplace=True)
HT_DATA.rename(columns={'后台退单原因': '退单原因'}, inplace=True)
#这里修改分类列名进行整合
Qtfl.rename(columns={'前台退单原因': '退单原因'}, inplace=True)
Htfl.rename(columns={'后台退单原因': '退单原因'}, inplace=True)
FL = pd.concat([Qtfl,Htfl], ignore_index=True) #拼接
#这里把前后台拼接后的数据进行分类
Zcl_data = pd.merge(Zcl, FL, on=['整体撤单原因'], how='left') #拼接
##
#计算重录总数
TDyuanyin = pd.concat([QT_DATA,HT_DATA], ignore_index=True) #拼接
DD = TDyuanyin[['退单重录数']]
DD = DD.div(DD.sum(axis=0),axis=1) #计算列占比
TDyuanyin = pd.merge(TDyuanyin,DD,left_index=True,right_index=True)

# 最后整改
TDyuanyin.rename(columns={'退单重录数_x': '退单重录数'}, inplace=True)
TDyuanyin.rename(columns={'退单重录数_y': '占比'}, inplace=True)
TDyuanyin = TDyuanyin.round({'占比': 4})  # 四舍五入
TDyuanyin['占比']=TDyuanyin['占比'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
TDyuanyin = TDyuanyin.append([{'前后台': '总计', '退单重录数': TDyuanyin.apply(lambda x: x.sum()).退单重录数}], ignore_index=True)


##
with pd.ExcelWriter('撤单对比报表'+'.xlsx') as writer:
    Tdzs.to_excel(writer, sheet_name='统计', startcol=0, index=False, header=True)
    TDyuanyin.to_excel(writer, sheet_name='重录工单二级明细', startcol=0, index=False, header=True)