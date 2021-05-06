import pandas as pd
import numpy as np
import time
start_time = time.time()


Dishi = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港','广西']},
    pd.Index(range(15)))
#导入需要数据
print('正在读取广西竣工详表...')
JG_data_Gx = pd.read_excel('4月1-10（广西）新家宽装机及时率报表.xlsx')
print('正在读取集团竣工详表...')
JG_data_Jt = pd.read_excel('新装千兆宽带竣工表(集团).xlsx')
print('正在读取回单质检详表...')
HD_data = pd.read_excel('0407回单质量检测指标地市报表详情.xls')
print('正在读取智能组网详表...')
ZW_data = pd.read_excel('0407智能组网分析报表详情.xls')
print('正在读取用户告知书分析详表...')
GZ_data_all = pd.read_excel('用户告知书分析报表.xlsx')
print('读取表格完毕!正在处理数据...')
##准备工作
HD_data_drop = HD_data[HD_data.工单主题.str.contains('FTTB') ==False]  #剔除B模式工单
JG_data_Jt = JG_data_Jt[JG_data_Jt.是否是校园宽带=='否'] #剔除校园工单
JG_data_Gx = JG_data_Gx[JG_data_Gx.工单标题.str.contains('1000M') ==True] #筛选千兆数据

GZ_data_all.fillna('空值',inplace=True) #  批量替换nan
#改列名
JG_data_Gx.rename(columns={'首次响应时长（H）': '首次响应时长'}, inplace=True)
JG_data_Gx.rename(columns={'工单历时（H）（剔除挂起时长）': '工单时长'}, inplace=True)
JG_drop= JG_data_Gx[JG_data_Gx.首次响应时长 >=0] #剔除首响时长为负数的数据
#城镇高品质
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '城镇')&(JG_data_Gx.用户等级 == '高品质')&(JG_data_Gx.工单时长 >=24 ),
               '是否超时'] = '超时'
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '城镇')&(JG_data_Gx.用户等级 == '高品质')&(JG_data_Gx.工单时长 < 24 ),
               '是否超时'] = '未超时'
#农村高品质
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '农村')&(JG_data_Gx.用户等级 == '高品质')&(JG_data_Gx.工单时长 >= 36 ),
               '是否超时'] = '超时'
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '农村')&(JG_data_Gx.用户等级 == '高品质')&(JG_data_Gx.工单时长 < 36 ),
               '是否超时'] = '未超时'
#城镇普通品质
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '城镇')&(JG_data_Gx.用户等级 == '普通')&(JG_data_Gx.工单时长 >= 48 ),
               '是否超时'] = '超时'
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '城镇')&(JG_data_Gx.用户等级 == '普通')&(JG_data_Gx.工单时长 < 48 ),
               '是否超时'] = '未超时'
#农村普通品质
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '农村')&(JG_data_Gx.用户等级 == '普通')&(JG_data_Gx.工单时长 >= 72 ),
               '是否超时'] = '超时'
JG_data_Gx.loc[(JG_data_Gx.区域类型 == '农村')&(JG_data_Gx.用户等级 == '普通')&(JG_data_Gx.工单时长 < 72 ),
               '是否超时'] = '未超时'
#计算装机竣工单量
Gx_jg = JG_data_Gx.groupby(['地市']).size().reset_index(name='装机竣工单量')
Gx_jg = Gx_jg.append([{'地市': '广西', '装机竣工单量': Gx_jg.apply(lambda x: x.sum()).装机竣工单量}], ignore_index=True)
#计算及时工单量
Gx_js = JG_data_Gx[JG_data_Gx.是否超时=='未超时'].groupby(['地市']).size().reset_index(name='及时工单量')
Gx_js = Gx_js.append([{'地市': '广西', '及时工单量': Gx_js.apply(lambda x: x.sum()).及时工单量}], ignore_index=True)
Gx_js = pd.merge(Gx_js, Gx_jg, on=['地市'], how='left') #拼接
#计算广西及时率
Gx_js['广西及时率'] = Gx_js.及时工单量/Gx_js.装机竣工单量
Gx_js = Gx_js.round({'广西及时率': 4})  # 四舍五入
Gx_js['广西及时率']=Gx_js['广西及时率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比


######
# 回单总数
QZ_DB = HD_data_drop.groupby(['地市']).size().reset_index(name='回单总数')
QZ_DB = QZ_DB.append([{'地市': '广西', '回单总数': QZ_DB.apply(lambda x: x.sum()).回单总数}], ignore_index=True)
QZ_DB = pd.merge(QZ_DB, Gx_js, on=['地市'], how='left') #拼接
#光功率达标数
QZ_DB1 = HD_data_drop[HD_data_drop.光功率=='达标'].groupby(['地市']).size().reset_index(name='光功率达标数')
QZ_DB1 = QZ_DB1.append([{'地市': '广西', '光功率达标数': QZ_DB1.apply(lambda x: x.sum()).光功率达标数}], ignore_index=True)
QZ_DB = pd.merge(QZ_DB, QZ_DB1, on=['地市'], how='left') #拼接
#光功率达标率
QZ_DB['光功率达标率'] = QZ_DB.光功率达标数/QZ_DB.回单总数
QZ_DB = QZ_DB.round({'光功率达标率': 4})  # 四舍五入
QZ_DB['光功率达标率']=QZ_DB['光功率达标率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
##
#城镇工单数
QZ_DB2 = HD_data_drop[HD_data_drop.区域类型=='城镇'].groupby(['地市']).size().reset_index(name='城镇总工单数')
QZ_DB2 = QZ_DB2.append([{'地市': '广西', '城镇总工单数': QZ_DB2.apply(lambda x: x.sum()).城镇总工单数}], ignore_index=True)
QZ_DB = pd.merge(QZ_DB, QZ_DB2, on=['地市'], how='left') #拼接
##
#城镇工单最终达标数
QZ_DB3 = HD_data_drop[(HD_data_drop.最终测速结果=='达标')&(HD_data_drop.区域类型=='城镇')].groupby(['地市']).size().reset_index(name='城镇最终达标数')
QZ_DB3 = QZ_DB3.append([{'地市': '广西', '城镇最终达标数': QZ_DB3.apply(lambda x: x.sum()).城镇最终达标数}], ignore_index=True)
QZ_DB = pd.merge(QZ_DB, QZ_DB3, on=['地市'], how='left') #拼接
QZ_DB['软探针测速达标率'] = QZ_DB.城镇最终达标数/QZ_DB.城镇总工单数
QZ_DB = QZ_DB.round({'软探针测速达标率': 4})  # 四舍五入
QZ_DB['软探针测速达标率']=QZ_DB['软探针测速达标率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
##
#首响平均时长
QZ_time = JG_drop[['地市', '首次响应时长']].groupby(['地市']).mean().reset_index()
QZ_time.rename(columns={'首次响应时长': '首响平均时长'}, inplace=True)
QZ_time = QZ_time.append(
    [{'地市': '广西', '首响平均时长': JG_drop[['地市', '首次响应时长']].mean().reset_index(name='首响平均时长').at[0, '首响平均时长']}],
    ignore_index=True)  # 计算全区
QZ_time= QZ_time = QZ_time.round({'首响平均时长': 2})  # 四舍五入

#
#装机平均工作时
QZ_time1 = JG_data_Gx[['地市', '工单时长']].groupby(['地市']).mean().reset_index()
QZ_time1.rename(columns={'工单时长': '装机平均工作时'}, inplace=True)
QZ_time1 = QZ_time1.append(
    [{'地市': '广西', '装机平均工作时': JG_data_Gx[['地市', '工单时长']].mean().reset_index(name='装机平均工作时').at[0, '装机平均工作时']}],
    ignore_index=True)  # 计算全区
QZ_time1= QZ_time1 = QZ_time1.round({'装机平均工作时': 2})  # 四舍五入
QZ_time = pd.merge(QZ_time, QZ_time1, on=['地市'], how='left') #拼接
Data_all = pd.merge(QZ_DB, QZ_time, on=['地市'], how='left')  #拼接
####################################################
GZ_data_all['智能组网是否合格1'] = ''
GZ_data_all['接入光功率1'] = ''
GZ_data_all['光猫出口数据'] = '是'
GZ_data_all['用户设备检查'] = '是'
GZ_data_all['机顶盒是否有线接入'] = ''
GZ_data_all['是否检测全屋WIFI信号'] = '是'
GZ_data_all['是否签订告知书'] = '是'
GZ_data_all['是否完成三必做'] = '是'
GZ_data_all['十步法是否执行到位'] = '合格'
##判断智能组网是否合格，装后评分>装前评分，如装后评分=100 自动减免
absolve = 100
GZ_data_all.loc[(GZ_data_all['装后得分'] > GZ_data_all['装前得分']), '智能组网是否合格1'] = '是'
GZ_data_all.loc[(GZ_data_all['装后得分'] <= GZ_data_all['装前得分']), '智能组网是否合格1'] = '否'
GZ_data_all.loc[(GZ_data_all['装后得分'] == absolve ), '智能组网是否合格1'] = '是'
GZ_data_all.loc[(GZ_data_all['智能组网宽带账号'] == '无组网工单' ), '智能组网是否合格1'] = '是'
##判断接入光功率是否合格
GZ_data_all.loc[(GZ_data_all['光功率是否达标'] == '达标'), '接入光功率1'] = '是'
GZ_data_all.loc[(GZ_data_all['光功率是否达标'] == '不达标'), '接入光功率1'] = '否'
##判断用户设备检查是否合格
GZ_data_all.loc[(GZ_data_all['路由器'] == '空值'), '用户设备检查'] = '否'
GZ_data_all.loc[(GZ_data_all['电脑网卡'] == '空值'), '用户设备检查'] = '否'
GZ_data_all.loc[(GZ_data_all['室内网线'] == '空值'), '用户设备检查'] = '否'
GZ_data_all.loc[(GZ_data_all['手机设备'] == '空值'), '用户设备检查'] = '否'
##判断机顶盒是否有线接入
GZ_data_all.loc[(GZ_data_all['机顶盒'] == '有线连接'), '机顶盒是否有线接入'] = '是'
GZ_data_all.loc[(GZ_data_all['机顶盒'] == '无线连接'), '机顶盒是否有线接入'] = '否'
##判断是否检测全屋WIFI信号
GZ_data_all.loc[(GZ_data_all['全屋WIFI信号'] == '空值'), '是否检测全屋WIFI信号'] = '否'
##判断是否签订告知书
GZ_data_all.loc[(GZ_data_all['是否具备千兆条件'] == '空值'), '是否签订告知书'] = '否'
##判断是否三必做
GZ_data_all.loc[(GZ_data_all['是否签订告知书'] == '否'), '是否完成三必做'] = '否'
GZ_data_all.loc[(GZ_data_all['是否留下光猫贴'] == '否'), '是否完成三必做'] = '否'
GZ_data_all.loc[(GZ_data_all['是否演示千兆业务'] == '否'), '是否完成三必做'] = '否'
##判断十步法是否执行到位
GZ_data_all.loc[(GZ_data_all['光猫出口数据'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['用户设备检查'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['机顶盒是否有线接入'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['是否检测全屋WIFI信号'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['是否签订告知书'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['智能组网是否合格1'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['接入光功率1'] == '否'), '十步法是否执行到位'] = '不合格'
GZ_data_all.loc[(GZ_data_all['是否完成三必做'] == '否'), '十步法是否执行到位'] = '不合格'
##开始计算质检单数
GZ_data_all_1000=GZ_data_all[(GZ_data_all.用户签约宽带=='1000M')]  #筛选千兆数据
GZ_ten = GZ_data_all_1000.groupby(['地市']).size().reset_index(name='质检单数')
GZ_ten = GZ_ten.append([{'地市': '广西', '质检单数': GZ_ten.apply(lambda x: x.sum()).质检单数}], ignore_index=True)
#开始计算合格单数
GZ_ten1 = GZ_data_all_1000[GZ_data_all_1000.十步法是否执行到位=='合格'].groupby(['地市']).size().reset_index(name='合格单数')
GZ_ten1 = GZ_ten1.append([{'地市': '广西', '合格单数': GZ_ten1.apply(lambda x: x.sum()).合格单数}], ignore_index=True)
GZ_ten = pd.merge(GZ_ten, GZ_ten1, on=['地市'], how='left') #拼接
GZ_ten['十步法合格率'] = GZ_ten.合格单数/GZ_ten.质检单数
GZ_ten = GZ_ten.round({'十步法合格率': 4})  # 四舍五入
GZ_ten['十步法合格率']=GZ_ten['十步法合格率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比
Data_all = pd.merge(Data_all, GZ_ten, on=['地市'], how='left')  #拼接

#排序
GZ_data_all=GZ_data_all[['十步法是否执行到位','智能组网是否合格1','接入光功率1','光猫出口数据','用户设备检查','机顶盒是否有线接入','是否检测全屋WIFI信号','是否完成三必做','是否签订告知书','是否留下光猫贴',
                         '是否演示千兆业务','地市','工单标题','工单类型','boss派单时间','boss归档完成时间','区域分类','区域类型','工单编码','是否千兆宽带','用户签约宽带','宽带号码','智能组网是否合格',
                         '接入光功率','光功率是否达标','光猫千兆口连接','光猫出口数据','出口速率是否达标','路由器','电脑网卡','室内网线','机顶盒','全屋WIFI信号','手机设备','是否留下光猫贴','是否演示千兆业务',
                             '是否具备千兆条件','是否三必做','十步法是否执行到位','是否通过RMS平台注册','智能组网宽带账号','装前得分','装后得分']]
GZ_data_all.rename(columns={'智能组网是否合格1': '智能组网是否合格（组网装后测评得分大于装前测评得分）'}, inplace=True)
GZ_data_all.rename(columns={'光猫出口数据': '光猫出口数据（测速）此项数据暂时忽略'}, inplace=True)
GZ_data_all.rename(columns={'用户设备检查': '用户设备检查（路由器、网线、网卡、手机）'}, inplace=True)
GZ_data_all.rename(columns={'接入光功率1': '接入光功率是否达标'}, inplace=True)
GZ_data_all.rename(columns={'是否完成三必做': '是否完成三必做（签订告知书、演示千兆业务、留下光猫贴）'}, inplace=True)
GZ_data_all.rename(columns={'网络千兆口连接': '光猫千兆口连接'}, inplace=True)
GZ_data_all.replace('空值','',inplace=True) #替换数据
#####################
#城镇高品质
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '城镇')&(JG_data_Jt.用户等级 == '高品质')&(JG_data_Jt.工单历时减去总时长分化时 >=24 ),
               '是否超时'] = '超时'
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '城镇')&(JG_data_Jt.用户等级 == '高品质')&(JG_data_Jt.工单历时减去总时长分化时 < 24 ),
               '是否超时'] = '未超时'
#农村高品质
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '农村')&(JG_data_Jt.用户等级 == '高品质')&(JG_data_Jt.工单历时减去总时长分化时 >= 36 ),
               '是否超时'] = '超时'
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '农村')&(JG_data_Jt.用户等级 == '高品质')&(JG_data_Jt.工单历时减去总时长分化时 < 36 ),
               '是否超时'] = '未超时'
#城镇普通品质
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '城镇')&(JG_data_Jt.用户等级 == '普通')&(JG_data_Jt.工单历时减去总时长分化时 >= 48 ),
               '是否超时'] = '超时'
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '城镇')&(JG_data_Jt.用户等级 == '普通')&(JG_data_Jt.工单历时减去总时长分化时 < 48 ),
               '是否超时'] = '未超时'
#农村普通品质
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '农村')&(JG_data_Jt.用户等级 == '普通')&(JG_data_Jt.工单历时减去总时长分化时 >= 72 ),
               '是否超时'] = '超时'
JG_data_Jt.loc[(JG_data_Jt.区域类型 == '农村')&(JG_data_Jt.用户等级 == '普通')&(JG_data_Jt.工单历时减去总时长分化时 < 72 ),
               '是否超时'] = '未超时'
#计算集团竣工数
Jk_js = JG_data_Jt.groupby(['地市']).size().reset_index(name='竣工数')
Jk_js = Jk_js.append([{'地市': '广西', '竣工数': Jk_js.apply(lambda x: x.sum()).竣工数}], ignore_index=True)
#计算集团及时工单量
Jk_js1 = JG_data_Jt[JG_data_Jt.是否超时=='未超时'].groupby(['地市']).size().reset_index(name='及时工单量')
Jk_js1 = Jk_js1.append([{'地市': '广西', '及时工单量': Jk_js1.apply(lambda x: x.sum()).及时工单量}], ignore_index=True)
Jk_js = pd.merge(Jk_js, Jk_js1, on=['地市'], how='left')  #拼接
#计算及时率
Jk_js['集团及时率'] = Jk_js.及时工单量/Jk_js.竣工数
Jk_js = Jk_js.round({'集团及时率': 4})  # 四舍五入
Jk_js['集团及时率']=Jk_js['集团及时率'].apply(lambda x: '%.2f%%' % (x * 100))  #转换百分比


Data_all.drop(['回单总数'],axis=1,inplace=True)
Data_all.drop(['及时工单量'],axis=1,inplace=True)
Data_all.drop(['装机竣工单量'],axis=1,inplace=True)
Data_all.drop(['光功率达标数'],axis=1,inplace=True)
Data_all.drop(['城镇总工单数'],axis=1,inplace=True)
Data_all.drop(['城镇最终达标数'],axis=1,inplace=True)
Data_all.drop(['质检单数'],axis=1,inplace=True)
Data_all.drop(['合格单数'],axis=1,inplace=True)
print('数据处理完毕,正在导出....')



with pd.ExcelWriter('盯控指标_千兆'+'.xlsx') as writer:
    Data_all.to_excel(writer, sheet_name='千兆指标', startcol=0, index=False, header=True)
    JG_data_Gx.to_excel(writer, sheet_name='千兆竣工详表(广西标准)', startcol=0, index=False, header=True)
    HD_data_drop.to_excel(writer, sheet_name='千兆质量检测', startcol=0, index=False, header=True)
    GZ_data_all.to_excel(writer, sheet_name='千兆告知书详表', startcol=0, index=False, header=True)
    ZW_data.to_excel(writer, sheet_name='智能组网清单', startcol=0, index=False, header=True)
    Jk_js.to_excel(writer, sheet_name='家宽及时率', startcol=0, index=False, header=True)
    JG_data_Jt.to_excel(writer, sheet_name='千兆竣工详表(集团模板)', startcol=0, index=False, header=True)
end_time = time.time()

print('处理完毕!!!总耗时%0.1f秒钟'%(start_time-end_time))
