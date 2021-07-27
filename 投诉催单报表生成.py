import pandas as pd
import datetime
import time

start_time = time.time()

today_now = datetime.datetime.now()  # 系统取当天日期
today = datetime.datetime.strptime('2021-7-3-15-00-00', '%Y-%m-%d-%H-%M-%S')  # 手动输入日期
# EXCEL路径

dishi = '(钦州)'
DISHI = '钦州'

COMPLAINT_PATH = '家庭业务投诉工单查询导出' + dishi + '.xlsx'

MATCH_PATH = '投诉工单监控列表' + dishi + '.xlsx'

AVERAGE_DAILY = '日均归档量' + dishi + '.xlsx'

GRID_PATH = '网格长名单' + dishi + '.xlsx'

LIST_PATH = '网格人员名单' + dishi + '.xlsx'

EX_PATH = '异常工单表.xlsx'

ex = pd.read_excel(EX_PATH)

list_data = pd.read_excel(LIST_PATH)
list_data = list_data.drop_duplicates(subset=['施工人员'], keep='first')


# 判断函数
def judge_table(obj, mat, key):
    map1 = obj[key].map(lambda x: x in list(mat[key]))  # 两个表之间一列进行判断
    map1 = pd.DataFrame(map1)
    map1.rename(columns={key: '判断'}, inplace=True)  # 改列名
    obj = pd.concat([obj, map1], axis=1)
    obj = obj[obj.判断 == False]
    obj.drop(['判断'], axis=1, inplace=True)
    return obj


# 计算总数函数
def data_total(data, cols, name, off=1):
    if off == 1:
        number = data.groupby([cols]).size().reset_index()
        number.rename(columns={0: '总数'}, inplace=True)
        number = number.append([{cols: '合计', '总数': number.apply(lambda x: x.sum()).总数}], ignore_index=True)
        number.rename(columns={'总数': name}, inplace=True)
    else:
        number = data.groupby([cols]).size().reset_index()
        number.rename(columns={0: name}, inplace=True)
    return number


# 批量修改表格格式为str
def format_str(data, name):
    data_types_dict = {name: str}
    data = data.astype(data_types_dict)
    return data


def make_excel(data, col, off=0):
    # 今日在途
    city = data_total(complaint, col, '今日在途', off=off)
    data = data.append([{col: '合计', '日均归档量': data.apply(lambda x: x.sum()).日均归档量}], ignore_index=True)
    city = pd.merge(data, city, on=[col], how='left')
    # 压单比
    city['压单比'] = city.今日在途 / city.日均归档量
    city = city.round({'压单比': 2})
    city.drop(['日均归档量'], axis=1, inplace=True)
    # 千兆在途总量
    city = pd.merge(city, data_total(complaint[complaint.千兆工单 == '是'], col, '千兆在途总量', off=off), on=[col], how='left')
    # 计算千兆超长超时工单
    overtime = complaint[complaint.是否要剔除 == 'FALSE']

    city = pd.merge(city, data_total(overtime[(overtime.千兆工单 == '是') & (overtime.超时未竣工量 == '超24-48小时未竣工')], col,
                                     '千兆超24-48小时未竣工', off=off), on=[col], how='left')
    city = pd.merge(city, data_total(overtime[(overtime.千兆工单 == '是') & (overtime.超时未竣工量 == '超48-100小时未竣工')], col,
                                     '千兆超48小时未竣工', off=off), on=[col], how='left')
    city = pd.merge(city, data_total(overtime[(overtime.千兆工单 == '是') & (overtime.超时未竣工量 == '超100小时未竣工')], col,
                                     '千兆超100小时未竣工', off=off), on=[col], how='left')
    # 计算超长超时工单
    city = pd.merge(city, data_total(overtime[(overtime[today] > 48) & (overtime.千兆工单 == '否')], col, '超48小时', off=off),
                    on=[col], how='left')
    city = pd.merge(city,
                    data_total(overtime[(overtime[today] > 168) & (overtime.千兆工单 == '否')], col, '超7天', off=off),
                    on=[col], how='left')
    city = pd.merge(city,
                    data_total(overtime[(overtime[today] > 360) & (overtime.千兆工单 == '否')], col, '超15天', off=off),
                    on=[col], how='left')
    city = pd.merge(city,
                    data_total(overtime[(overtime[today] > 720) & (overtime.千兆工单 == '否')], col, '超30天', off=off),
                    on=[col], how='left')
    city = city.fillna(0)
    # 超48小时占比(千兆+普通)  ＜15%
    city['超48小时占比(千兆+普通)＜15%'] = (city.千兆超48小时未竣工 + city.千兆超100小时未竣工 + city.超48小时) / city.今日在途
    city = city.fillna(0)
    city['超48小时占比(千兆+普通)＜15%'] = city['超48小时占比(千兆+普通)＜15%'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
    # 重复投诉单量
    city = pd.merge(city, data_total(complaint[complaint.重复投诉次数 >= 1], col, '重复投诉单数量', off=off), on=[col], how='left')
    city = city.fillna(0)
    # 重复投诉率
    city['重复投诉单率'] = city.重复投诉单数量 / city.今日在途
    city = city.fillna(0)
    city['重复投诉单率'] = city['重复投诉单率'].apply(lambda x: '%.2f%%' % (x * 100))  # 转换百分比
    return city


# 在途表
complaint = pd.read_excel(COMPLAINT_PATH)
complaint['最后质检通过时间'].fillna('空', inplace=True)
complaint = format_str(complaint, '来电号码')
complaint = format_str(complaint, '宽带账号')

# 匹配表
match = pd.read_excel(MATCH_PATH, skiprows=1)

complaint.loc[(complaint['责任区县'] == '钦南区'), '责任区县'] = '城区'
complaint.loc[(complaint['责任区县'] == '钦北区'), '责任区县'] = '城区'
match.loc[(match['所属区县'] == '钦南区'), '所属区县'] = '城区'
match.loc[(match['所属区县'] == '钦北区'), '所属区县'] = '城区'

complaint = complaint[(complaint.最后质检通过时间 == '空') & (complaint.工单类型 != '和家亲')
                      & (complaint.工单类型 != '问题解决单') & (complaint.责任地市 == DISHI)]  # 筛选数据

complaint = pd.merge(complaint, match[['工单流水号', '施工人员']], on='工单流水号', how='left')

match = pd.merge(match, complaint[['工单流水号']], on='工单流水号')
match = judge_table(match, ex, '工单流水号')
complaint_data = pd.DataFrame(
    complaint[['首次到单时间', '广西时限', '集团时限', '投诉分类', '工单状态', '工单流水号', '工单类型', '速率']])
# 日均归档表
r_city = pd.read_excel(AVERAGE_DAILY, sheet_name=0)  # 地市

r_county = pd.read_excel(AVERAGE_DAILY, sheet_name=1)  # 区县

r_grid = pd.read_excel(AVERAGE_DAILY, sheet_name=2)  # 网格

# 计算工作时
complaint_data['广西剩余工作时'] = (pd.to_datetime(complaint_data.广西时限) - today_now).dt.days * 24 + (
        pd.to_datetime(complaint_data.广西时限) - today_now).dt.seconds / 3600
complaint_data[today_now] = (today_now - pd.to_datetime(complaint_data.首次到单时间)).dt.days * 24 + (
        today_now - pd.to_datetime(complaint_data.首次到单时间)).dt.seconds / 3600
complaint_data[today] = (today - pd.to_datetime(complaint_data.首次到单时间)).dt.days * 24 + (
        today - pd.to_datetime(complaint_data.首次到单时间)).dt.seconds / 3600
# 计算剔除夜间时间工作时间
complaint_data['集团时限'] = pd.to_datetime(complaint_data['集团时限'])
complaint_data['首次到单时间'] = pd.to_datetime(complaint_data['首次到单时间'])
# 现在时
complaint_data['现在时'] = today_now
# 现在时早八点
complaint_data['现在时早八点'] = complaint_data['现在时'].dt.strftime('%Y-%m-%d')
complaint_data['现在时早八点'] = complaint_data.现在时早八点 + ' 08:00:00'
complaint_data['现在时早八点'] = pd.to_datetime(complaint_data['现在时早八点'])
# 现在时晚二十点
complaint_data['现在时晚二十点'] = complaint_data['现在时'].dt.strftime('%Y-%m-%d')
complaint_data['现在时晚二十点'] = complaint_data.现在时晚二十点 + ' 20:00:00'
complaint_data['现在时晚二十点'] = pd.to_datetime(complaint_data['现在时晚二十点'])

# 到单时间的当日20点
complaint_data['到单二十点'] = complaint_data['首次到单时间'].dt.strftime('%Y-%m-%d')
complaint_data['到单二十点'] = complaint_data.到单二十点 + ' 20:00:00'
complaint_data['到单二十点'] = pd.to_datetime(complaint_data['到单二十点'])
# 到单时间的当日8点
complaint_data['到单早八点'] = complaint_data['首次到单时间'].dt.strftime('%Y-%m-%d')
complaint_data['到单早八点'] = complaint_data.到单早八点 + ' 08:00:00'
complaint_data['到单早八点'] = pd.to_datetime(complaint_data['到单早八点'])
# 到单第二日8点
complaint_data['到单时间'] = pd.to_datetime(complaint_data['首次到单时间'], format='%Y-%m-%d')
complaint_data['到单时间'] = (pd.to_datetime(complaint_data['首次到单时间']) + datetime.timedelta(days=1))
complaint_data['到单时间'] = complaint_data['到单时间'].dt.strftime('%Y-%m-%d')
complaint_data['到单时间'] = complaint_data.到单时间 + ' 08:00:00'
complaint_data['到单时间'] = pd.to_datetime(complaint_data['到单时间'])
# 集团八点
complaint_data['八点集团时限'] = pd.to_datetime(complaint_data['集团时限'], format='%Y-%m-%d')
complaint_data['八点集团时限'] = complaint_data['八点集团时限'].dt.strftime('%Y-%m-%d')
complaint_data['八点集团时限'] = complaint_data.八点集团时限 + ' 20:00:00'
complaint_data['集团早八点'] = complaint_data.八点集团时限 + ' 08:00:00'
complaint_data['集团早八点'] = pd.to_datetime(complaint_data['集团早八点'])
complaint_data['八点集团时限'] = pd.to_datetime(complaint_data['八点集团时限'])
# 到单修改时间
complaint_data['到单修改时间'] = complaint_data['首次到单时间']
complaint_data.loc[(complaint_data['首次到单时间'] > complaint_data['到单二十点']), '到单修改时间'] = complaint_data['到单时间']
complaint_data.loc[(complaint_data['到单修改时间'] < complaint_data['到单早八点']), '到单修改时间'] = complaint_data['到单早八点']
# 集团修改时间
complaint_data['集团修改时间'] = complaint_data['集团时限']
complaint_data.loc[(complaint_data['集团时限'] > complaint_data['八点集团时限']), '集团修改时间'] = complaint_data['八点集团时限']
complaint_data.loc[(complaint_data['集团修改时间'] < complaint_data['集团早八点']), '集团修改时间'] = complaint_data['集团早八点']
# 现在时修改时间
complaint_data['现在时修改时间'] = complaint_data['现在时']
complaint_data.loc[(complaint_data['现在时'] > complaint_data['现在时晚二十点']), '现在时修改时间'] = complaint_data['现在时晚二十点']
complaint_data.loc[(complaint_data['现在时修改时间'] < complaint_data['现在时早八点']), '现在时修改时间'] = complaint_data['现在时早八点']
# 现在日期(计算天数用)
complaint_data['当天日期'] = pd.to_datetime(complaint_data['现在时修改时间'], format='%Y-%m-%d')
complaint_data['当天日期'] = complaint_data['当天日期'].dt.strftime('%Y-%m-%d')
complaint_data['当天日期'] = pd.to_datetime(complaint_data['当天日期'])

# 集团日期(计算天数用)
complaint_data['集团日期'] = pd.to_datetime(complaint_data['集团修改时间'], format='%Y-%m-%d')
complaint_data['集团日期'] = complaint_data['集团日期'].dt.strftime('%Y-%m-%d')
complaint_data['集团日期'] = pd.to_datetime(complaint_data['集团日期'])
# 到单日期(计算天数用)
complaint_data['到单日期'] = pd.to_datetime(complaint_data['到单修改时间'], format='%Y-%m-%d')
complaint_data['到单日期'] = complaint_data['到单日期'].dt.strftime('%Y-%m-%d')
complaint_data['到单日期'] = pd.to_datetime(complaint_data['到单日期'])
# 计算天数
complaint_data['天数'] = (pd.to_datetime(complaint_data['集团日期']) - pd.to_datetime(complaint_data['到单日期'])).dt.days
complaint_data['当前天数'] = (pd.to_datetime(complaint_data['当天日期']) - pd.to_datetime(complaint_data['到单日期'])).dt.days
# 计算集团剩余时间
complaint_data['集团剩余时间'] = (pd.to_datetime(complaint_data['集团修改时间']) - pd.to_datetime(complaint_data['到单修改时间'])).astype(
    'timedelta64[s]') / 3600
complaint_data['集团剩余时间'] = complaint_data.集团剩余时间 - (complaint_data.天数 * 12)
# 计算工单历时
complaint_data['工单历时'] = (pd.to_datetime(complaint_data['现在时修改时间']) - pd.to_datetime(complaint_data['到单修改时间'])).astype(
    'timedelta64[s]') / 3600
complaint_data['工单历时'] = complaint_data.工单历时 - (complaint_data.当前天数 * 12)
# 集团剩余工作时间
complaint_data['集团剩余工作时'] = complaint_data.集团剩余时间 - complaint_data.工单历时

# 删列
complaint_data = complaint_data.drop(
    ['现在时', '现在时早八点', '现在时晚二十点', '到单二十点', '到单早八点', '到单时间', '八点集团时限', '集团早八点',
     '到单修改时间', '集团修改时间', '现在时修改时间', '当天日期', '集团日期', '到单日期', '天数', '当前天数'],
    axis=1)

complaint_data = pd.merge(complaint_data, match[['施工人员', '工单流水号']], on='工单流水号', how='left')
complaint_data = pd.merge(complaint_data, list_data, on='施工人员', how='left')
complaint_data.rename(columns={'责任区县': '区县', '网格名称': '网格'}, inplace=True)

# 客户姓名
complaint_data[
    '客户名称'] = complaint.姓名 + '【宽带账号' + complaint.宽带账号 + ',来电号码' + complaint.来电号码 + '】,' + complaint.安装地址 + ',' + complaint.工单流水号 + '。注意:集团时限' + complaint.集团时限 + ',广西时限' + complaint.广西时限 + ',请按先到时限优先处理，谢谢!'

complaint_data['是否要剔除'] = 'FALSE'

complaint_data['千兆工单'] = '否'

complaint_data['是否异常'] = ''

complaint_data.loc[(complaint_data['工单类型'] == '待归档'), '是否要剔除'] = 'TRUE'

complaint_data.loc[(complaint_data['速率'] == '1000M'), '千兆工单'] = '是'

complaint_data.loc[complaint_data[today] > 24, '超时未竣工量'] = '超24-48小时未竣工'
complaint_data.loc[complaint_data[today] > 48, '超时未竣工量'] = '超48-100小时未竣工'
complaint_data.loc[complaint_data[today] > 100, '超时未竣工量'] = '超100小时未竣工'
# 修改日期格式为str
complaint_data = format_str(complaint_data, '广西时限')
complaint_data = format_str(complaint_data, '集团时限')
complaint_data = format_str(complaint_data, '首次到单时间')
# 排序

complaint_data = pd.DataFrame(complaint_data[
                                  ['集团剩余时间', '工单历时', '广西剩余工作时', '集团剩余工作时', today_now, today, '区县', '网格', '施工人员', '客户名称',
                                   '首次到单时间', '广西时限', '集团时限', '投诉分类', '工单状态', '是否要剔除', '千兆工单', '是否异常', '超时未竣工量']])

grid = pd.read_excel(GRID_PATH)

complaint = pd.merge(complaint_data, complaint)

grid = pd.merge(grid, make_excel(r_grid, '网格', 1), on='网格', how='right')
# =====================================================================================
# 1-超时未竣工工单通报
# 千兆超时 >48
QZ_Overtime = complaint[(complaint.千兆工单 == '是') & (complaint.是否要剔除 == 'FALSE') & (complaint[today] > 24)]
QZ_Overtime = QZ_Overtime[['区县', '网格', '施工人员', '工单流水号', '姓名', '客户电话', '联系电话1', '宽带账号', '来电号码', '安装地址', today]]
QZ_Overtime.rename(columns={today: '历时小时'}, inplace=True)  # 改列名
# 普通超时 48-100
PT_Overtime = complaint[
    (complaint.千兆工单 == '否') & (complaint.是否要剔除 == 'FALSE') & (complaint[today] > 48)]
PT_Overtime = PT_Overtime[['区县', '网格', '施工人员', '工单流水号', '姓名', '客户电话', '联系电话1', '宽带账号', '来电号码', '安装地址', today]]
PT_Overtime.rename(columns={today: '历时小时'}, inplace=True)  # 改列名

# ======================================================================================
# 2-即将超时工单通报
# 千兆即将超时 >-24,<3
QZ_Timeout = complaint[
    (complaint.千兆工单 == '是') & (complaint.是否要剔除 == 'FALSE') & (complaint.集团剩余工作时 >= -24)]
QZ_Timeout_data = QZ_Timeout[['区县', '网格', '施工人员', '客户名称', '广西剩余工作时', '集团剩余工作时']]
# 普通即将超时 >-24,<3
PT_Timeout = complaint[
    (complaint.千兆工单 == '否') & (complaint.是否要剔除 == 'FALSE') & (complaint.集团剩余工作时 >= -24) & (complaint.集团剩余工作时 <= 3)]
PT_Timeout_data = PT_Timeout[['区县', '网格', '施工人员', '客户名称', '广西剩余工作时', '集团剩余工作时']]

Wangge = pd.DataFrame(r_grid.iloc[0:, 0:1])
# 千兆总计
total = pd.merge(Wangge, data_total(QZ_Timeout, '网格', '千兆即将超时工单'), on='网格', how='left')
total = total.append([{'网格': '合计', '千兆即将超时工单': total.apply(lambda x: x.sum()).千兆即将超时工单}], ignore_index=True)
# 普通总计
total = pd.merge(total, data_total(PT_Timeout, '网格', '普通即将超时工单'), on='网格', how='left')
# 合计
total = total.fillna(0)
total['合计'] = total.千兆即将超时工单 + total.普通即将超时工单

title1 = pd.DataFrame({'千兆超时未竣工在途工单': ['']})
title2 = pd.DataFrame({'普通超48-100小时未竣工在途工单': ['']})
title3 = pd.DataFrame({'千兆即将超时工单': ['']})
title4 = pd.DataFrame({'普通宽带即将超时工单': ['']})

complaint = judge_table(complaint, ex, '工单流水号')
with pd.ExcelWriter('投诉催单报表' + dishi + '.xlsx') as writer:  # 写入结果为当前路径
    complaint.to_excel(writer, sheet_name='在途单', startcol=0, index=False, header=True)
    match.to_excel(writer, sheet_name='匹配表', startcol=0, index=False, header=True)
    make_excel(r_city, '责任地市', 0).to_excel(writer, sheet_name='地市与区县', startcol=0, index=False, header=True)
    make_excel(r_county, '区县', 1).to_excel(writer, sheet_name='地市与区县', startcol=0, startrow=2, index=False,
                                           header=True)
    grid.to_excel(writer, sheet_name='网格', startcol=0, index=False, header=True)

    title1.to_excel(writer, sheet_name='1-超时未竣工工单通报', startcol=1, index=False, header=True)
    QZ_Overtime.to_excel(writer, sheet_name='1-超时未竣工工单通报', startrow=1, index=False, header=True)

    title2.to_excel(writer, sheet_name='1-超时未竣工工单通报', startcol=13, index=False, header=True)
    PT_Overtime.to_excel(writer, sheet_name='1-超时未竣工工单通报', startrow=1, startcol=12, index=False, header=True)

    total.to_excel(writer, sheet_name='2-即将超时工单通报', startrow=0, startcol=0, index=False, header=True)

    title3.to_excel(writer, sheet_name='2-即将超时工单通报', startrow=0, startcol=5, index=False, header=True)
    QZ_Timeout_data.to_excel(writer, sheet_name='2-即将超时工单通报', startrow=1, startcol=5, index=False, header=True)

    title4.to_excel(writer, sheet_name='2-即将超时工单通报', startrow=0, startcol=12, index=False, header=True)
    PT_Timeout_data.to_excel(writer, sheet_name='2-即将超时工单通报', startrow=1, startcol=12, index=False, header=True)

end_time = time.time()
print('处理完毕!!!总耗时%0.0f秒钟' % (end_time - start_time))
