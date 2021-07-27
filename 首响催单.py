import pandas as pd
import datetime
import time

city = '钦州'

# 两种模式，可取实时时间和自定义时间(统计工作时用)
# today_now = datetime.datetime.now()  # 系统取实时时间(精确到时分秒)
today_now = datetime.datetime.strptime('2021-7-22 13:52:00', '%Y-%m-%d %H:%M:%S')  # 自定义制表时间

COMPLAINT_PATH = '家庭业务投诉工单查询导出.xlsx'

LIST_PATH = '程序所需数据表.xlsx'


# 判断函数(最好只用于数据较小的表格)
def match_table(obj, mat, key):
    map1 = obj[key].map(lambda x: x in list(mat[key]))  # 两个表之间一列进行判断
    map1 = pd.DataFrame(map1)
    map1.rename(columns={key: '判断'}, inplace=True)  # 改列名
    obj = pd.concat([obj, map1], axis=1)
    return obj


# 批量修改表格格式为str
def format_str(data, name):
    data_types_dict = {name: str}
    data = data.astype(data_types_dict)
    return data


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


# 剔除重复人员
list_data = pd.read_excel(LIST_PATH, sheet_name='网格人员名单')  # 人员名单表
list_data = list_data.drop_duplicates(subset=['施工人员'], keep='first')
# 剔除重复异常工单
ex = pd.read_excel(LIST_PATH, sheet_name='异常工单')
ex = ex.drop_duplicates(subset=['工单流水号'], keep='first')
# 网格长名单
captain = pd.read_excel(LIST_PATH, sheet_name='网格长名单')
captain.loc[captain.shape[0], '网格'] = '合计'
# 区县
county = pd.read_excel(LIST_PATH, sheet_name='区县')
county = pd.DataFrame(county[['区县']])
county.loc[county.shape[0], '区县'] = '合计'

complaint = pd.read_excel(COMPLAINT_PATH)
complaint['最后质检通过时间'].fillna('空', inplace=True)
complaint['联系客户时间（客户接通当次的装维拨打时间）'].fillna('空', inplace=True)

complaint = complaint[(complaint.最后质检通过时间 == '空') & (complaint.工单类型 != '和家亲')
                      & (complaint.工单类型 != '问题解决单') & (complaint.责任地市 == city)]  # 筛选数据

complaint_data = pd.DataFrame(
    complaint[['首次到单时间', '责任地市', '首次联系客户时间', '首次调度施工人员', '速率', '联系客户时间（客户接通当次的装维拨打时间）', '工单流水号']])
complaint_data.rename(columns={'首次到单时间': '工单开始时间', '首次联系客户时间': '首次预约时间', '首次调度施工人员': '施工人员', '责任地市': '地市'},
                      inplace=True)  # 改列名

complaint_data['是否千兆'] = 'FALSE'
complaint_data['是否首响'] = '是'
complaint_data['是否异常'] = '否'
complaint_data['是否超时'] = '剩余30分钟以内'

complaint_data.loc[(complaint_data['速率'] == '1000M'), '是否千兆'] = 'TRUE'
complaint_data.loc[(complaint_data['联系客户时间（客户接通当次的装维拨打时间）'] == '空'), '是否首响'] = '否'

complaint_data = pd.merge(complaint_data, list_data, on='施工人员', how='left')

complaint_data.rename(columns={'装维网格名称': '网格', '所属区县': '区县'}, inplace=True)  # 改列名

complaint_data = match_table(complaint_data, ex, '工单流水号')

complaint_data.loc[(complaint_data['判断'] == True), '是否异常'] = '是'

# 计算工作时间(剔除夜间时间)
complaint_data['工单开始时间'] = pd.to_datetime(complaint_data['工单开始时间'])  # 转换日期格式

# 现在时
complaint_data['现在时'] = today_now
# 开始时间的当日20点
complaint_data['开始当日二十点'] = complaint_data['工单开始时间'].dt.strftime('%Y-%m-%d')
complaint_data['开始当日二十点'] = complaint_data.开始当日二十点 + ' 20:00:00'
complaint_data['开始当日二十点'] = pd.to_datetime(complaint_data['开始当日二十点'])
# 开始时间的当日8点
complaint_data['开始当日八点'] = complaint_data['工单开始时间'].dt.strftime('%Y-%m-%d')
complaint_data['开始当日八点'] = complaint_data.开始当日八点 + ' 08:00:00'
complaint_data['开始当日八点'] = pd.to_datetime(complaint_data['开始当日八点'])
# 到单第二日8点
complaint_data['开始第二日八点'] = pd.to_datetime(complaint_data['工单开始时间'], format='%Y-%m-%d')
complaint_data['开始第二日八点'] = (pd.to_datetime(complaint_data['工单开始时间']) + datetime.timedelta(days=1))
complaint_data['开始第二日八点'] = complaint_data['开始第二日八点'].dt.strftime('%Y-%m-%d')
complaint_data['开始第二日八点'] = complaint_data.开始第二日八点 + ' 08:00:00'
complaint_data['开始第二日八点'] = pd.to_datetime(complaint_data['开始第二日八点'])
# 实际计算开始时间
complaint_data['实际开始时间'] = complaint_data['工单开始时间']
complaint_data.loc[(complaint_data['实际开始时间'] > complaint_data['开始当日二十点']), '实际开始时间'] = complaint_data['开始第二日八点']
complaint_data.loc[(complaint_data['实际开始时间'] < complaint_data['开始当日八点']), '实际开始时间'] = complaint_data['开始当日八点']

# 开始日期(计算天数用)
complaint_data['开始日期'] = pd.to_datetime(complaint_data['实际开始时间'], format='%Y-%m-%d')
complaint_data['开始日期'] = complaint_data['开始日期'].dt.strftime('%Y-%m-%d')
complaint_data['开始日期'] = pd.to_datetime(complaint_data['开始日期'])

# 现在日期(计算天数用)
complaint_data['当天日期'] = pd.to_datetime(complaint_data['现在时'], format='%Y-%m-%d')
complaint_data['当天日期'] = complaint_data['当天日期'].dt.strftime('%Y-%m-%d')
complaint_data['当天日期'] = pd.to_datetime(complaint_data['当天日期'])
# 相差天数
complaint_data['相差天数'] = (pd.to_datetime(complaint_data['当天日期']) - pd.to_datetime(complaint_data['开始日期'])).dt.days
# 计算工单剩余时间
complaint_data['工单剩余时间'] = (pd.to_datetime(complaint_data['现在时']) - pd.to_datetime(complaint_data['实际开始时间'])).astype(
    'timedelta64[s]') / 3600
complaint_data['工单剩余时间'] = complaint_data.工单剩余时间 - (complaint_data.相差天数 * 12)
complaint_data = complaint_data.round({'工单剩余时间': 2})  # 四舍五入

complaint_data.rename(columns={'工单剩余时间': today_now}, inplace=True)  # 改列名

# 判断超时时间
complaint_data.loc[(complaint_data[today_now] >= 0.17), '是否超时'] = '剩余20分钟以内'
complaint_data.loc[(complaint_data[today_now] >= 0.33), '是否超时'] = '剩余10分钟以内'
complaint_data.loc[(complaint_data[today_now] >= 0.5), '是否超时'] = '超30分钟'
complaint_data.loc[(complaint_data[today_now] >= 1.0), '是否超时'] = '超60分钟'
complaint_data.loc[(complaint_data[today_now] > 1.01), '是否超时'] = '超60分钟以上'
# 修改日期格式为str
complaint_data = format_str(complaint_data, '工单开始时间')
# 只取需要的数据
complaint_data = pd.DataFrame(
    complaint_data[[today_now, '地市', '是否首响', '是否异常', '是否超时', '区县', '网格', '施工人员', '工单开始时间', '首次预约时间', '是否千兆', '工单流水号']])

complaint = pd.merge(complaint_data, complaint, on='工单流水号')


# 统计数据
# 只取需要统计的数据
def make_excel(data, col, off=0):
    statistical_table = data[(data.是否首响 == '否') & (data.是否异常 == '否')]

    table = data_total(statistical_table, col, '未首响工单数', off)

    table = pd.merge(table,
                     data_total(statistical_table[statistical_table[today_now] <= 1], col, '可有效首响工单', off),
                     on=col, how='left')

    table = pd.merge(table,
                     data_total(statistical_table[statistical_table[today_now] < 0.5], col, '小于30分钟未首响工单数', off),
                     on=col, how='left')

    table = pd.merge(table,
                     data_total(
                         statistical_table[
                             (statistical_table[today_now] >= 0.5) & (statistical_table[today_now] < 1)],
                         col, '大于等于30分钟未首响工单数', off),
                     on=col, how='left')
    table = pd.merge(table,
                     data_total(statistical_table[statistical_table[today_now] >= 1], col, '大于等于60分钟未首响工单数', off),
                     on=col, how='left')
    table = table.fillna(0)
    return table


# 第二页(地市)
city_table = make_excel(complaint, '地市')
# 第二页(区县)
county_table = make_excel(complaint, '区县', 1)
county_table = pd.merge(county, county_table, on='区县', how='left')
county_table = county_table.fillna(0)
# 第二页(网格)
grid_table = make_excel(complaint, '网格', 1)
grid_table = pd.merge(captain, grid_table, on='网格', how='left')
grid_table = grid_table.fillna(0)

# 第四页(通报)
no_response = complaint[(complaint.是否首响 == '否') & (complaint.是否异常 == '否')]

no_response = pd.DataFrame(
    no_response[['是否超时', '区县', '网格', '施工人员', '工单流水号', '姓名', '客户电话', '联系电话1', '宽带账号', '来电号码', '安装地址']])

no_response.rename(columns={'是否超时': '首响超时', '姓名': '客户姓名'}, inplace=True)  # 改列名

# 按照指定区县排序
c = pd.DataFrame({'区县': ['城区', '钦州港区', '灵山县', '浦北县']})

no_response = pd.merge(c, no_response, on='区县')

col = list(no_response.columns.values)
col.insert(0, col[1])
col.pop(1 + 1)
no_response = no_response[col]

with pd.ExcelWriter(f'({city})有效首响催单通报.xlsx') as writer:  # 写入结果为当前路径
    complaint.to_excel(writer, sheet_name='在途详单', startcol=0, index=False, header=True)
    city_table.to_excel(writer, sheet_name='区县', startcol=0, index=False, header=True)
    county_table.to_excel(writer, sheet_name='区县', startcol=0, startrow=3, index=False, header=True)
    grid_table.to_excel(writer, sheet_name='网格', startcol=0, index=False, header=True)
    no_response.to_excel(writer, sheet_name='通报', startcol=0, index=False, header=True)
