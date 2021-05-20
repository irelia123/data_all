import pandas as pd
import time, datetime

# 导入日期
time = datetime.datetime.now()
todaytime = datetime.datetime.strptime('2021-5-13', '%Y-%m-%d')
todaytime1 = todaytime.strftime('%m') + '/' + todaytime.strftime('%d')  # 日期转换str：XX/XX


# 计算平均值函数
def time_mean(data, name='平均时长'):
    gf_time = data[['区县', '总时长']].groupby(['区县']).mean().reset_index()
    gf_time = gf_time.append(
        [{'区县': '合计', '总时长': data[['区县', '总时长']].mean().reset_index(name='总时长').at[0, '总时长']}],
        ignore_index=True)  # 计算全区
    gf_time = gf_time.round({'总时长': 2})  # 四舍五入
    gf_time.rename(columns={'总时长': name}, inplace=True)
    return gf_time


# 计算总数函数
def data_total(data, name):
    number = data.groupby(['区县']).size().reset_index()
    number.rename(columns={0: '总数'}, inplace=True)
    number = number.append([{'区县': '合计', '总数': number.apply(lambda x: x.sum()).总数}], ignore_index=True)
    number.rename(columns={'总数': name}, inplace=True)
    return number


def make_excel(department):
    # 读取表格
    QuXian = pd.DataFrame({'区县': ['江州区', '扶绥县', '宁明县', '龙州县', '大新县', '凭祥市', '天等县', '合计']})
    Group_failure = pd.read_excel('截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx', sheet_name=1)
    Failure_follow = pd.read_excel('截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx', sheet_name=2)
    # 准备工作
    Failure_follow.rename(columns={'通知/移交时间': '时间'}, inplace=True)  # 改列名
    Failure_follow.rename(columns={'县份': '区县'}, inplace=True)  # 改列名
    Failure_follow['现在时'] = (time - pd.to_datetime(Failure_follow.时间)).dt.days * 24 + (
            time - pd.to_datetime(Failure_follow.时间)).dt.seconds / 3600  # 计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
    # 筛选所需要的数据
    GF_Kx = Group_failure[(Group_failure.影响用户数 >= 10) & (Group_failure.责任部门 == department)]  # 筛选数据
    FF_Kx = Failure_follow[(Failure_follow.影响用户 >= 10) & (Failure_follow.责任部门 == department)]  # 筛选数据
    FF_48 = FF_Kx[FF_Kx.现在时 >= 48]
    Recovery = FF_Kx[FF_Kx.恢复时间 != '未恢复']  # 筛选数据
    Accumulation = FF_Kx[FF_Kx.恢复时间 == '未恢复']  # 筛选数据
    GF_Kxd = Group_failure[Group_failure.责任部门 == department]  # 筛选数据
    Kx_City = GF_Kx[GF_Kx.所属区域 == '城镇']  # 城镇数据
    Kx_Rural = GF_Kx[GF_Kx.所属区域 == '农村']  # 农村数据
    # 转换日期格式
    FF_Kx = pd.DataFrame(FF_Kx)
    FF_Kx['时间'] = FF_Kx['时间'].dt.strftime('%Y/%m/%d')
    Recovery = pd.DataFrame(Recovery)
    Recovery['恢复时间'] = pd.to_datetime(Recovery['恢复时间'], format='%Y/%m/%d')
    Recovery['恢复时间'] = Recovery['恢复时间'].dt.strftime('%Y/%m/%d')
    Newly_added = FF_Kx[FF_Kx.时间.str.contains(todaytime1) == True]  # 筛选今日数据的数据，，，，，false为剔除
    Recovery = Recovery[Recovery.恢复时间.str.contains(todaytime1) == True]  # 筛选今日数据的数据，，，，，false为剔除
    # 城镇
    GF_City = pd.merge(QuXian, data_total(Kx_City, '城镇'), on=['区县'], how='left')  # 拼接
    # 农村
    Failure = pd.merge(GF_City, data_total(Kx_Rural, '农村'), on=['区县'], how='left')  # 拼接
    # 线路故障总数
    Failure['线路故障总数'] = Failure.城镇 + Failure.农村
    # 计算平均时长
    Failure = pd.merge(Failure, time_mean(Kx_City, '城镇平均时长'), on=['区县'], how='left')  # 拼接
    Failure = pd.merge(Failure, time_mean(Kx_Rural, '农村平均时长'), on=['区县'], how='left')  # 拼接
    Failure = pd.merge(Failure, time_mean(GF_Kx, '线路故障平均时长'), on=['区县'], how='left')  # 拼接
    # 这里先算总数
    Failure1 = pd.merge(QuXian, data_total(GF_Kxd[(GF_Kxd.总时长 <= 18) & (GF_Kxd.所属区域 == '城镇')], '数'), on=['区县'],
                        how='left')
    Failure1 = pd.merge(Failure1, data_total(GF_Kxd[(GF_Kxd.总时长 <= 36) & (GF_Kxd.所属区域 == '农村')], '数1'), on=['区县'],
                        how='left')
    Failure1 = pd.merge(Failure1, data_total(GF_Kxd, '数2'), on=['区县'], how='left')
    Failure1 = Failure1.fillna(0)
    # 整体处理及时率
    Failure['整体处理及时率'] = (Failure1.数 + Failure1.数1) / Failure1.数2
    Failure['整体处理及时率'] = Failure['整体处理及时率'].apply(lambda x: '%.2f%%' % (x * 100))
    # 影响用户数
    userNumber = pd.DataFrame(GF_Kx.groupby('区县')['影响用户数'].sum()).reset_index()
    userNumber = userNumber.append([{'区县': '合计', '影响用户数': userNumber.apply(lambda x: x.sum()).影响用户数}],
                                   ignore_index=True)
    Failure = pd.merge(Failure, userNumber, on=['区县'], how='left')  # 拼接
    # 今日新增
    Failure = pd.merge(Failure, data_total(Newly_added, '今日新增'), on=['区县'], how='left')  # 拼接
    # 今日恢复
    Failure = pd.merge(Failure, data_total(Recovery, '今日恢复'), on=['区县'], how='left')  # 拼接
    # 积压故障
    Failure = pd.merge(Failure, data_total(Accumulation, '积压故障'), on=['区县'], how='left')  # 拼接
    # 超48小时工单数
    Failure = pd.merge(Failure, data_total(FF_48, '超48小时工单数'), on=['区县'], how='left')  # 拼接
    # 超48小时应影响用户数
    userNumber48 = pd.DataFrame(FF_48.groupby('区县')['影响用户'].sum()).reset_index()
    userNumber48 = userNumber48.append([{'区县': '合计', '影响用户': userNumber48.apply(lambda x: x.sum()).影响用户}],
                                   ignore_index=True)
    userNumber48.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名
    Failure = pd.merge(Failure, userNumber48, on=['区县'], how='left')  # 拼接
    Failure = Failure.fillna(0)
    return Failure


def make_excel2(department):
    # 读取表格
    QuXian = pd.DataFrame({'区县': ['江州区', '扶绥县', '宁明县', '龙州县', '大新县', '凭祥市', '天等县', '合计']})
    Group_failure = pd.read_excel('截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx', sheet_name=1)
    Failure_follow = pd.read_excel('截至2021年5月12日有线专业线路故障修复完成情况（影响用户数超10户）(新版).xlsx', sheet_name=2)
    Group_failure['责任部门'] = '整体'
    Failure_follow['责任部门'] = '整体'
    # 准备工作
    Failure_follow.rename(columns={'通知/移交时间': '时间'}, inplace=True)  # 改列名
    Failure_follow.rename(columns={'县份': '区县'}, inplace=True)  # 改列名
    Failure_follow['现在时'] = (time - pd.to_datetime(Failure_follow.时间)).dt.days * 24 + (
            time - pd.to_datetime(Failure_follow.时间)).dt.seconds / 3600  # 计算时间工公式    备注：公式是时间算的不能按元Excel数据修改
    # 筛选所需要的数据
    GF_Kx = Group_failure[(Group_failure.影响用户数 >= 10) & (Group_failure.责任部门 == department)]  # 筛选数据
    FF_Kx = Failure_follow[(Failure_follow.影响用户 >= 10) & (Failure_follow.责任部门 == department)]  # 筛选数据
    FF_48 = FF_Kx[FF_Kx.现在时 >= 48]
    Recovery = FF_Kx[FF_Kx.恢复时间 != '未恢复']  # 筛选数据
    Accumulation = FF_Kx[FF_Kx.恢复时间 == '未恢复']  # 筛选数据
    GF_Kxd = Group_failure[Group_failure.责任部门 == department]  # 筛选数据
    Kx_City = GF_Kx[GF_Kx.所属区域 == '城镇']  # 城镇数据
    Kx_Rural = GF_Kx[GF_Kx.所属区域 == '农村']  # 农村数据
    # 转换日期格式
    FF_Kx = pd.DataFrame(FF_Kx)
    FF_Kx['时间'] = FF_Kx['时间'].dt.strftime('%Y/%m/%d')
    Recovery = pd.DataFrame(Recovery)
    Recovery['恢复时间'] = pd.to_datetime(Recovery['恢复时间'], format='%Y/%m/%d')
    Recovery['恢复时间'] = Recovery['恢复时间'].dt.strftime('%Y/%m/%d')
    Newly_added = FF_Kx[FF_Kx.时间.str.contains(todaytime1) == True]  # 筛选今日数据的数据，，，，，false为剔除
    Recovery = Recovery[Recovery.恢复时间.str.contains(todaytime1) == True]  # 筛选今日数据的数据，，，，，false为剔除
    # 城镇
    GF_City = pd.merge(QuXian, data_total(Kx_City, '城镇'), on=['区县'], how='left')  # 拼接
    # 农村
    Failure = pd.merge(GF_City, data_total(Kx_Rural, '农村'), on=['区县'], how='left')  # 拼接
    # 线路故障总数
    Failure['线路故障总数'] = Failure.城镇 + Failure.农村
    # 计算平均时长
    Failure = pd.merge(Failure, time_mean(Kx_City, '城镇平均时长'), on=['区县'], how='left')  # 拼接
    Failure = pd.merge(Failure, time_mean(Kx_Rural, '农村平均时长'), on=['区县'], how='left')  # 拼接
    Failure = pd.merge(Failure, time_mean(GF_Kx, '线路故障平均时长'), on=['区县'], how='left')  # 拼接
    # 这里先算总数
    Failure1 = pd.merge(QuXian, data_total(GF_Kxd[(GF_Kxd.总时长 <= 18) & (GF_Kxd.所属区域 == '城镇')], '数'), on=['区县'],
                        how='left')
    Failure1 = pd.merge(Failure1, data_total(GF_Kxd[(GF_Kxd.总时长 <= 36) & (GF_Kxd.所属区域 == '农村')], '数1'), on=['区县'],
                        how='left')
    Failure1 = pd.merge(Failure1, data_total(GF_Kxd, '数2'), on=['区县'], how='left')
    Failure1 = Failure1.fillna(0)
    # 整体处理及时率
    Failure['整体处理及时率'] = (Failure1.数 + Failure1.数1) / Failure1.数2
    Failure['整体处理及时率'] = Failure['整体处理及时率'].apply(lambda x: '%.2f%%' % (x * 100))
    # 影响用户数
    userNumber = pd.DataFrame(GF_Kx.groupby('区县')['影响用户数'].sum()).reset_index()
    userNumber = userNumber.append([{'区县': '合计', '影响用户数': userNumber.apply(lambda x: x.sum()).影响用户数}],
                                   ignore_index=True)
    Failure = pd.merge(Failure, userNumber, on=['区县'], how='left')  # 拼接
    # 今日新增
    Failure = pd.merge(Failure, data_total(Newly_added, '今日新增'), on=['区县'], how='left')  # 拼接
    # 今日恢复
    Failure = pd.merge(Failure, data_total(Recovery, '今日恢复'), on=['区县'], how='left')  # 拼接
    # 积压故障
    Failure = pd.merge(Failure, data_total(Accumulation, '积压故障'), on=['区县'], how='left')  # 拼接
    # 超48小时工单数
    Failure = pd.merge(Failure, data_total(FF_48, '超48小时工单数'), on=['区县'], how='left')  # 拼接
    # 超48小时应影响用户数
    userNumber48 = pd.DataFrame(FF_48.groupby('区县')['影响用户'].sum()).reset_index()
    userNumber48 = userNumber48.append([{'区县': '合计', '影响用户': userNumber48.apply(lambda x: x.sum()).影响用户}],
                                   ignore_index=True)
    userNumber48.rename(columns={'影响用户': '超48小时影响用户数'}, inplace=True)  # 改列名
    Failure = pd.merge(Failure, userNumber48, on=['区县'], how='left')  # 拼接
    Failure = Failure.fillna(0)
    return Failure
Data = pd.DataFrame(make_excel2('整体'))
Data = Data[['今日新增','积压故障','超48小时工单数','超48小时影响用户数']]

with pd.ExcelWriter('群障报表' + '.xlsx') as writer:
    make_excel('客响').to_excel(writer, sheet_name='汇总', startcol=0,startrow=0, index=False, header=True)
    make_excel('传输').to_excel(writer, sheet_name='汇总', startcol=0, startrow=11,index=False, header=True)
    make_excel2('整体').to_excel(writer, sheet_name='汇总', startcol=0,startrow=22, index=False, header=True)
    Data.to_excel(writer, sheet_name='汇总', startcol=0,startrow=33, index=False, header=True)
