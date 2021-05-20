import datetime
import pandas as pd



#################################
today = datetime.datetime.now()  # 实时
today2 = today.strftime('%Y') + '年' + str(today.month) + '月' + today.strftime('%d') + '日'  # 日期转换str：2019年9月10日
# 计算日期完毕!
Bymbh = pd.DataFrame(
    {'地市': ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港', '全区']},
    pd.Index(range(15)))

def restructure(tables_name,n):
    table = pd.read_excel(tables_name, skiprows=2)  # 读表
    table = table.reset_index(drop=True)
    table.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长'}, inplace=True)  # 改列名
    table.rename(columns={'高品质装机时长（农村）': '农村高品质装机时长'}, inplace=True)  # 改列名
    table.rename(columns={'普通品质装机时长（城镇）': '城镇普通品质装机时长'}, inplace=True)  # 改列名
    table.rename(columns={'普通品质装机时长（农村）': '农村普通品质装机时长'}, inplace=True)  # 改列名
    table.rename(columns={'装机时长（整体）': '整体装机时长'}, inplace=True)  # 改列名
    table.rename(columns={'装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名

    table_get = table[['地市', '城镇高品质装机时长', '农村高品质装机时长', '城镇普通品质装机时长', '农村普通品质装机时长', '城镇装机时长', '农村装机时长', '整体装机时长', '高品质装机及时率', '普通品质装机及时率', '整体装机及时率']]
    col = table_get.columns.size  # 总列数
    result = table_get['地市']
    result = result.to_frame()
    for i in range(1, col):
        a = table_get.columns[i]
        linshi2 = table_get[['地市', a]]
        linshi2['排名'] = linshi2[a][:n].rank(axis=0, ascending=True, method='dense')  # 输出排名
        result = pd.merge(result, linshi2, on ='地市')
    result.rename(columns={'排名_x':'排名','排名_y':'排名'},inplace=True)
    return result
q = restructure('魔百和装机及时率报表.xls',14)
w =restructure('IMS装机及时率报表.xls',13)
e = restructure('和目装机及时率报表.xls',12)
r = restructure('平安乡村及时率报表.xls',14)
t =restructure('智能组网装机及时率报表.xls',14)

with pd.ExcelWriter('月报'+'.xlsx') as writer:
    q.to_excel(writer, sheet_name='魔百和装机及时率', startcol=0, index=False, header=True)
    w.to_excel(writer, sheet_name='IMS装机及时率', startcol=0, index=False, header=True)
    e.to_excel(writer, sheet_name='和目装机及时率', startcol=0, index=False, header=True)
    r.to_excel(writer, sheet_name='平安乡村装机及时率', startcol=0, index=False, header=True)
    t.to_excel(writer, sheet_name='智能组网装机及时率', startcol=0, index=False, header=True)
