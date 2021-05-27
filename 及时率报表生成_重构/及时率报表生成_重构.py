import pandas as pd

def jsl(tables_name, n):  # 表名、排名行数
    table = pd.read_excel(tables_name, skiprows=2)  # 读表
    table = table.reset_index(drop=True)
    table.rename(columns={'高品质装机时长（城镇）': '城镇高品质装机时长', '高品质装机时长（农村）': '农村高品质装机时长',
                          '普通品质装机时长（城镇）': '城镇普通品质装机时长', '普通品质装机时长（农村）': '农村普通品质装机时长',
                          '装机时长（整体）': '整体装机时长', '装机及时率（整体）': '整体装机及时率'}, inplace=True)  # 改列名
    table_get = table[['地市', '城镇高品质装机时长', '农村高品质装机时长', '城镇普通品质装机时长', '农村普通品质装机时长',
                       '城镇装机时长', '农村装机时长', '整体装机时长', '高品质装机及时率', '普通品质装机及时率', '整体装机及时率']]  # 提取相关内容
    col = table_get.columns.size  # 总列数
    result = table_get['地市']
    result = result.to_frame()  # 搞一个DataFrame，后面用于拼接
    for i in range(1, col):
        a = table_get.columns[i]
        nrank = table_get[['地市', a]]
        nrank['排名'] = nrank[a][:n].rank(axis=0, ascending=True, method='dense')  # 输出排名
        result = pd.merge(result, nrank, on='地市')
    result.rename(columns={'排名_x': '排名', '排名_y': '排名'}, inplace=True)
    return result

with pd.ExcelWriter('及时率报表'+'.xlsx') as writer:
    jsl('魔百和装机及时率报表.xls', 14).to_excel(writer, sheet_name='魔百和装机及时率', startcol=0, index=False, header=True)
    jsl('IMS装机及时率报表.xls', 13).to_excel(writer, sheet_name='IMS装机及时率', startcol=0, index=False, header=True)
    jsl('和目装机及时率报表.xls', 12).to_excel(writer, sheet_name='和目装机及时率', startcol=0, index=False, header=True)
    jsl('平安乡村及时率报表.xls', 14).to_excel(writer, sheet_name='平安乡村装机及时率', startcol=0, index=False, header=True)
    jsl('智能组网装机及时率报表.xls', 14).to_excel(writer, sheet_name='智能组网装机及时率', startcol=0, index=False, header=True)  # 生成表格
