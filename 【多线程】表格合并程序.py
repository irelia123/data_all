import threading
import os
import pandas as pd
import re

hebingbiao = pd.DataFrame()  # 合并表

path = os.getcwd()  # py代码当前所在目录
shuchu = path + '\\合并表\\'  # 合成表输出路径

threads = []  # 线程句柄数组


def dubiao(root, file):  # 读取Excel文件的函数
    print(file, "正在读取...")
    df = pd.read_excel(os.path.join(root, file))
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)
    print(file, "读取完毕！")
    return df


class thread(threading.Thread):  # 线程入口
    def __init__(self, root, file):  # 线程初始化
        threading.Thread.__init__(self, name='线程' + file)
        self.file = str(file)
        self.root = str(root)

    def run(self):  # 线程主体代码

        df = dubiao(self.root, self.file)

        lock.acquire()  # 线程锁【锁定】：当多条线程访问同一个变量时，为了防止冲突，加锁让线程有序运行，
        global hebingbiao  # 访问外部变量必须要声明global
        hebingbiao = hebingbiao.append(df)
        print(self.file, "已成功合并！")
        lock.release()  # 线程锁【解锁】：当前代码执行完毕时，必须先解锁，下一条线程才能允许进入


lock = threading.Lock()  # 创建线程锁

for root, dirs, files in os.walk(path, topdown=True):  # 遍历当前目录
    for file in files:  # 遍历文件名
        lswjm = os.path.splitext(file)
        if re.search(r'xls', lswjm[1]) and root == path:  # 判断文件后缀
            threads.append(thread(root, file))  # 创建一条线程，将线程所需要的参数传递进去

for t in threads:  # 开启线程
    t.start()

for t in threads:  # 阻塞线程 当所有线程结束后才能运行下一步
    t.join()

for i in range(3):  # 检测表中第一行是否是表头，检测三次
    # 如果当前行存在空单元格（Unnamed），以第二行作为表头，删除第一行，往复循环
    if re.search(r'Unnamed', str(hebingbiao.columns.values)):
        hebingbiao.columns = hebingbiao.iloc[0].values
        hebingbiao = hebingbiao.drop([0])
        hebingbiao = hebingbiao.reset_index(drop=True)

for i, vl in enumerate(hebingbiao.columns.values):  # 遍历列名
    if vl == '序号':  # 如果存在序号列，进行索引重置
        hebingbiao[vl] = [b + 1 for b in range(len(hebingbiao[vl]))]
        break

print("正在写入", shuchu)

if not os.path.exists(shuchu):  # 输出目录不存在则创建目录
    os.makedirs(shuchu)
# 保留原格式，但导出速度较慢
hebingbiao.to_excel(shuchu + '总表.xlsx', index=False, header=True)

# 快速导出，不保留格式
# hebingbiao.to_csv(shuchu + '总表.csv', index=False, header=True)


print("写入完成！")
