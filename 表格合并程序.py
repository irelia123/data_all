#导入模块.
import os
import pandas as pd
import re
import xlwt


A = pd.DataFrame()

path = os.getcwd()
shuchu = path + '\\合并表\\'

for root, dirs, files in os.walk(path,topdown=True):
    for file in files:
        lswjm = os.path.splitext(file)
        if re.search(r'xls',lswjm[1]) and root == path:
            print(file)
            df = pd.read_excel(os.path.join(root, file))
            df = df.dropna(how='all')
            df = df.reset_index(drop=True)
            A = A.append(df)

th = 0
for i,vl in enumerate(A.columns.values):

    if re.search(r'Unnamed', vl) and th ==0 and i < 3:
        for jc, xh in enumerate(A.iloc[0].values):
            print(xh)
            try:#遇到NAN报错
                if re.search(r'[序号]', xh):
                    th = 1
                    A.columns = A.iloc[0].values
                    A = A.drop([0])
                    A = A.reset_index(drop=True)
                    break
            except:
                break

    if th == 1:
        break

for i,vl in enumerate(A.columns.values):
    if re.search(r'[序号]', vl):
        A[vl] = [b+1 for b in range(len(A[vl]))]
        break




print("正在写入", shuchu)

if not os.path.exists(shuchu):
    os.makedirs(shuchu)
A.to_csv(shuchu+'总表.csv', index=False, header=True)

print("结束了")

