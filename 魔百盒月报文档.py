#设置固定模板，模板中需要填入数据的位置要设置好域，然后脚本读取代码，用render函数将数据填入文档中并保存。
#哪里的数据出问题，就找到它的域名，然后到代码中找到域名相应的位置

import pandas as pd
from docxtpl import DocxTemplate,InlineImage

path = '2-1．魔百和、IMS、附属产品装机指标情况-6月.xlsx'
date = '2021年6月'

#输出达标情况
def jieguo(table,pd_content,jizhunzhi,tiaozhanzhi,qufan):
    r_tiaozhan = []
    rj_nt = []
    nr_jizhun = []
    index = table.index.values.tolist()
    index.remove('全区')
    for i in index:
        if  table.loc[i,pd_content]*qufan > jizhunzhi*qufan:
            nr_jizhun.append(i)
        elif table.loc[i,pd_content]*qufan < tiaozhanzhi*qufan:
            r_tiaozhan.append(i)
        else:
            rj_nt.append(i)
    num_tiaozhan = str(len(r_tiaozhan))
    num_jt = str(len(rj_nt))
    num_jz = str(len(nr_jizhun))
    if r_tiaozhan != []:
        if rj_nt != []:
            if nr_jizhun != []:
                 content = '有'+num_tiaozhan+'个市公司达到挑战值：'+'、'.join(r_tiaozhan)+';'+'有'+num_jt+'个市公司达基准值未达挑战值：'+'、'.join(rj_nt)+';'+'有'+num_jz+'个市公司未达基准值：'+'、'.join(nr_jizhun)
            else:
                 content = '有'+num_tiaozhan+'个市公司达到挑战值：'+'、'.join(r_tiaozhan)+';'+'有'+num_jt+'个市公司达基准值未达挑战值：'+'、'.join(rj_nt)
        else:
            if nr_jizhun != []:
               content = '有'+num_tiaozhan+'个市公司达到挑战值：'+'、'.join(r_tiaozhan)+';'+'有'+num_jz+'个市公司未达基准值：'+'、'.join(nr_jizhun)
            else:
                 content = '各市公司均达到挑战值'
    else:
        if rj_nt != []:
            if nr_jizhun != []:
                content = '有'+num_jt+'个市公司达基准值未达挑战值：'+'、'.join(rj_nt)+';'+'有'+num_jz+'个市公司未达基准值：'+'、'.join(nr_jizhun)
            else:
                content = '有'+num_jt+'个市公司达基准值未达挑战值：'+'、'.join(rj_nt)
        else:
            if nr_jizhun != []:
                content = '各市公司均未达到基准值'
    return content   
#判断全区是否达标挑战值、基准值
def shifou(value,jizhun,tiaozhan_value,qufan):
    if value*qufan < jizhun*qufan:
        if value*qufan < tiaozhan_value*qufan:
            content = '达到挑战值'
        else:
            content = '达到基准值'
    else:
        content = '未达基准值'
    return content

def mubiaozhi(num):
    if num > 96:
        content = '未'
    else:
        content = ''
    return content
#判断全区是否达标目标值
def tiaozhan_mb(table):
    nr_mubiao = []
    index = table.index.values.tolist()
    index.remove('全区')
    for i in index:
        if  table.loc[i,'装移机平均时长.6'] > 96:
            nr_mubiao.append(i)
    num = str(len(nr_mubiao))
    if nr_mubiao != []:
        content = '有'+num+'个市公司未达到目标值：'+'、'.join(nr_mubiao)
    else:
        content = '各市公司均已达到目标值'
    return content
    
mbh_data =pd.read_excel(path, skiprows=2)  

#魔百盒装机时长
#城镇高品质
cg_data = mbh_data.iloc[:15,:3]
cg_data.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市，方便选取各地市的数据
cgzj = cg_data.loc['全区','装移机平均时长']

#农村高品质
ng_data = mbh_data.iloc[:15,3:6]
ng_data.set_index('地市.1',drop=True, append=False, inplace=True, verify_integrity=False)
ngzj = ng_data.loc['全区','装移机平均时长.1']

#城镇普通品质
cp_data = mbh_data.iloc[:15,6:9]
cp_data.set_index('地市.2',drop=True, append=False, inplace=True, verify_integrity=False)
cpzj = cp_data.loc['全区','装移机平均时长.2']

#农村普通品质
np_data = mbh_data.iloc[:15,9:12]
np_data.set_index('地市.3',drop=True, append=False, inplace=True, verify_integrity=False)
npzj = np_data.loc['全区','装移机平均时长.3']
#魔百和装机及时率
#全区高品质
gq_data =mbh_data.iloc[:15,22:25]
gq_data.set_index('地市.7',drop=True, append=False, inplace=True, verify_integrity=False)
g_jsl = gq_data.loc['全区','装移机及时率']

#全区普通品质
pq_data =mbh_data.iloc[:15,25:28]
pq_data.set_index('地市.8',drop=True, append=False, inplace=True, verify_integrity=False)
p_jsl = pq_data.loc['全区','装移机及时率.1']

#IMS装机时长
IMS_data = pd.read_excel(path,sheet_name='IMS装机及时率', skiprows=2)
IMScg_data = IMS_data.iloc[0:15,:3]
#剔除无数据地市
IMScg_data = IMScg_data[~IMScg_data['装移机平均时长'].isin(['-'])]
IMScg_data.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)
cgIMS = IMScg_data.loc['全区','装移机平均时长']

#农村高品质IMS
IMSng_data = IMS_data.iloc[0:15,3:6]
#剔除无数据地市
IMSng_data = IMSng_data[~IMSng_data['装移机平均时长.1'].isin(['-'])]
IMSng_data.set_index('地市.1',drop=True, append=False, inplace=True, verify_integrity=False)
ngIMS = IMSng_data.loc['全区','装移机平均时长.1']

#城镇普通品质IMS
IMScp_data = IMS_data.iloc[0:15,6:9]
#剔除无数据地市
IMScp_data = IMScp_data[~IMScp_data['装移机平均时长.2'].isin(['-'])]
IMScp_data.set_index('地市.2',drop=True, append=False, inplace=True, verify_integrity=False)
cpIMS = IMScp_data.loc['全区','装移机平均时长.2']

#农村普通品质IMS
IMSnp_data = IMS_data.iloc[0:15,9:12]
#剔除无数据地市
IMSnp_data = IMSnp_data[~IMSnp_data['装移机平均时长.3'].isin(['-'])]
IMSnp_data.set_index('地市.3',drop=True, append=False, inplace=True, verify_integrity=False)
npIMS = IMSnp_data.loc['全区','装移机平均时长.3']

#IMS装机及时率
#全区高品质
imsgq_data =IMS_data.iloc[:15,22:25]
imsgq_data = imsgq_data[~imsgq_data['装移机及时率'].isin(['-'])]
imsgq_data.set_index('地市.7',drop=True, append=False, inplace=True, verify_integrity=False)
g_ghzjlv = imsgq_data.loc['全区','装移机及时率']

#全区普通品质
imspq_data =IMS_data.iloc[:15,25:28]
imspq_data.set_index('地市.8',drop=True, append=False, inplace=True, verify_integrity=False)
p_ghzjlv = imspq_data.loc['全区','装移机及时率.1']

#和目、平安乡村、智能组网装机时长
hemu_data =pd.read_excel(path,sheet_name='和目装机及时率', skiprows=2)  
pingan_data =pd.read_excel(path,sheet_name='平安乡村装机及时率', skiprows=2) 
zhineng_data =pd.read_excel(path,sheet_name='智能组网装机及时率', skiprows=2) 
hmall_data = hemu_data.iloc[:15,18:21]
pingall_data = pingan_data.iloc[:15,18:21]
znall_data = zhineng_data.iloc[:15,18:21]
hmall_data.set_index('地市.6',drop=True, append=False, inplace=True, verify_integrity=False)
pingall_data.set_index('地市.6',drop=True, append=False, inplace=True, verify_integrity=False)
znall_data.set_index('地市.6',drop=True, append=False, inplace=True, verify_integrity=False)
hmzj = hmall_data.loc['全区','装移机平均时长.6']
paxczj = pingall_data.loc['全区','装移机平均时长.6']
znzwzj = znall_data.loc['全区','装移机平均时长.6']

hmjsl_data =hemu_data.iloc[:15,28:31]
hmjsl_data.set_index('地市.9',drop=True, append=False, inplace=True, verify_integrity=False)
hmzjlv = hmjsl_data.loc['全区','装移机及时率.2']

panjsl_data =pingan_data.iloc[:15,28:31]
panjsl_data.set_index('地市.9',drop=True, append=False, inplace=True, verify_integrity=False)
paxczjlv = panjsl_data.loc['全区','装移机及时率.2']

znengjsl_data =zhineng_data.iloc[:15,28:31]
znengjsl_data.set_index('地市.9',drop=True, append=False, inplace=True, verify_integrity=False)
znzwzjlv = znengjsl_data.loc['全区','装移机及时率.2']

contexts = []
context = {"cgzj":cgzj,"shifou1":shifou(cgzj,24,16,1),"tiaozhan1":jieguo(cg_data,'装移机平均时长',24,16,1),
          "ngzj":ngzj,"shifou2":shifou(ngzj,36,28,1),"tiaozhan2":jieguo(ng_data,'装移机平均时长.1',36,28,1),
          "cpzj":cpzj,"shifou3":shifou(cpzj,48,40,1),"tiaozhan3":jieguo(cp_data,'装移机平均时长.2',48,40,1),
          "npzj":npzj,"shifou4":shifou(npzj,72,64,1),"tiaozhan4":jieguo(np_data,'装移机平均时长.3',72,64,1),
          "g_jsl":'%.2f%%' % (g_jsl * 100),"shifou5":shifou(g_jsl,0.93,0.96,-1),"tiaozhan5":jieguo(gq_data,'装移机及时率',0.93,0.96,-1),
          "p_jsl":'%.2f%%' % (p_jsl * 100),"shifou6":shifou(p_jsl,0.93,0.96,-1),"tiaozhan6":jieguo(pq_data,'装移机及时率.1',0.93,0.96,-1),
          "cgIMS":round(cgIMS,2),"shifou7":shifou(cgIMS,24,16,1),"tiaozhan7":jieguo(IMScg_data,'装移机平均时长',24,16,1),
          "ngIMS":round(ngIMS,2),"shifou8":shifou(ngIMS,36,28,1),"tiaozhan8":jieguo(IMSng_data,'装移机平均时长.1',36,28,1),
          "cpIMS":round(cpIMS,2),"shifou9":shifou(cpIMS,48,40,1),"tiaozhan9":jieguo(IMScp_data,'装移机平均时长.2',48,40,1),
          "npIMS":round(npIMS,2),"shifou10":shifou(npIMS,72,64,1),"tiaozhan10":jieguo(IMSnp_data,'装移机平均时长.3',72,64,1),
          "g_ghzjlv":'%.2f%%' % (g_ghzjlv * 100),"shifou11":shifou(g_ghzjlv,0.93,0.96,-1),"tiaozhan11":jieguo(imsgq_data,'装移机及时率',0.93,0.96,-1),
          "p_ghzjlv":'%.2f%%' % (p_ghzjlv * 100),"shifou12":shifou(p_ghzjlv,0.93,0.96,-1),"tiaozhan12":jieguo(imspq_data,'装移机及时率.1',0.93,0.96,-1),
          "hmzj":hmzj,"shifou13":mubiaozhi(hmzj),"tiaozhan13":tiaozhan_mb(hmall_data),
          "paxczj":paxczj,"shifou14":mubiaozhi(paxczj),"tiaozhan14":tiaozhan_mb(pingall_data),
          "znzwzj":znzwzj,"shifou15":mubiaozhi(znzwzj),"tiaozhan15":tiaozhan_mb(znall_data),
          "hmzjlv":'%.2f%%' % (hmzjlv * 100),"shifou16":shifou(hmzjlv,0.93,0.96,-1),"tiaozhan16":jieguo(hmjsl_data,'装移机及时率.2',0.93,0.96,-1),
          "paxczjlv":'%.2f%%' % (paxczjlv * 100),"shifou17":shifou(paxczjlv,0.93,0.96,-1),"tiaozhan17":jieguo(panjsl_data,'装移机及时率.2',0.93,0.96,-1),
          "znzwzjlv":'%.2f%%' % (znzwzjlv * 100),"shifou18":shifou(znzwzjlv,0.93,0.96,-1),"tiaozhan18":jieguo(znengjsl_data,'装移机及时率.2',0.93,0.96,-1)
          } #变量名称与Word文档中的占位符要一一对应
contexts.append(context)

for context in contexts:
    doc = DocxTemplate(r"魔百盒月报模板.docx")  #需要填入的Word文档的的地址
    doc.render(context)   #将数据填入模板
    doc.save("2-3．关于"+date+"家庭宽带扩展业务装维管理工作情况的通报.docx")  #保存
    
    
