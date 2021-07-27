import pandas as pd
from docxtpl import DocxTemplate,InlineImage

date = '6月'  #手动修改日期
#需要填入Word的Excel工作簿的地址
path_jk = "1-1．家宽装机指标分析（2021年6月).xlsx"  #手动修改需要读取数据的Excel
path_rb = '程序生成日报2021年7月01日（剔除校园工单）.xlsx'
path_dkzb = '重点盯控指标-6月.xlsx'

def col_name(table,dishi,zyzb,paiming):
    table.rename(columns={dishi: '地市'}, inplace=True)  # 改列名
    table.rename(columns={zyzb: '装移机平均时长'}, inplace=True)  # 改列名
    table.rename(columns={paiming: '排名'}, inplace=True)
    table.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)
    table['装移机平均时长'] = round(table['装移机平均时长'],2)

def tiaozhan(table,pd_content,challe_data,jizhunzhi,qufan):
    #未达挑战值
    not_challe1 = []
    not_jizhun = []
    for i in index:
        if  table.loc[i,pd_content]*qufan > jizhunzhi*qufan:
            not_jizhun.append(i)
        elif table.loc[i,pd_content]*qufan > challe_data*qufan:
            not_challe1.append(i)
    
    if not_jizhun != []:
        if not_challe1 != []:
            num_challe = len(not_challe1)
            num_challe =str(num_challe)
            num_jz = len(not_jizhun)
            num_jz =str(num_jz)
            content = '有'+num_challe+'个市公司达基准值未达挑战值：'+'、'.join(not_challe1)+';'+'有'+num_jz+'个市公司未达基准值：'+'、'.join(not_jizhun)
        else:
            num_jz = len(not_jizhun)
            num_jz =str(num_jz)
            content = '有'+num_jz+'个市公司未达基准值：'+'、'.join(not_jizhun)
    else:
        if not_challe1 != []:
           num_challe = len(not_challe1)
           num_challe =str(num_challe) 
           content = '有'+num_challe+'个市公司达基准值未达挑战值：'+'、'.join(not_challe1)
        else:
             content = '各市公司均达挑战值'
    return content

def mubiao(table,content,mubiaozhi,qf,word):
    not_mubiao = []
    for i in index:
        if  table.loc[i,content] *qf > mubiaozhi*qf:
            not_mubiao.append(i)
    if not_mubiao != []:
        num_mubiao = len(not_mubiao)
        num_mubiao = str(num_mubiao)
        content = '有'+num_mubiao+'个市公司未达到'+word+':'+ '、'.join(not_mubiao)
    else:
        content = '各市公司均达'+ word
    return content

def tiaozhan1(num,tiaozhanzhi,str_tiaozhanzhi,qf):
    if num*qf > tiaozhanzhi*qf:
        result = "达到" + str_tiaozhanzhi+"）"
    else:
        result = "未达到" + str_tiaozhanzhi+"）"
    return result

jk_data =pd.read_excel(path_jk, skiprows=2)  
qlc_data = pd.read_excel(path_jk,sheetname = '全流程时长、及时率',skiprows=1)
yuyue_data = pd.read_excel(path_jk,sheetname = '预约及时率（广西）',skiprows=1)
huifang_data = pd.read_excel(path_jk,sheetname = '1.4家宽装机回访（短信触点）',skiprows=1)
jihuo_rate1 = pd.read_excel(path_jk,sheetname = '1.5正装机一次成功率',skiprows=1) 
tuidanlv = pd.read_excel(path_jk,sheetname = '1.2退单率',skiprows=1) 
hdguifan = pd.read_excel(path_jk,sheetname = '1.6装机回单规范率',skiprows=1) 
hdzguifan = pd.read_excel(path_jk,sheetname = '1.10预约挂起规范率',skiprows=1) 

urban_g = jk_data.iloc[:15,:3].fillna('')
contryside_g = jk_data.iloc[:15,3:6].fillna('')
urban_p = jk_data.iloc[:15,6:9].fillna('')
contryside_p = jk_data.iloc[:15,9:12].fillna('')

challenge1 = 16
challenge2 = 28
challenge3 = 40
challenge4 = 64

col_name(contryside_g,'地市.1','装移机平均时长.1','排名.1')
col_name(urban_p,'地市.2','装移机平均时长.2','排名.2')
col_name(contryside_p,'地市.3','装移机平均时长.3','排名.3')
urban_g.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市

contexts = []
urban_g = urban_g.round({'装移机平均时长': 2})

duration1 = urban_g.loc['全区','装移机平均时长']
duration2 = contryside_g.loc['全区','装移机平均时长']
duration3 = urban_p.loc['全区','装移机平均时长']
duration4 = contryside_p.loc['全区','装移机平均时长']

index = ['南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']

jsrate_g = jk_data.iloc[:15,22:25].fillna('')  #获取高品质及时率
jsrate_p = jk_data.iloc[:15,25:28].fillna('')  #获取普通品质及时率
jsrate_g.rename(columns={"地市.7": '地市',"排名.7": '排名'},inplace=True)  # 改列名
jsrate_p.rename(columns={"地市.8": '地市',"装移机及时率.1": '装移机及时率',"排名.8": '排名'},  inplace=True)  # 改列名
jsrate_g.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
jsrate_p.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
g_jsrate1 = jsrate_g.loc['全区','装移机及时率']
p_jsrate1 = jsrate_p.loc['全区','装移机及时率']

#装移机全流程处理时长
shichang_qlc = qlc_data.iloc[:15,8:11]
shichang_qlc.rename(columns={"地市.2": '地市',"装移机全流程处理时长.1": '装移机全流程处理时长',"排名.2": '排名'},  inplace=True)  # 改列名
shichang_qlc.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
duration_qlc1 = shichang_qlc.loc['全区','装移机全流程处理时长']

#家宽装机预约及时率
yuyue_data = yuyue_data.iloc[:15,:6]
yuyue_data.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
yuyue_data['预约及时率'] = pd.to_numeric(yuyue_data['预约及时率'], errors='coerce')
yuyue_g = yuyue_data.loc['全区','高品质装机首次响应时长']
yuyue_p = yuyue_data.loc['全区','普通装机首次响应时长']
yuyue_rate = yuyue_data.loc['全区','预约及时率']


#家宽装机触点短信调研情况
huifang_data = huifang_data.iloc[:15,:26]
huifang_data.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
canping_user = huifang_data.iloc[-1]["评价量"]
canpinglv = huifang_data.iloc[-1]['参评率']
manyidu = huifang_data.iloc[-1]['家宽装机场景满意度=（安装结果*20%+上门及时性*30%+安装专业性*30%+安装人员服务*20%-1）/9*100']

# 装移机一次激活成功率
jihuo_rate = jihuo_rate1.iloc[:15,:10]
jihuo_rate.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
jihuolv = jihuo_rate.iloc[-1]['一次成功率']
jihuo_rate_old = jihuo_rate1.iloc[20:35,:10]
jihuo_rate_old.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
jihuolv_old = jihuo_rate_old.iloc[-1]['一次成功率']

#退单情况
new_tdl = tuidanlv.iloc[:15,32:63]
new_tdl.set_index('地市.1',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
old_tdl = tuidanlv.iloc[20:35,32:63]
zongzj = new_tdl.iloc[-1]['总装机单数量.1']
zongtd = new_tdl.iloc[-1]['退单数（城市）.1']+new_tdl.iloc[-1]['退单数（农村）.1']
ztd_rate =new_tdl.iloc[-1]['总退单率.1']
ztd_rate_old = old_tdl.iloc[-1]['总退单率.1']

def zhangjiang(new,old,fuhao):
    huanbi = (new-old)/old
    if huanbi >= 0:
        if fuhao != '':
            zhangsheng = '上涨'+str(round(huanbi * 100,2)) + fuhao
        else:
            zhangsheng = '上涨'+str(round(huanbi,2))
    else:
        if fuhao != '':
            zhangsheng = '下降'+str(round(huanbi * 100*-1,2)) + fuhao
        else:
            zhangsheng = '下降'+str(round(huanbi*-1,2))
    return zhangsheng
reason_user = new_tdl.iloc[-1]['退单率-用户原因.1']
reason_qt = new_tdl.iloc[-1]['退单率-前台原因.1']
reason_net = new_tdl.iloc[-1]['退单率-网络原因.1']
reason_jianshe = new_tdl.iloc[-1]['退单率-建设原因.1']
reason_net_old =old_tdl.iloc[-1]['退单率-网络原因.1']
#装移机回单规范率
hdguifan = hdguifan.iloc[:16,]
hdguifan_rate1 = hdguifan.iloc[1:16,21:]
hdguifan_rate1.rename(columns={'地市.1': '地市'}, inplace=True)  # 改列名
hdguifan_rate1.rename(columns={'合格率.1': '合格率'}, inplace=True)  # 改列名
hdguifan_rate1.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
hdguifan_rate = hdguifan_rate1.loc['全区']['合格率']

#缓装待装规范率
hdzguifan_rate = hdzguifan.iloc[:15,:6]
hdzguifan_rate.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为地市
hdzguifanlv = hdzguifan_rate.iloc[-1]['缓装待装规范率']

context = {"duration1": duration1, "content1": tiaozhan(urban_g,'装移机平均时长',challenge1,24,1),
           "duration2": duration2,"content2": tiaozhan(contryside_g,'装移机平均时长',challenge2,36,1),
           "duration3": duration3,"content3": tiaozhan(urban_p,'装移机平均时长',challenge3,48,1),
           "duration4": duration4,"content4": tiaozhan(contryside_p,'装移机平均时长',challenge4,72,1),
           "g_jsrate1": '%.2f%%' % (g_jsrate1 * 100),"p_jsrate1": '%.2f%%' % (p_jsrate1 * 100),
           "content5":tiaozhan(jsrate_g,'装移机及时率',0.96,0.93,-1),
           "content6":tiaozhan(jsrate_p,'装移机及时率',0.96,0.93,-1),
           "duration_qlc1":duration_qlc1,
           "content7":mubiao(shichang_qlc,"装移机全流程处理时长",48,1,'目标值'),
           "yuyue_g":round(yuyue_g,2),"yuyue_p":round(yuyue_p,2),"yuyue_rate":'%.2f%%' % (yuyue_rate * 100),
           "content8":mubiao(yuyue_data,"预约及时率",0.98,-1,'挑战值'),
           "canping_user":canping_user,
           "canpinglv":'%.2f%%' % (canpinglv * 100),"manyidu":manyidu,"content9":tiaozhan(huifang_data,'安装结果满意度（20%）',99,97,-1),
           "jihuolv":'%.2f%%' % (jihuolv * 100),"content10":mubiao(jihuo_rate,'一次成功率',0.7,-1,'目标值'),
           "zongzj":zongzj,"zongtd":zongtd,"ztd_rate":'%.2f%%' % (ztd_rate * 100),"zhangsheng":zhangjiang(ztd_rate,ztd_rate_old,'%'),"content11":tiaozhan(new_tdl,'总退单率.1',0.18,0.2,1),
           "reason_user":'%.2f%%' % (reason_user * 100),"reason_qt":'%.2f%%' % (reason_qt * 100),"reason_net":'%.2f%%' % (reason_net * 100),
           "reason_jianshe":'%.2f%%' % (reason_jianshe * 100),"hdguifan_rate":'%.2f%%' % (hdguifan_rate * 100),
           "content12":tiaozhan(hdguifan_rate1,'合格率',0.95,0.90,-1),"hdzguifanlv":'%.2f%%' % (hdzguifanlv * 100),
           "content13":tiaozhan(hdzguifan_rate,'缓装待装规范率',98,95,1),"date":date,"result1":tiaozhan1(duration1,16,'挑战值（16小时',-1),
           "result2":tiaozhan1(duration2,28,'挑战值（28小时',-1),
           "result3":tiaozhan1(duration3,40,'挑战值（40小时',-1),"result4":tiaozhan1(duration4,64,'挑战值（64小时',-1),
           "result5":tiaozhan1(g_jsrate1,0.96,'挑战值（96%',1),
           "result6":tiaozhan1(p_jsrate1,0.96,'挑战值（96%',1),"result7":tiaozhan1(duration_qlc1,48,'目标值（48小时',-1),
           "result8":tiaozhan1(yuyue_rate,0.98,'挑战值（98%',1),
           "result9":tiaozhan1(manyidu,99,'挑战值（99',1),"result10":tiaozhan1(jihuolv,0.7,'目标值（70%',1),
           "result11":tiaozhan1(reason_net,0.05,'目标值（5%',-1),
           "result12":tiaozhan1(hdguifan_rate,0.95,'挑战值（95%',1),"result13":tiaozhan1(hdzguifanlv,0.95,'挑战值（95%',1),
           "tiaozhan_td":tiaozhan1(ztd_rate,0.18,'挑战值（18%',-1)
          } #变量名称与Word文档中的占位符要一一对应
contexts.append(context)

for context in contexts:
    doc = DocxTemplate(r"家宽模板.docx")  #需要填入的Word文档的的地址
    doc.render(context)
    doc.save("1-3-1．2021年"+date+"家庭宽带装机管理工作报告.docx")



############################### 家客 #############################################
qz = pd.read_excel(path_dkzb,skiprows = 1 )
qz_new = qz.loc[:13,['指标','广西','南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']]
qz_new.set_index('指标',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为指标
qz_new = qz_new.T
qz_old = qz.loc[17:30,['指标','广西','南宁', '桂林', '柳州', '玉林', '百色', '河池', '贵港', '钦州', '梧州', '北海', '崇左', '来宾', '贺州', '防城港']]
qz_old.set_index('指标',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为指标
qz_old = qz_old.T
duration_qz = qz_new.loc['广西']['千兆装机时长']
duration_qzsx = qz_new.loc['广西']['千兆装机首次响应时长']
qz_ggl = pd.read_excel(path_dkzb,skiprows = 35)
qz_ggl.set_index('地市',drop=True, append=False, inplace=True, verify_integrity=False)  #修改索引为指标
num_bdb = qz_ggl.loc['广西']['不达标数量']
dblv_qz = qz_new.loc['广西']['千兆新装光功率达标率']
dblv_qz_old = qz_old.loc['广西']['千兆新装光功率达标率']

dblv_cs = qz_new.loc['广西']['千兆新装宽带测速达标率']
dblv_cs_old = qz_old.loc['广西']['千兆新装宽带测速达标率']
yadanbi = qz_new.loc['广西']['装机在途工单压单比']
zhijianlv = qz_new.loc['广西']['质检合格率']

def db(contrast,num,con):
    if contrast < num:
        content = con
    else:
        content = '未'+con
    return content
duration_qz_old = qz_old.loc['广西']['千兆装机时长']
duration_qzsx_old = qz_old.loc['广西']['千兆装机首次响应时长']
def db_city(table,content,mubiaozhi,word,qf):
    not_dabiao = []
    for i in index:
        if  table.loc[i,content] *qf > mubiaozhi*qf:
            not_dabiao.append(i)
    if not_dabiao != []:
        content = '、'.join(not_dabiao)+'市公司未达'+ word
    else:
        content = '各市公司均达标'
    return content
def db_all(table,content,qf,mubiaozhi):
    not_dabiao = []
    dabiao = []
    for i in index:
        if  table.loc[i,content]*qf > mubiaozhi*qf:
            not_dabiao.append(i)
        else:
            dabiao.append(i)
    if not_dabiao != []:
        content = '、'.join(dabiao)+'市公司已达标'+';'+'、'.join(not_dabiao)+'市公司未达标'
    else:
        content = '各市公司均达标'
    return content

def tiaozhan2(table,content,mubiaozhi,qf,word):
    not_mubiao = []
    for i in index:
        if  table.loc[i,content] *qf > mubiaozhi*qf:
            not_mubiao.append(i)
    if not_mubiao != []:
        content =  '、'.join(not_mubiao)+'未达挑战值'
    else:
        content = '各市公司均达挑战值'
    return content

contexts_jke = []
context_jke = {"duration_qz":round(duration_qz,2),"dabiao_gx1":db(duration_qz,8,'达到目标要求（8工作时）'),"huanbi_qz":zhangjiang(duration_qz,duration_qz_old,''),
               "dabiao_dishi1":db_city(qz_new,'千兆装机时长',8,'标',1),
               "duration_qzsx":round(duration_qzsx,2),"dabiao_dishi2":db_city(qz_new,'千兆装机首次响应时长',0.5,'标',1),"huanbi_jh":zhangjiang(jihuolv,jihuolv_old,'%'),
               "dabiao_gx2":db(duration_qzsx,0.5,'达到目标要求（0.5）'),"huanbi_qzsx":zhangjiang(duration_qzsx,duration_qzsx_old,''),
               "dabiao_dishi3":tiaozhan(jihuo_rate,'一次成功率',0.9,0.8,-1),"content_zj":mubiao(qz_new,'质检合格率',0.9,-1,"目标值"),
               "jihuolv":'%.2f%%' % (jihuolv * 100),"dabiao_gx3":tiaozhan1(jihuolv,0.9,'挑战值（90%',1),"dabiao_ggl":db_all(qz_new,'千兆新装光功率达标率',-1,1),
               "dabiao_cs":db_all(qz_new,'千兆新装宽带测速达标率',-1,1),"dabiao_ydb":db_all(qz_new,'装机在途工单压单比',1,3),
               "dabiao_td":db_all(new_tdl,'总退单率.1',1,0.2),"tiaozhan_yuyue":tiaozhan2(yuyue_data,"预约及时率",0.98,-1,'挑战值'),
               "num_bdb":num_bdb,"dblv_qz":'%.2f%%' % (dblv_qz * 100),"dabiao_gx4":tiaozhan1(dblv_qz,1,'达目标值（100%',1),
               "yuyue_g":round(yuyue_g,2),"yuyue_p":round(yuyue_p,2),"yuyue_rate":'%.2f%%' % (yuyue_rate * 100),
               "tiaozhan_tdan":tiaozhan1(ztd_rate,0.2,'目标值（20%',-1),
               "reason_user":'%.2f%%' % (reason_user * 100),"reason_qt":'%.2f%%' % (reason_qt * 100),"reason_net":'%.2f%%' % (reason_net * 100),
               "result11":tiaozhan1(reason_net,0.05,'目标值（5%',-1),"huanbi_ggl":zhangjiang(dblv_qz,dblv_qz_old,'%'),
               "huanbi_net":zhangjiang(reason_net,reason_net_old,'pp'),"tiaozhan_jishilv":tiaozhan1(yuyue_rate,0.98,'挑战值（98%',1),
               "dblv_cs":'%.2f%%' % (dblv_cs * 100), "dabiao_gx5":db(dblv_cs,1,'达到目标值（100%）'),"huanbi_cs":zhangjiang(dblv_cs,dblv_cs_old,'%'),
               "yadanbi":yadanbi,"dabiao_gx6":db(yadanbi,3,'达阶段目标值（3）'),"zhijianlv":'%.2f%%' % (zhijianlv * 100),"result_zj":db(zhijianlv,0.9,'目标值（90%）'),
               "reason_jianshe":'%.2f%%' % (reason_jianshe * 100),"ztd_rate":'%.2f%%' % (ztd_rate * 100),"zhangsheng":zhangjiang(ztd_rate,ztd_rate_old,'pp')}
contexts_jke.append(context_jke)

for context_jke in contexts_jke:
    doc_jke = DocxTemplate(r"家客模板.docx")  #需要填入的Word文档的的地址
    doc_jke.render(context_jke)
    doc_jke.save('关于2021年'+date+'家客专业运维管理工作情况通报.docx')



