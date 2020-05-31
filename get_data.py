# -*- coding: utf-8 -*-
import re,xlrd,random

fnn6_nr=['Fun', 'GetHandle', 'KuaiJieJian', 'TxtWrite', 'WindowCapture',
         'aircv', 'cv2', 'numpy', 'pyautogui', 'pyperclip', 'time',
         'traceback', 'win32api', 'win32con', 'win32gui', 'win32ui']
del fnn6_nr
from fun6 import *
import get_config
import setting

zb=GetHandle().jbzb_run()
WindowCapture.window_capture(setting.quan_pt)  # 截屏


#******************************************取数据***********************************************
class GetData():
    @staticmethod #返回函数的静态方法,可以实现实例化使用,也可不实例化调用 C.f()
    def di_zhi_chu_li(beizhu):
        """地址处理:传入存在改地址的备注，返回address"""
        beizhu=beizhu.replace('\n','')#去除备注中的换行
        dz01=re.findall('改地址[:：](.*?)┋',beizhu)
        d=''
        if dz01:
            d=''.join(e for e in dz01[-1] if e.isalnum() or e in ':：;；._-')#去除备注里的符号
            c='(?:联系方式|电话|手机).{,5}?(?=1\d{10})|(?:所在地区|详细地址|收货人|姓名|地址)[:：]'
            d=re.sub(c,'',d)    #替换掉 联系方式|电话|手机|所在地区|详细地址|收货人|姓名|地址等字符
        if d:d=re.split('[;；]',d)
        if len(d)>3:return False,False
        dh=[]
        xm=[]
        dz=[]
        for i in d:
            cd=len(i)   #长度
            sz=len(re.findall('\d',i))  #数字个数
            sj=len(re.findall('省|市|区|县|乡|镇|村|路|号|街|道',i))   #包含地址字符个数
            if sz>9:
                if cd-sz<6:dh+=[i]
                else:dz+=[i]
            else:
                if sj>2:dz+=[i]
                elif sj==0:
                    if cd<8:xm+=[i]
                    else:dz+=[i]
                else:
                    if cd>10:dz+=[i]
                    else:xm+=[i]
        for i in [dh,dz,xm]:
            try:
                while 1:i.remove('')
            except:pass

        num1 = 0
        if len(dh)==0:dh=''
        elif len(dh)==1:dh=dh[0]
        else:
            dh1=''
            for i in dh:
                if not dh1:
                    num1=len(i)
                    dh1=i
                elif num1>len(i):
                    num1=len(i)
                    dh1=i
            dh=dh1
        if len(xm)==0:xm=''
        elif len(xm)==1:xm=xm[0]
        else:
            xm1=''
            for i in xm:
                if not xm1:
                    num1=len(i)
                    xm1=i
                elif num1>len(i):
                    num1=len(i)
                    xm1=i
            xm=xm1
        if len(dz)==0:dz=''
        elif len(dz)==1:dz=dz[0]
        else:
            dz1=''
            for i in dz:
                if not dz1:
                    num1=len(i)
                    dz1=i
                elif num1<len(i):
                    num1=len(i)
                    dz1=i
            dz=dz1
        sqx=''
        if dz:
            t1= xlrd.open_workbook(setting.diqu_e).sheets()[2]
            li=[]
            dz1=''.join(re.findall('.*(?<=[区县市旗岛域辖镇乡台阁仔])',dz))
            if not dz1:dz1=dz
            for i in range(t1.nrows):#循环行
                h=t1.row_values(i)
                zhi=0;zhi2=0
                for i1 in h[3:6]:#分析当前行的456格
                    if i1 in dz1:zhi2+=5
                    elif i1[:2] in dz1:zhi2+=3
                    for i2 in range(len(i1)):#i1地址单元格
                        if i2<2 and i1[i2] in dz1:zhi+=2
                        elif i1[i2] in dz1:zhi+=1
                if zhi2<6:continue#权重不足时跳出当次循环
                zhi=zhi+zhi2
                if zhi>8:li.append([i,zhi])#i行，zhi权重
            zhi=[i[1] for i in li]#值的列表
            zhi1=sorted(zhi)[-2:]
            if not zhi1:print('改地址备注权重不足以判断地区：备注内容非地址',dz);return False,False
            elif len(zhi1)==1 or zhi1[0]!=zhi1[1]:
                if zhi1[-1]>10:sqx=t1.row_values(li[zhi.index(zhi1[-1])][0])[:5]
                else:print('改地址备注权重不足以判断地区：备注内容不全',dz);return False,False
            else:
                if zhi1[0]>10:
                    g5=t1.row_values(li[zhi.index(zhi1[-1])][0])[4]
                    #sqx=[t1.row_values(i)[:3] for i in range(t1.nrows) if (t1.row_values(i)[4]==g5 and '其它区' in t1.row_values(i)[5])][0]
                    try:sqx=[t1.row_values(i)[:5] for i in range(t1.nrows)
                             if (t1.row_values(i)[4]==g5 and '其它区' in t1.row_values(i)[5])][0]
                    except:print('其他区不存在1111111111111111',dz);return False,False
                else:print('修改地址备注异常222222222222',dz);return False,False
        shengshi=[]
        sqx1=[]
        if sqx:
            sqx1= [int(x) for x in sqx[:3]]
            shengshi= [x for x in sqx[3:]]
        address=(xm,dh,dz,sqx1)#姓名、电话、地址、地区编码的列表,省份城市列表
        for i in address:
            if i:return address,shengshi
        #print('电话',dh);print('姓名',xm);print('地址',dz);print('编码',sqx);print('省份城市',shengshi)
        return False,False

    @staticmethod
    def ddsj(jbzb,ddclzx_d,ptdh=False):
        """订单数据：返回当前查询到的所有订单数据,参数[句柄|需要返回的订单字段|总订单字段数]"""
        len1=len(ddclzx_d)   #订单页字段数
        ddclzx_d = {i: ddclzx_d[i] for i in ddclzx_d if i in setting.ddzds}    #需要的订单字段及位置的字典
        b=Fun.quan_xuan(jbzb,setting.jia_zz,setting.quan_pt)#实现判断订单明细是否能加载成功或者存在
        if not b:return {}
        #zong.xiuGaiBiaoQian()#设置标签为测试
        b=[i.split('\t') for i in b.split('\r\n校验')]#关注前11个--店铺名\仓库名\省\市\区\地址\客服备注\客户备注\快递\异常原因\平台单号
        print(b[0])
        print(ddclzx_d)
        b=[i for i in b if (i[int(ddclzx_d['订单标签'])]=='' or i[int(ddclzx_d['订单标签'])]=='测试')
           and len(i)-1==len1 and i[int(ddclzx_d['异常原因'])]=='' and i[int(ddclzx_d['订单类型'])]=='原单']
        b={i[int(ddclzx_d['平台订单'])]:{j:i[int(ddclzx_d[j])] for j in ddclzx_d if j!='平台订单'} for i in b}
        if ptdh:return [i for i in b]
        else:return b

    @staticmethod
    def spsj(jbzb, spxx_d):
        """商品数据：返回当前查询到的所有商品数据,参数[句柄|需要返回的商品字段|商品总字段数]"""
        zb=jbzb[1]
        len2=len(spxx_d)   #商品页字段数
        spxx_d = {i: spxx_d[i] for i in spxx_d if i in setting.spzds}    #需要的商品字段及位置的字典
        b=Fun.quan_xuan(jbzb,setting.jia_zz,setting.quan_pt)#实现判断订单明细是否能加载成功或者存在
        if not b:return {}
        sp=dict()
        f=''
        for j in range(200):
            pyperclip.copy("")
            for i in range(36):#此循环实现判断订单货品明细是否能加载成功
                pyautogui.click(zb['商品信息'])#商品信息
                time.sleep(0.1)
                pyautogui.click(zb['商品全选'])#全选商品
                KuaiJieJian(17,67)
                time.sleep(0.1)
                c=pyperclip.paste()
                if re.findall('平台商品|普通商品',c):
                    c=[i.split('\t') for i in c.split('\r\n校验')]#商品编码\品名\数量\规格\商品类型
                    c = [{j:i[int(spxx_d[j])] for j in spxx_d} for i in c if len(i)-1==len2]
                    break
            pyautogui.click(zb['订单信息'])#订单信息
            pyautogui.press('tab',3)
            KuaiJieJian(17,67)
            f001 = pyperclip.paste()
            # 不等于上次复制的订单编号,字符为纯数字,长度大于13
            if f!=pyperclip.paste() and 23>len(f001)>13:
                f=f001
                sp[f]={'订单货品':c}
            else:
                break
            pyautogui.click(zb['订单列表'])
            time.sleep(0.1)
            pyautogui.press('down')
            for i in range(666):#此循环实现判断订单明细是否加载中
                time.sleep(0.2)
                if WindowCapture.match_img(setting.jia_zz,setting.quan_pt,0.95):continue
                else:break
        return sp

    @staticmethod
    def ddspsj(dd,sp ):
        """返回订单商品数据,参数:[订单数据|商品数据]"""
        b={i:{**dd[i],**sp[i]} for i in sp if i in dd}
        return b

    @staticmethod
    def panduanpeizhi(v1,v2):#判段配置，传入配置值和订单对应字段值，符合配置返回True,不符合返回False
        #if v1==None:return True
        if isinstance(v1,str) and re.findall('\d+<%s<\d+',v1):return eval(v1% v2)
        elif isinstance(v1,str):
            tj=re.findall('^\[包含\]|^\[不包含\]|^\[空\]|^\[非空\]|^\[等于\]|^\[不等于\]',v1[:5])[0]
            v1=v1.replace(tj,'')
            v1=v1.split('|')
            if tj=='[包含]':
                if [i for i in v1 if i in v2]:return True
                else:return False
            elif tj=='[不包含]':
                if [i for i in v1 if i in v2]:return False
                else:return True
            if tj=='[等于]':
                if v1[0]==v2:return True
                else:return False
            elif tj=='[不等于]':
                if v1[0]==v2:return False
                else:return True
            if tj=='[空]':
                if v2:return False
                else:return True
            elif tj=='[非空]':
                if v2:return True
                else:return False

    def panduan(self):
        """传入订单货品数据和配置内容，返回zdsh（自动修改）、zdsh1（自动审核平台单号）、人工处理（平台单号）"""
        conf = get_config.GetCofig.read_configure(setting.peizhi_e)  # 配置文件内容
        ddclzx_dict= get_config.GetCofig.get_layout(setting.ddclzx)  #返回所有订单字及段位置的字典
        spxx_dict = get_config.GetCofig.get_layout(setting.spxx)  #返回商品信息页所有字段位置的字典
        Fun.cxdd(zb,cx=True)
        Fun.cxpanduan(setting.jia_zz,setting.quan_pt)
        sp = GetData.spsj(zb, spxx_dict)
        dd = GetData.ddsj(zb, ddclzx_dict)
        ddspsj01 = GetData.ddspsj(dd,sp)
        print(ddspsj01)

        #rgcl,gkd,gzp,k=Config
        zdsh=dict()
        zdsh['改快递']={conf[3][i]:[] for i in conf[3]}
        zdsh['改赠品'],zdsh['关闭订单'],zdsh['改地址']=[],[],{}
        zdsh1=[]
        reng_gong=[]
        print('将要分析的订单数是：%s'%len(ddspsj01))
        for i in ddspsj01:#i订单编号
            if ddspsj01[i]['客服备注'] or ddspsj01[i]['订单买家留言']:
                TxtWrite.write_in_txt(setting.cfrz,f'####:订单{i}存在备注不需要处理')
                continue #有备注不处理
            dianpu001=re.findall('SEPTWOLVES雅赋专卖店|少年狼箱包|拼多多美之瑞专卖店',ddspsj01[i]['网店名称'])
            if dianpu001:
                TxtWrite.write_in_txt(setting.cfrz,f'人工审核:订单{i}属于不需要处理的店铺{dianpu001}\n')
                continue#这几个店铺不处理
            try:
                spbt,guige,pingming,shuliang='','','',0
                for j1 in ddspsj01[i]['订单货品']:
                    spbt+=j1['商品标题']
                    guige+=j1['平台规格']
                    pingming+=j1['品名']
                    if '赠品' in j1['商品类型']:continue
                    shuliang+=int(j1['数量'])     #数量不计算赠品的
                ddspsj01[i]['商品标题'] = spbt
                ddspsj01[i]['平台规格']=guige
                ddspsj01[i]['品名']=pingming
                ddspsj01[i]['数量']=shuliang
                panduan_rg=False
                #筛选改地址
                dz01=re.findall('改地址[:：](.*?)┋',ddspsj01[i]['客服备注'])
                if dz01:
                    dz01,shengshi=self.di_zhi_chu_li(ddspsj01[i]['客服备注'])
                    if dz01:
                        zdsh['改地址'][i]=dz01#符合条件，修改备注
                        if shengshi:
                            ddspsj01[i]['省份']=shengshi[0]
                            ddspsj01[i]['城市']=shengshi[1]
                    else:
                        TxtWrite.write_in_txt(setting.cfrz,f'人工审核:订单{i}改地址识别原因需要人工审核\n')
                        panduan_rg=True#判断人工处理1

                #判断人工处理2
                r=ddspsj01[i]['客服备注'].split('┋')
                if len([i7 for i7 in r if  not re.findall('快递[:：]|长度[:：]|改地址[:：]',i7)])>1:
                    panduan_rg=True
                    TxtWrite.write_in_txt(setting.cfrz,f'人工审核:订单{i}备注含有需要人工处理的内容\n')

                #判断人工处理3
                for i11 in conf[0]:#i1一行
                    i1={i:i11[i] for i in i11 if not re.findall('处理方式',i)}
                    pan1=1
                    for i4 in i1:
                        if not self.panduanpeizhi(i1[i4],ddspsj01[i][i4]):#人工筛选
                            pan1=0
                            break
                    if pan1:
                        panduan_rg=True
                        TxtWrite.write_in_txt(setting.cfrz, f'人工审核:订单{i}符合人工处理配置{ddspsj01[i][i4]}\n')
                        break

                #筛选改快递
                for i11 in conf[1]:#i1一行
                    i1={i:i11[i] for i in i11 if not re.findall('快递公司',i)}#当前选择条件（一行）
                    pan1=1
                    for i4 in i1:
                        if not self.panduanpeizhi(i1[i4],ddspsj01[i][i4]):#快递筛选
                            pan1=0
                            break
                    if pan1:
                        if '客服备注指定' not in i11['快递公司']:
                            kd=i11['快递公司'].split('|')
                            kd=random.choice(kd)
                        else:
                            kd=re.findall('快递[:：](.*?)┋',ddspsj01[i]['客服备注'])[-1]#中通
                            kd=[i for i in conf[3] if kd in i]#['中通快递拉杆箱', '中通速递']
                            if len(kd)!=1:
                                panduan_rg=True
                                TxtWrite.write_in_txt(setting.cfrz, f'人工审核:订单{i}备注中内容不明确,需要人工审核\n')
                                break
                            else:kd=kd[0]
                        if kd not in ddspsj01[i]['快递公司']:
                            zdsh['改快递'][conf[3][kd]].append(i)
                            TxtWrite.write_in_txt(setting.cfrz,'改快递:订单%s符合条件%s改快递为[%s]\n'%(i,i1,kd))
                        break

                #筛选改赠品
                for i11 in conf[2]:#i1一行
                    i1={i:i11[i] for i in i11 if not re.findall('处理方式|快递公司',i)}
                    clfs=i11['处理方式']
                    try:
                        sykd01=i11['快递公司']
                    except:sykd01=''
                    pan1=1
                    for i4 in i1:
                        if not self.panduanpeizhi(i1[i4],ddspsj01[i][i4]):
                            pan1=0
                            break
                    if pan1:########################
                        if sykd01=='' or sykd01==kd:
                            zdsh['改赠品'].append(i)
                            TxtWrite.write_in_txt(setting.cfrz,'改赠品:订单%s符合条件%s改打火机为卡包\n'%(i,i1))
                            break#向后移动了一格，避免了i1为空的情况
            except:
                print('*************:',(i1[i4],ddspsj01[i][i4]))
                panduan_rg=True
                TxtWrite.write_in_txt(setting.cfrz,'人工审核:订单%s分析错误需要人工处理！！\n'% i)
                TxtWrite.error(setting.error_info)
                print(f'订单{i}错误需要人工处理！！\n',ddspsj01[i],'\n')
            if panduan_rg:reng_gong.append(i)
            else:
                TxtWrite.write_in_txt(setting.cfrz, '自动审核:订单%s分析可以自动审核！！\n' % i)
                zdsh1.append(i)
        return zdsh,zdsh1,reng_gong#返回需要修改的订单、所有可以自动审核订单、人工处理订单

    def gai_kd(self): #改快递
        pass

    def reng_gong_sh(self): #人工审核
        pass

    def gai_zp(self): #改赠品
        pass


#******************************************取数据***********************************************


if __name__ == '__main__':
    #ddclzx_dict1, len11 = get_config.GetCofig.get_layout(setting.ddclzx, setting.ddzds)  # 订单处理中心,字段位置的字典
    #spxx_dict1, len21 = get_config.GetCofig.get_layout(setting.spxx, setting.spzds)  # 商品信息 ,字段位置的字典
    #conf1 = get_config.GetCofig.read_configure(setting.peizhi_e)  # 配置文件内容
    #dd1 = GetData.ddsj(zb, ddclzx_dict1,len11)
    #sp1 = GetData.spsj(zb, spxx_dict1,len21)
    #ddspsj011 = GetData.ddspsj(dd1,sp1)
    b001=GetData().panduan()
    pass



