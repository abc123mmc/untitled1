# -*- coding: utf-8 -*-
import re,time
import win32api,win32con,win32gui
import pyautogui#模拟鼠标键盘操作
import pyperclip#剪切板操作
import abc123,zhaotu
import get_data,zong


try:get_jbzb=zong.get_jbzb
except:get_data.Ri_zhi()
#******************************************订单操作***********************************************
def gaidizh(name=0,tel=0,dz=0,sqx=0):#改地址，传入name、tel、dz、sqx，修改收件消息为传入参数
    pyautogui.doubleClick(*zong.zb['第一条订单'])#双击击第一条订单
    pyautogui.click(*zong.zb['zb'])
    pyautogui.press('end')
    time.sleep(0.5)
    pan_tiao=1
    for i in range(16):#此循环实现判断是否订单明细页是否能加载成功
        time.sleep(0.5)
        b=pyperclip.copy("")
        pyautogui.click(*zong.zb['zb'])
        pyautogui.press('end')
        time.sleep(0.5)
        pyautogui.press('tab',3)
        abc123.kuaiJieJian(17,67)
        b=pyperclip.paste()
        if b:pan_tiao=0;break
    if pan_tiao:print('''gaidizh(name=0,tel=0,dz=0,sqx=0):#改地址**********跳出''');return
    pyautogui.press('tab',3)
    pyautogui.press('down')
    def zhantie(name):##
        if name:
            abc123.kuaiJieJian(17,65)
            pyautogui.press('del')
            time.sleep(0.2)
            pyperclip.copy(name)
            abc123.kuaiJieJian(17,86)
    zhantie(name)##
    pyautogui.press('tab')
    zhantie(tel)##
    pyautogui.press('tab',3)
    def udt(u,d):#
        pyautogui.press('up',u)
        pyautogui.press('down',d)
        pyautogui.press('tab')
    if dz:
        udt(34,sqx[0])#
        udt(1,sqx[1])#
        udt(18,sqx[2])#
        zhantie(dz)##
    pyautogui.click(*zong.zb['保存'])#保存
    pan_tiao=1
    for i in range(16):#此循环实现判断是否保存成功
        time.sleep(0.5)
        b=pyperclip.copy("")
        abc123.kuaiJieJian(17,67)
        b=pyperclip.paste()
        if '系统信息' in b:pyautogui.press('enter')
        else:pan_tiao=0;break
    if pan_tiao:print('''gaidizh(name=0,tel=0,dz=0,sqx=0):#改地址111**********跳出''')  ;return  
    pan_tiao=1
    for i in range(26):#此循环实现判断是否保存成功
        time.sleep(0.5)
        b=pyperclip.copy("")
        pyautogui.click(*zong.zb['zb'])
        pyautogui.press('end')
        time.sleep(0.5)
        pyautogui.press('tab',3)
        abc123.kuaiJieJian(17,67)
        b=pyperclip.paste()
        if b:pan_tiao=0;break
    for i in range(2):
        pyautogui.click(zong.zb['关闭销售订单明细'][0]-30,zong.zb['关闭销售订单明细'][1])
        pyautogui.click(*zong.zb['关闭销售订单明细'])
    time.sleep(1)
    b=get_jbzb.zsjb(0,'销售订单明细')
    if b:
        time.sleep(10)#异常时等待10秒
        pyautogui.press('tab')
        pyautogui.press('enter')
        gaidizh(name,tel,dz,sqx)

def shenghe():#审核
    pyautogui.click(*zong.zb['批量处理'])
    pyautogui.click(*zong.zb['批量强制客审'])#批量强制客审
    time.sleep(2)
    xtxx=get_jbzb.zjb(0,'WindowsForms10.Window.8.app.')
    for i in xtxx:
        if '系统信息' in win32gui.GetWindowText(i):xtxx=i
    if type(xtxx)==int:
        fou=zsjb(xtxx,'否')[0]
        left, top, right, bottom =win32gui.GetWindowRect(w)
        pyautogui.click(right-10,top+10)
    for i in range(20):
        time.sleep(1)
        b=get_jbzb.zjb(0,c='WindowsForms10.Window.8.app.')
        b=[i for i in b if win32gui.IsWindowVisible(i)]
        if len(b)>1:
            for i in b:
                if '错误' in win32gui.GetWindowText(i):
                    left, top, right, bottom =win32gui.GetWindowRect(i)
                    pyautogui.click(right-25,top+10)
        else:print('审核完成')  ;del b;break

def gaiShangPing(hp1='DHJ',hp2='PKB'):
    '''改货品:传入订单、货品1和2，把对应订单的货品1换成货品2'''
    pyautogui.click(*zong.zb['批量修改'])
    pyautogui.click(*zong.zb['批量修改商品'])
    pyperclip.copy(hp1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyperclip.copy(hp2)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab',2)
    pyautogui.press('enter')
    t1=time.time()
    for i in range(20):
        time.sleep(1)
        b=get_jbzb.zjb(0,c='WindowsForms10.Window.8.app.')
        b=[i for i in b if win32gui.IsWindowVisible(i)]
        if len(b)>1:
            for i in b:
                if '错误' in win32gui.GetWindowText(i):
                    left, top, right, bottom =win32gui.GetWindowRect(i)
                    pyautogui.click(right-25,top+10)
        else:print('标签修改成功')  ;break
    print('修改赠品延时%s秒'%(time.time()-t1))

def gaiKuaiDi(kd):
    '''改快递:传入订单、快递编码，把对应订单 快递进行修改'''
    pyautogui.click(*zong.zb['批量修改'])
    pyautogui.click(*zong.zb['批量修改快递'])
    xtxx=get_jbzb.zjb(0,c='WindowsForms10.Window.8.app.')
    for i in xtxx:
        if '系统信息' in win32gui.GetWindowText(i):xtxx=i
    if type(xtxx)==int:
        fou=zsjb(xtxx,'否')[0]
        left, top, right, bottom =win32gui.GetWindowRect(w)
        pyautogui.click(right-10,top+10)
    pyautogui.press('space')
    pyautogui.press('down',kd-1)
    pyautogui.press('enter')
    pyautogui.press('tab')
    pyautogui.press('enter')
    t1=time.time()
    for i in range(50):
        b=get_jbzb.zjb(0,c='WindowsForms10.Window.8.')
        b=[i for i in b if win32gui.IsWindowVisible(i)]
        if len(b)>1:
            time.sleep(0.4)
            for i in b:
                if '错误' in win32gui.GetWindowText(i):
                    left, top, right, bottom =win32gui.GetWindowRect(i)
                    pyautogui.click(right-25,top+10)
        else:break#print('快递修改成功')
    print('修改快递延时%s秒'%(time.time()-t1))

def cxdd(ptdh01='',i=['','']):#查询订单：传入付款开始和结束时间,查询订单（1）
    pyautogui.click(*zong.zb['日期1'])
    if i[0]:
        pyperclip.copy(i[0])
        pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(*zong.zb['日期2'])
    if i[1]:
        pyperclip.copy(i[1])
        pyautogui.hotkey('ctrl', 'v')
    win32gui.SendMessage(zong.jb['平台订单输入'], win32con.WM_SETTEXT, None,ptdh01)
    time.sleep(1)
    pyautogui.click(*zong.zb['查询'])
    time.sleep(2)
    cxpanduan()#查询判断
    e=zong.quanXuan()
    return e

def cxpanduan():#判断查询
    for i in range(60):#判断加载中是否存在达到判断订单是否加载成功
        time.sleep(0.2)
        if zhaotu.matchImg('./图片/加载中.jpg',0.95):time.sleep(1)
        else:return True
    if win32api.MessageBox(0,'订单查询失败\n点击确定退出应用\
    \n点击取消忽略错误继续', u'提示框',win32con.MB_OKCANCEL)==1:
        raise RuntimeError('testError')

#******************************************订单操作***********************************************

def get_danhao():
    '''选择一小时前的'''
    t=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()-3600))
    b=cxdd(i=['',''])
    b=[i.split('\t') for i in b.split('\r\n校验')]
    c=[]
    for i in b:
        try:
            if re.findall('\d{10}',i[11]):#验证订单编号是否正确
                if i[12]=='' or i[12]=='测试':c.append(i[11])
        except:print(i,'订单分割存不成功')
    b=[]
    for i in range(0,len(c),100):
        b.append(c[i:i+100]) 
    return b

def go_go():
    e=get_danhao()
    for i in e:
        ptdh01=','.join(i)
        b=cxdd(ptdh01)
        b=[i.split('\t') for i in b.split('\r\n校验')]
        c=[]
        for i in b:
            try:
                if re.findall('\d{10}',i[11]):#验证订单编号是否正确
                    if (i[12]=='' or i[12]=='测试') and i[10]=='':c.append(i[11])
            except:print(i,'订单分割存不成功')
        
        ptdh01=','.join(c)  
        if ptdh01 and cxdd(ptdh01):
            hp=get_data.ddsp()#获取订单商品数据
            Config=get_data.readConfigure()#获取配置内容数据
            b,c,d=get_data.panduan(hp,Config)
            time.sleep(1)
            
            for i in b['改快递']:
                ptdh01=','.join(b['改快递'][i])
                if b['改快递'][i]:
                    e=cxdd(ptdh01)
                    if e:
                        gaiKuaiDi(i)
                        e=cxdd('改快递%s已完成'% i)
            for i in b['改地址']:
                e=cxdd(i)
                if not e:continue
                gaidizh(*b['改地址'][i])
                get_jbzb.xujiludnr('修改地址',str(i)+',')

            ptdh01=','.join(b['改赠品'])
            if ptdh01:
                e=cxdd(ptdh01)
                if e:
                    gaiShangPing()#修改商品
            ptdh01=','.join(d)
            if ptdh01:
                e=cxdd(ptdh01)
                if e:
                    zong.xiuGaiBiaoQian(0)#修改订单标签为1未知
                    e=cxdd('人工处理已更改订单标签为未知')
            ptdh01=','.join(c)
            if ptdh01:
                e=cxdd(ptdh01)
                e=[i.split('\t') for i in e.split('\r\n校验')]#关注前11个--店铺\仓库\省\市\区\地址\客服备注\客户备注\快递\异常原因\平台单号
                e1={}
                for i in e:
                    try:
                        e1[i[11]]=i[7]#客服备注
                    except:print(i,'订单分割存在不成功')
                e=[]#订单备注未被修改的订单
                for i in e1:
                    if e1[i]==hp[i]['客服备注']:e+=[i]
                    else:get_jbzb.xujiludnr('备注变更订单',str(i)+',')

                ptdh01=','.join(e)
                if ptdh01:
                    e=cxdd(ptdh01)
                if e:
                    shenghe()
                    e=cxdd('当次获取数据审核完成')
            time.sleep(3)
    e=cxdd('所有获取到的订单审核完成*******')


if __name__ == '__main__':
    try:go_go()
    except:get_data.Ri_zhi()
    pass


