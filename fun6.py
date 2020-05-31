import win32ui

import time
import pyperclip#剪切板操作

import cv2
import numpy
import aircv

import traceback#用于错误处理

#win32api|win32con|win32gui|pyautogui#模拟鼠标键盘操作||GetHandle[zjb|zsjb|jbzb_run]
from get_jbzb import *


class WindowCapture:
    @staticmethod
    def window_capture(filename):#截屏，传入保存文件名，截屏并保存
        hwnd = 0 # 窗口的编号，0号表示当前活跃窗口
        hwnd_dc = win32gui.GetWindowDC(hwnd)# 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)# 根据窗口的DC获取mfcDC
        save_dc = mfc_dc.CreateCompatibleDC()# mfcDC创建可兼容的DC
        save_bit_map = win32ui.CreateBitmap()# 创建bigmap准备保存图片
        moniter_dev = win32api.EnumDisplayMonitors(None, None)# 获取监控器信息
        w = moniter_dev[0][2][2]
        h = moniter_dev[0][2][3]
        # print w,h　　　#图片大小
        save_bit_map.CreateCompatibleBitmap(mfc_dc, w, h)# 为bitmap开辟空间
        save_dc.SelectObject(save_bit_map)# 高度saveDC，将截图保存到saveBitmap中
        save_dc.BitBlt((0, 0), (w, h), mfc_dc, (0, 0), win32con.SRCCOPY)# 截取从左上角（0，0）长宽为（w，h）的图片
        save_bit_map.SaveBitmapFile(save_dc, filename)

    @staticmethod
    def match_img(imgobj,imgsrc,confidence=0.5,quzhi='result'):#传入小图、大图、相似度、取值，返回小图所在大图的坐标
        """取值分别为result【中心点坐标】，rectangle【四角坐标】，confidence【相似度】"""
        imobj = cv2.imdecode(numpy.fromfile(imgobj,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
        imsrc = cv2.imdecode(numpy.fromfile(imgsrc,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
        match_result = aircv.find_all_template(imsrc,imobj,confidence)
        if match_result:return [i[quzhi] for i in match_result]
        else:return False#;print('%s查找失败'% imgobj[:-4])

class KuaiJieJian:#快捷键
    def __init__(self,*arg):
        """Ctrl、A键的ASCII码分别为17和65 ,Ctrl+A快捷键:kuaiJieJian(17,65)"""
        """Ctrl、C键的ASCII码分别为17和67 ,Ctrl+A快捷键:kuaiJieJian(17,67)"""
        for i in arg:win32api.keybd_event(i,0,0,0)
        time.sleep(0.2)
        for i in arg[::-1]:win32api.keybd_event(i,0,win32con.KEYEVENTF_KEYUP,0)  #释放按键,arg[::-1]元组倒序

class TxtWrite():
    @staticmethod
    def write_in_txt(file,info):
        with open(file,'a') as f:
            f.write(info)

    @staticmethod
    def error(file):
        now = int(time.time())
        time_array = time.localtime(now)
        other_style_time = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
        # if traceback.format_exc()!='NoneType: None\n':
        with open(file, 'a') as f:
            traceback.print_exc(file=f)
            f.write(f'*******************************{other_style_time}************************************\n\n')

class Fun():
    @staticmethod
    def quan_xuan(jbzb,xiaot,datu):#全选，并返回复制内容
        zb=jbzb[1]
        pyautogui.click(zb['订单全选'][0]+50,zb['订单全选'][1]+20)
        pyautogui.press('home')
        for i in range(10):#此循环实现判断订单明细是否加载中
            if WindowCapture.match_img(xiaot,datu,0.95):continue
            else:
                for i in [0,-5,10,-15,20,-25,30]:
                    b=pyperclip.copy("")
                    pyautogui.click(zb['订单全选'][0]+i,zb['订单全选'][1])
                    time.sleep(0.2)
                    KuaiJieJian(17,67)  #复制
                    time.sleep(0.8)
                    b=pyperclip.paste()
                    print(len(b),207)
                    if len(b)>50:return b
            time.sleep(0.5)
        return False

    @staticmethod
    def cxdd(jbzb,ptdh01='',i=['',''],cx=False):#查询订单,参数:句柄|平台单号|付款开始和结束时间|是否为无条件查询
        zb=jbzb[1]
        jb=jbzb[0]
        if not cx:
            pyautogui.click(zb['日期1'])
            if i[0]:
                pyperclip.copy(i[0])
                pyautogui.hotkey('ctrl', 'v')
            pyautogui.click(zb['日期2'])
            if i[1]:
                pyperclip.copy(i[1])
                pyautogui.hotkey('ctrl', 'v')
            win32gui.SendMessage(jb['平台订单输入'], win32con.WM_SETTEXT, None,ptdh01)
            time.sleep(2)
        pyautogui.click(zb['查询'])
        time.sleep(1)

    @staticmethod
    def cxpanduan(jia_zz,quan_pt):#判断查询
        for i in range(60):#判断加载中是否存在达到判断订单是否加载成功
            time.sleep(0.2)
            if WindowCapture.match_img(jia_zz,quan_pt,0.95):time.sleep(1)
            else:return True
        if win32api.MessageBox(0,'订单查询失败\n点击确定退出应用\
        \n点击取消忽略错误继续', u'提示框',win32con.MB_OKCANCEL)==1:
            raise RuntimeError('testError')

    @staticmethod
    def xuanzeshijian():#时间段
        t1,li=time.time(),[1,4,8,12,24,72]
        lli=[]
        for i in range(len(li)-1):
            t2=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(t1-3600*li[i+1]))
            t3=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(t1-3600*li[i]))
            lli.append([t2,t3])
        return lli

    @staticmethod
    def _zhantie(val):##
        if val:
            KuaiJieJian(17,65)#全选
            pyautogui.press('del')#删除
            time.sleep(0.2)
            pyperclip.copy(val)#写入剪切板
            KuaiJieJian(17,86)#粘贴

    def gai_dz(self,jbzb,name=0, tel='', dz=0, sqx=()):  #改地址，传入name、tel、dz、sqx，修改收件消息为传入参数
        zb=jbzb[1]
        pyautogui.doubleClick(zb['第一条订单'])  # 双击击第一条订单
        pyautogui.click(zb['zb'])
        pyautogui.press('end')
        time.sleep(0.5)
        pan_tiao = 1
        for i in range(16):  #此循环实现判断是否订单明细页是否能加载成功
            time.sleep(0.5)
            b=pyperclip.copy("")
            pyautogui.click(zb['zb'])
            pyautogui.press('end')
            time.sleep(0.5)
            pyautogui.press('tab',3)
            KuaiJieJian(17,67)
            b=pyperclip.paste()
            if b:pan_tiao=0;break
        if pan_tiao:print('''gaidizh(name=0,tel=0,dz=0,sqx=0):#改地址**********跳出''');return
        pyautogui.press('tab',3)
        pyautogui.press('down')
        self._zhantie(name)##
        pyautogui.press('tab')
        self._zhantie(tel)##
        pyautogui.press('tab',3)
        def udt(u,d):#
            pyautogui.press('up',u)
            pyautogui.press('down',d)
            pyautogui.press('tab')
        if dz:
            udt(34,sqx[0])#
            udt(1,sqx[1])#
            udt(18,sqx[2])#
            self._zhantie(dz)##
        pyautogui.click(zb['保存'])#保存
        pan_tiao=1
        for i in range(16):#此循环实现判断是否保存成功
            time.sleep(0.5)
            b=pyperclip.copy("")
            KuaiJieJian(17,67)
            b=pyperclip.paste()
            if '系统信息' in b:pyautogui.press('enter')
            else:pan_tiao=0;break
        if pan_tiao:print('''gaidizh(name=0,tel=0,dz=0,sqx=0):#改地址111**********跳出''')  ;return
        pan_tiao=1
        for i in range(26):#此循环实现判断是否保存成功
            time.sleep(0.5)
            b=pyperclip.copy("")
            pyautogui.click(zb['zb'])
            pyautogui.press('end')
            time.sleep(0.5)
            pyautogui.press('tab',3)
            KuaiJieJian(17,67)
            b=pyperclip.paste()
            if b:pan_tiao=0;break
        for i in range(2):
            pyautogui.click(zb['关闭销售订单明细'][0]-30,zb['关闭销售订单明细'][1])
            pyautogui.click(zb['关闭销售订单明细'])
        time.sleep(1)
        b=GetHandle().zsjb(0,'销售订单明细')
        if b:
            time.sleep(10)#异常时等待10秒
            pyautogui.press('tab')
            pyautogui.press('enter')
            self.gai_dz(zb,name,tel,dz,sqx)

    @staticmethod
    def gai_sp(jbzb,hp1='DHJ',hp2='PKB'):
        """改货品:传入订单、货品1和2，把对应订单的货品1换成货品2"""
        zb=jbzb[1]
        pyautogui.click(zb['批量修改'])
        pyautogui.click(zb['批量修改商品'])
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
            b=GetHandle().zjb(0,c='WindowsForms10.Window.8.app.')
            b=[i for i in b if win32gui.IsWindowVisible(i)]
            if len(b)>1:
                for i in b:
                    if '错误' in win32gui.GetWindowText(i):
                        left, top, right, bottom =win32gui.GetWindowRect(i)
                        pyautogui.click(right-25,top+10)
            else:print('标签修改成功')  ;break
        print('修改赠品延时%s秒'%(time.time()-t1))

    @staticmethod
    def gai_kd(jbzb,kd):
        """改快递:传入订单、快递编码，把对应订单 快递进行修改"""
        zb=jbzb[1]
        pyautogui.click(zb['批量修改'])
        pyautogui.click(zb['批量修改快递'])
        xtxx=GetHandle.zjb(0,'WindowsForms10.Window.8.app.')
        for i in xtxx:
            if '系统信息' in win32gui.GetWindowText(i):xtxx=i
        if type(xtxx)==int:
            fou=GetHandle().zsjb(xtxx,'否')[0]
            left, top, right, bottom =win32gui.GetWindowRect(fou)
            pyautogui.click(right-10,top+10)
        pyautogui.press('space')
        pyautogui.press('down',kd-1)
        pyautogui.press('enter')
        pyautogui.press('tab')
        pyautogui.press('enter')
        t1=time.time()
        for i in range(50):
            b=GetHandle.zjb(0,'WindowsForms10.Window.8.')
            b=[i for i in b if win32gui.IsWindowVisible(i)]
            if len(b)>1:
                time.sleep(0.4)
                for i in b:
                    if '错误' in win32gui.GetWindowText(i):
                        left, top, right, bottom =win32gui.GetWindowRect(i)
                        pyautogui.click(right-25,top+10)
            else:break#print('快递修改成功')
        print('修改快递延时%s秒'%(time.time()-t1))

    @staticmethod
    def shenghe(jbzb):#审核
        zb=jbzb[1]
        pyautogui.click(zb['批量处理'])
        pyautogui.click(*zb['批量强制客审'])#批量强制客审
        time.sleep(2)
        xtxx=GetHandle.zjb(0,'WindowsForms10.Window.8.app.')
        for i in xtxx:
            if '系统信息' in win32gui.GetWindowText(i):xtxx=i
        if type(xtxx)==int:
            fou=GetHandle().zsjb(xtxx,'否')[0]
            left, top, right, bottom =win32gui.GetWindowRect(fou)
            pyautogui.click(right-10,top+10)
        for i in range(20):
            time.sleep(1)
            b=GetHandle().zjb(0,'WindowsForms10.Window.8.app.')
            b=[i for i in b if win32gui.IsWindowVisible(i)]
            if len(b)>1:
                for i in b:
                    if '错误' in win32gui.GetWindowText(i):
                        left, top, right, bottom =win32gui.GetWindowRect(i)
                        pyautogui.click(right-25,top+10)
            else:print('审核完成')  ;del b;break

    @staticmethod
    def gai_bq(jbzb,bqs01=6):  # 修改订单标签
        """1未知，6测试"""
        zb=jbzb[1]
        time.sleep(1)
        pyautogui.click(*zb['批量修改'])
        pyautogui.click(*zb['批量修改标签'])
        pyautogui.press('space')
        pyautogui.press('down', bqs01 - 1)
        pyautogui.press('enter')
        pyautogui.press('tab')
        pyautogui.press('enter')
        t1 = time.time()
        for i in range(20):
            time.sleep(1)
            b = GetHandle.zjb(0,'WindowsForms10.Window.8.app.')
            b = [i for i in b if win32gui.IsWindowVisible(i)]
            if len(b) > 1:
                for i in b:
                    if '错误' in win32gui.GetWindowText(i):
                        left, top, right, bottom = win32gui.GetWindowRect(i)
                        pyautogui.click(right - 25, top + 10)
            else:
                print('标签修改成功');break
        print('订单修改标签延时%s秒' % (time.time() - t1))

if __name__ == '__main__':
      pass

