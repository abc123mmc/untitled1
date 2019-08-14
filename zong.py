# -*- coding: utf-8 -*-
#过度模块，其中【quanXuan，xiuGaiBiaoQian】方法和【jb,zb】变量被 模块get_data 和 run引用
import time
import win32gui
import pyautogui#模拟鼠标键盘操作
import pyperclip#剪切板操作
import get_jbzb
import traceback#用于错误处理

def quanXuan():#全选，并返回复制内容
    import abc123,zhaotu
    for i in range(10):#此循环实现判断订单明细是否加载中
        if zhaotu.matchImg('./图片/订单是否存在.jpg',0.95):return False
        else:
            for i in [0,-5,10,-15,20,-25,30]:
                b=pyperclip.copy("")
                pyautogui.click(zb['订单全选'][0]+i,zb['订单全选'][1])
                time.sleep(0.2)
                abc123.kuaiJieJian(17,67)
                time.sleep(0.8)
                b=pyperclip.paste()
                print(len(b),207)
                if len(b)>50:return b


def xiuGaiBiaoQian(bqs01=6):#修改订单标签
    '''1未知，6测试'''
    time.sleep(1)
    pyautogui.click(*zb['批量修改'])
    pyautogui.click(*zb['批量修改标签'])
    pyautogui.press('space')
    pyautogui.press('down',bqs01-1)
    pyautogui.press('enter')
    pyautogui.press('tab')
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
    print('订单修改标签延时%s秒'%(time.time()-t1))


try:jb,zb=get_jbzb.jbzb_run()
except:
    now = int(time.time())
    timeArray = time.localtime(now)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S",timeArray)
    if traceback.format_exc()!='NoneType: None\n':
        f=open(r"执行日志.txt",'a')
        traceback.print_exc(file=f)
        f.write('*******************************'+otherStyleTime+'************************************\n\n')
        f.flush()  
        f.close()


if __name__=='__main__':
    xiuGaiBiaoQian(1)
    quanXuan()
    pass
