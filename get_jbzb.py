import win32api,win32con,win32gui
import pyautogui#模拟鼠标键盘操作


class GetHandle():
    @staticmethod
    def zjb(hwnd,btlx=None):
        """ 传入父窗口的句柄，和要查找的标题或者类型（默认为None）,寻找父句柄下的所有子句柄 """
        handle = win32gui.FindWindowEx(hwnd,0,None,None)
        handlelist=[]
        while handle>0:
            if btlx:
                if btlx in win32gui.GetWindowText(handle) or btlx in win32gui.GetClassName(handle):
                    handlelist.append(handle)
            else:handlelist.append(handle)
            handle = win32gui.FindWindowEx(hwnd,handle,None,None)
        return handlelist

    def _find_idx_sub_handle(self,hwnd,btlx,handlelist=[]):
        handle = win32gui.FindWindowEx(hwnd, 0, None, None)
        while handle>0:
            pan_xuanze=0
            if btlx:
                if btlx in win32gui.GetWindowText(handle) or btlx in win32gui.GetClassName(handle):pan_xuanze=1
            else:pan_xuanze=1
            if pan_xuanze:handlelist.append(handle)
            self._find_idx_sub_handle(handle,btlx,handlelist)
            handle = win32gui.FindWindowEx(hwnd, handle, None, None)

    def zsjb(self,hwnd,btlx=None):
        """ 传入父窗口的句柄,和要查找的标题或者类型（默认为None），寻找父句柄下的所有子孙句柄 """
        handlelist=[]
        self._find_idx_sub_handle(hwnd,btlx,handlelist)
        return handlelist

    def jbzb_run(self):
       jb={}
       handle=self.zsjb(0,'维富友')
       if len(handle)!=1:
           win32api.MessageBox(0,'本机打开的维富友软件不唯一，请检查是否打开或者多打开', u'提示框',win32con.MB_SYSTEMMODAL)
           raise RuntimeError('testError')
       handle=self.zsjb(handle[0],'订单处理中心')[0]#订单处理中心
       handle=self.zjb(handle)[0]#订单处理中心下一级句柄
       shaixuan=self.zsjb(self.zjb(handle)[4])
       zb={i[0]:win32gui.GetWindowRect(shaixuan[i[1]])[0:2] for i in [['日期1',57],['日期2',55],['查询',51]]}
       jb['平台订单输入']=shaixuan[72]
       left,top,right,bottom=win32gui.GetWindowRect(self.zjb(handle)[-1])#售中处理、退货两行
       zb['批量处理']=(39*7,bottom/4+top*3/4)
       zb['保存']=(39,top+10)#销售订单明细页
       zb['关闭销售订单明细']=(330,top-15)#销售订单明细页
       zb['批量强制客审']=(136,bottom*3/4+top/4);zb['批量终止']=(580,bottom*3/4+top/4)
       zb['批量修改']=(39*9,bottom/4+top*3/4);
       for i in [['快递',1],['仓库',3],['备注',5],['标签',7],['商品',9]]:zb['批量修改'+i[0]]=(50*i[1],bottom*3/4+top/4)
       zb['批量修改快递']=(60*1,bottom*3/4+top/4);zb['批量修改仓库']=(60*3,bottom*3/4+top/4);
       left,top,right,bottom=win32gui.GetWindowRect(self.zjb(handle)[0]);zb['订单列表']=(left,top);zb['订单全选']=(left+50,top+10)###########
       handle=self.zjb(handle)[3]#订单详情，主要用到下面的订单信息
       left,top,right,bottom=win32gui.GetWindowRect(handle);pyautogui.click(left,top,1);pyautogui.press('right',2,1)
       handle=self.zsjb(self.zsjb(handle,'订单信息')[0])#订单信息
       left,top,right,bottom=win32gui.GetWindowRect(handle[96]);zb['商品信息']=(left,top-20);zb['商品全选']=(right-10,bottom-15)
       left,top,right,bottom=win32gui.GetWindowRect(handle[62]);zb['订单信息']=(right,top-20)
       e,r=zb['订单全选'];     zb['第一条订单']=(e+100,r+20)
       e,r=zb['关闭销售订单明细'];     zb['zb']=(e+370,r)
       return jb,zb

if __name__ == '__main__':
    pass
    z=GetHandle()
    jb,zb=z.jbzb_run()


