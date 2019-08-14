import time
import win32gui,win32ui,win32con, win32api
def window_capture(filename):#截屏，传入保存文件名，截屏并保存
      hwnd = 0 # 窗口的编号，0号表示当前活跃窗口
      hwndDC = win32gui.GetWindowDC(hwnd)# 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
      mfcDC = win32ui.CreateDCFromHandle(hwndDC)# 根据窗口的DC获取mfcDC
      saveDC = mfcDC.CreateCompatibleDC()# mfcDC创建可兼容的DC
      saveBitMap = win32ui.CreateBitmap()# 创建bigmap准备保存图片
      MoniterDev = win32api.EnumDisplayMonitors(None, None)# 获取监控器信息
      w = MoniterDev[0][2][2]
      h = MoniterDev[0][2][3]
      # print w,h　　　#图片大小
      saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)# 为bitmap开辟空间
      saveDC.SelectObject(saveBitMap)# 高度saveDC，将截图保存到saveBitmap中
      saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)# 截取从左上角（0，0）长宽为（w，h）的图片
      saveBitMap.SaveBitmapFile(saveDC, filename)

import cv2
import numpy
import aircv
def matchImg(imgobj,confidence=0.5,quzhi='result'):#传入小图、相似度、取值，返回小图所在屏幕的坐标
      '''取值分别为result【中心点坐标】，rectangle【四角坐标】，confidence【相似度】'''
      imobj = cv2.imdecode(numpy.fromfile(imgobj,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
      imgsrc='pm001.jpg'
      window_capture(imgsrc)#截屏
      imsrc = cv2.imdecode(numpy.fromfile(imgsrc,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
      match_result = aircv.find_all_template(imsrc,imobj,confidence)
      if match_result:return [i[quzhi] for i in match_result]
      else:return False# ;print('%s查找失败'% imgobj[:-4])

def matchImg1(imgobj,imgsrc,confidence=0.5,quzhi='result'):#传入小图、大图、相似度、取值，返回小图所在大图的坐标
      '''取值分别为result【中心点坐标】，rectangle【四角坐标】，confidence【相似度】'''
      ## imdecode读取的是rgb，如果后续需要opencv处理的话，需要转换成bgr，转换后图片颜色会变化
      ##cv_img=cv2.cvtColor(cv_img,cv2.COLOR_RGB2BGR)
      imobj = cv2.imdecode(numpy.fromfile(imgobj,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
      imsrc = cv2.imdecode(numpy.fromfile(imgsrc,dtype=numpy.uint8),-1)#解决imread不能读取中文路径的问题
      #imsrc = aircv.imread(imgsrc)#读取图像
      #imobj = aircv.imread(imgobj)#读取图像
      #match_result = aircv.find_template(imsrc,imobj,confidence)
      match_result = aircv.find_all_template(imsrc,imobj,confidence)
      #match_result[0]={'result': (1481.0, 856.0), 'rectangle': ((1459, 845), (1459, 867), (1503, 845), (1503, 867)), 'confidence': 0.6873950958251953}
      if match_result:return [i[quzhi] for i in match_result]
      else:return False#;print('%s查找失败'% imgobj[:-4])

if __name__ == '__main__':
    import pyautogui#模拟鼠标键盘操作
    a=window_capture("订单处理中心.jpg")
    print(a)
    b=matchImg('./图片/全选.jpg','./图片/订单处理中心.jpg')
    #pyautogui.click(*b)
