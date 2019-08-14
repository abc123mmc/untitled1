import os
import re
import string
#此函数的作用为递归查找文件夹下所有的文件
def dirlist(mainpath, allfilelist):
   filelist = os.listdir(mainpath)
   for filename in filelist:
      filepath = os.path.join(mainpath, filename)
      if os.path.isdir(filepath):
         dirlist(filepath, allfilelist)
      else:
         allfilelist.append(filepath)
   return allfilelist

#此函数的作用为在某个文件中匹配特定字符串
def findstr(filename,keyword):
   global everyline
   right=[]
   fp=open( filename,'r',encoding='UTF-8')
   #fp=open( filename, 'rb')
   for everyline in fp:
       print
      if re.search(keyword,everyline,re.I):
         right.append(filename)
         break
   return right

if __name__ == "__main__":
    keyword='44556'
    allfile=dirlist(r"C:\Users\Administrator\PycharmProjects\liaotian01",[])
    for i in range(len(allfile)):
        if findstr(allfile[i],keyword):
            print(allfile[i])
