import re,time,xlrd


def diZhiChuLi(beizhu):
    '''地址处理:传入存在改地址的备注，返回address'''
    beizhu=beizhu.replace('\n','')#去除备注中的换行
    dz01=re.findall('改地址[:：](.*?)┋',beizhu)
    d=''
    if dz01:
        d=''.join(e for e in dz01[-1] if e.isalnum() or e in ':：;；._-')#去除备注里的符号
        c='(?:联系方式|电话|手机).{,5}?(?=1\d{10})|(?:所在地区|详细地址|收货人|姓名|地址)[:：]'
        d=re.sub(c,'',d)
    if d:d=re.split('[;；]',d)
    if len(d)>3:return False
    dh=[]
    xm=[]
    dz=[]
    for i in d:
        cd=len(i)
        sz=len(re.findall('\d',i))
        sj=re.findall('省|市|区|县|乡|镇|村|路|号|街|道|',i)
        try:
            while 1:sj.remove('')
        except:pass
        sj=len(sj)
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
        t1= xlrd.open_workbook(r'地区.xlsx').sheets()[2]
        li=[] ;dz1=''.join(re.findall('.*(?<=[区县市旗岛域辖镇乡台阁仔])',dz))
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
        if not zhi1:print('改地址备注权重不足以判断地区：备注内容非地址',dz);return False
        elif len(zhi1)==1 or zhi1[0]!=zhi1[1]:
            if zhi1[-1]>10:sqx=t1.row_values(li[zhi.index(zhi1[-1])][0])[:3]
            else:print('改地址备注权重不足以判断地区：备注内容不全',dz);return False
        else:
            if zhi1[0]>10:
                g5=t1.row_values(li[zhi.index(zhi1[-1])][0])[4]
                #sqx=[t1.row_values(i)[:3] for i in range(t1.nrows) if (t1.row_values(i)[4]==g5 and '其它区' in t1.row_values(i)[5])][0]
                try:sqx=[t1.row_values(i)[:3] for i in range(t1.nrows) if (t1.row_values(i)[4]==g5 and '其它区' in t1.row_values(i)[5])][0]
                except:print('其他区不存在1111111111111111',dz);return False
            else:print('修改地址备注异常222222222222',dz);return False
    if sqx:sqx = [int(x) for x in sqx]
    address=(xm,dh,dz,sqx)#姓名、电话、地址、地区编码的列表
    for i in address:
        if i:return address
    #print(address);print('电话',dh);print('姓名',xm);print('地址',dz);print('编码',sqx)
    return False


t=xlrd.open_workbook('改地址.xls')
s=t.sheet_by_index(0)
h=s.nrows
for i in range(h):
    hv=s.row_values(i)
    bz=hv[0].replace('\n','')
    b=diZhiChuLi(bz)
    if not b:
        print(bz)
'''
    dz01=re.findall('改地址[:：](.*?)┋',bz)
    if dz01:
        d=''.join(e for e in dz01[-1] if e.isalnum() or e in ':：;；._-')#去除备注里的符号
        c='(?:联系方式|电话|手机).{,5}?(?=1\d{10})|(?:所在地区|详细地址|收货人|姓名|地址)[:：]'
        d=re.sub(c,'',d)
        #if len(d)<10:print(d)
        if d:
            d=re.split('[;；]',d)
            dh=xm=dz=''
            for i in d:
                if re.findall(r"1[3456789]\d{9}$",i):dh=i
                elif len(i)<5:xm=i
                else:dz=i
            #print('%s	%s	%s'%(dh,xm,dz))

'''
