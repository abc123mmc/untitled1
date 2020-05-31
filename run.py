# -*- coding: utf-8 -*-
from get_data import *
#pyautogui.FAILSAFE = False



def get_danhao():
    """选择一小时前的订单"""
    t=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()-3600))
    Fun.cxdd(zb,i=['',t])
    Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
    ddclzx_dict = get_config.GetCofig.get_layout(setting.ddclzx)  # 返回所有订单字及段位置的字典
    c=GetData.ddsj(zb,ddclzx_dict,ptdh=True)
    b=[]
    for i in range(0,len(c),60):
        b.append(c[i:i+60])
    return b

def go_go():
    e=get_danhao()
    for i in e:
        ptdh01=','.join(i)
        Fun.cxdd(zb,ptdh01)
        Fun.cxpanduan(setting.jia_zz,setting.quan_pt)
        b,c,d = GetData().panduan()
        time.sleep(1)
        for j in b['改快递']:
            ptdh01 = ','.join(b['改快递'][j])
            if ptdh01:
                Fun.cxdd(zb, ptdh01)
                Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
                Fun.quan_xuan(zb,setting.jia_zz,setting.quan_pt)
                Fun.gai_kd(zb,j)
                Fun.cxdd(zb,'改快递%s已完成' % j)
                TxtWrite.write_in_txt(setting.gkd_f, f'{ptdh01}\n\n')
        for j in b['改地址']:
            Fun.cxdd(zb,j)
            Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
            Fun.gai_dz(*b['改地址'][j])
        if b['改地址']:
            padh01=','.join([i for i in b['改地址']])
            TxtWrite.write_in_txt(setting.gkd_f, f'{padh01}\n\n')
        for j in b['改赠品']:
            ptdh01=','.join(b['改赠品'])
            if ptdh01:
                time.sleep(2)
                Fun.cxdd(zb, ptdh01)
                Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
                Fun.quan_xuan(zb, setting.jia_zz, setting.quan_pt)
                Fun.gai_sp(zb)
                TxtWrite.write_in_txt(setting.gzp_f, f'{ptdh01}\n\n')

        ptdh01=','.join(d)
        print(f'人工审核:{ptdh01}')
        if ptdh01:#改标签,人工审核
            time.sleep(2)
            Fun.cxdd(zb, ptdh01)
            Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
            Fun.quan_xuan(zb, setting.jia_zz, setting.quan_pt)
            Fun.gai_bq(zb,bqs01=1)
            TxtWrite.write_in_txt(setting.gbq_f, f'{ptdh01}\n\n')

        ptdh01=','.join(c)
        print(f'可以审核:{ptdh01}')
        if ptdh01:
            time.sleep(2)
            Fun.cxdd(zb, ptdh01)
            Fun.cxpanduan(setting.jia_zz, setting.quan_pt)
            Fun.quan_xuan(zb, setting.jia_zz, setting.quan_pt)
            Fun.shenghe(zb)
            t=time.localtime()
            TxtWrite.write_in_txt(setting.zdsh_f,
                                  f'{t[0]}-{t[1]}-{t[2]} {t[3]}:{t[4]}:{5} >>> {ptdh01}\n\n')
    Fun.cxdd(zb,'所有获取到的订单审核完成*******')

if __name__ == '__main__':
    try:go_go()
    except:TxtWrite.error(setting.error_info)
    pass


