import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main

# print(os.path.join(os.path.dirname(__file__),'modules'))

def cus_feedback(place='minghu',cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701',adj_bfr='yes',adj_src='prg',gui=''):
    p=main.MingHu(place=place,adj_bfr=adj_bfr,adj_src=adj_src,gui=gui)
    p.draw(cus=cus,ins=ins,start_time=start_time,end_time=end_time)

def group_input():
    p=main.GroupDataInput()
    p.data_input()

def today_feedback(place='minghu',cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210324'):
    p=main.FeedBackAfterClass(place=place)
    p.draw(cus=cus,ins=ins,date_input=date_input)

def auto_xls(cus_name_input='',mode='prgrm',gui=''):
    p=main.MingHu()
    new_cus_fn=p.auto_cus_xls(cus_name_input=cus_name_input,mode=mode,gui=gui)
    return new_cus_fn


if __name__=='__main__':
    #反馈
    # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')

    #批量录入
    group_input()

    #当天课后生成
    # today_feedback(cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210619')

    #新增会员
    # auto_xls(cus_name_input='测试',mode='prgrm',gui='')



