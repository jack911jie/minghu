import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main

# print(os.path.join(os.path.dirname(__file__),'modules'))

def cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701'):
    p=main.MingHu()
    p.draw(cus=cus,ins=ins,start_time=start_time,end_time=end_time)

def group_input():
    p=main.GroupDataInput()
    p.data_input()


if __name__=='__main__':
    #反馈
    # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')

    #批量录入
    group_input()





