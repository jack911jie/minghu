import os
import sys
# sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),'modules'))
import main as minghu_main

def runMH(cus='MH000唐青剑',ins='MHINS002韦越棋',start_time='20200101',end_time=''):
    p=minghu_main.MingHu()
    p.draw(cus=cus,ins=ins,start_time=start_time,end_time=end_time)

if __name__=='__main__':
    runMH(cus='MH000唐青剑',ins='MHINS001陆伟杰',start_time='20200101',end_time='')