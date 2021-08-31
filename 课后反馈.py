import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main

def today_feedback(cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210324'):
    p=main.FeedBackAfterClass()
    p.draw(cus=cus,ins=ins,date_input=date_input)


if __name__=='__main__':
    # 当天课后生成
    today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')