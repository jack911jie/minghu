import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main


class FeedBack:

    def today_feedback(self,cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210324'):
        p=main.FeedBackAfterClass()
        p.draw(cus=cus,ins=ins,date_input=date_input)

    def today_feedback_group(self,ins='MHINS002韦越棋',date_input='20210727',open_dir='no'):
        p=main.FeedBackAfterClass()
        p.group_afterclass(ins=ins,date_input=date_input,open_dir=open_dir)

if __name__=='__main__':
################################################################################################
    p=FeedBack()
    p.today_feedback_group(ins='MHINS002韦越棋',date_input='20210803',open_dir='no')