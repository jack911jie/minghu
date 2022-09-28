import os
import sys
import pandas as pd
import re


class Data:
    def __init__(self):
        pass

    def coach_data(self,work_dir='D:\\工作目录\\铭湖健身\\数据',fn='教练工作日志.xlsx'):
        xlsx=os.path.join(work_dir,fn)

        df_coach_list=list(pd.read_excel(xlsx,sheet_name=None))

        coach=[]

        _n=0
        for li in df_coach_list:
            if re.match(r'20\d\d-\d\d',li):
                coach.append(pd.read_excel(xlsx,sheet_name=li))
        
        _df_coach=pd.concat(coach)
        df_coach=_df_coach[~pd.isnull(_df_coach['会员姓名'])]

        print(df_coach)
        df_coach.to_clipboard()





if __name__=='__main__':
    p=Data()
    p.coach_data()