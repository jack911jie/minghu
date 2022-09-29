import os
import sys
import pandas as pd
import re


class Data:
    def __init__(self):
        pass

    def contain_space(self,lst):
        sp=list(filter(lambda x: re.match(r'Unnamed:.*',x),lst))
        if sp:
            return True
        else:
            return False


    def coach_data(self,work_dir='D:\\工作目录\\铭湖健身\\数据',fn='教练工作日志.xlsx'):
        xlsx=os.path.join(work_dir,fn)

        df_coach_list=list(pd.read_excel(xlsx,sheet_name=None))

        coach=[]

        _n=0
        for li in df_coach_list:
            if re.match(r'20\d\d-\d\d',li):
                sht=pd.read_excel(xlsx,sheet_name=li)
                if  self.contain_space(list(sht.columns)):
                    print('在 ',li,' 标题行中有空数据。')
                    exit(0)
                else:
                    coach.append(sht)
        
        _df_coach=pd.concat(coach)
        df_coach=_df_coach[~pd.isnull(_df_coach['会员姓名'])]
        df_coach=df_coach.iloc[:,1:]

        save_dir=os.path.join(work_dir,'教练工作日志合并.xlsx')
        df_coach.to_excel(save_dir,index=False)
        os.startfile(work_dir)
        print('完成')



if __name__=='__main__':
    p=Data()
    p.coach_data(work_dir='e:\\工作目录\\铭湖健身\\数据',fn='教练工作日志.xlsx')