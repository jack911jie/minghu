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
                    
                    print('在 ',li,' 标题行中有空数据。\n',list(sht.columns))
                    exit(0)
                else:
                    coach.append(sht)
        
        _df_coach=pd.concat(coach)
        df_coach=_df_coach[~pd.isnull(_df_coach['会员姓名'])]
        df_coach=df_coach.iloc[:,1:]

        return df_coach

    def batch_coach_data(self,work_dir='D:\\工作目录\\铭湖健身\\数据'):
        df_list=[]
        for fn in os.listdir(work_dir):
            if re.match(r'教练工作日志-\d{4}.xlsx',fn):
                df=self.coach_data(work_dir=work_dir,fn=fn)
                df_list.append(df)
        
        df_all=pd.concat(df_list)

        return df_all

    def exp_coach_data(self,save_dir_input='D:\\工作目录\\铭湖健身\\数据'):
        df_all=self.batch_coach_data(work_dir=save_dir_input)
        save_name=os.path.join(save_dir_input,'教练工作日志合并.xlsx')
        df_all.to_excel(save_name,index=False)
        os.startfile(save_dir_input)
        print('完成')



if __name__=='__main__':
    p=Data()
    # p.coach_data(work_dir='e:\\工作目录\\铭湖健身\\数据',fn='教练工作日志.xlsx')
    p.exp_coach_data(save_dir_input='E:\\temp\\minghu')
    # k=p.batch_coach_data(work_dir='E:\\temp\\minghu')
    # print(k)