import os
import json
import pandas as pd
import numpy as np
from datetime import datetime
import re

class MingHu:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        with open (os.path.join(self.dir,'config.linux'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line) 

        self.cus_file_dir=config['会员档案文件夹']

    def read_cus(self,cus='MH001韦美霜',start_time='20160101',end_time=''):
        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        interval=end_time-start_time
        interval_days=interval.days        
        nat = np.datetime64('NaT')
        xls_name=os.path.join(self.cus_file_dir,cus+'.xlsx')
        infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','次数','教练姓名','教练评语']
        # infos.rename(columns={'0':'时间','Unnamed: 1':'形式','Unnamed: 2':'目标肌群', \
        #                       'Unnamed: 3':'有氧项目','Unnamed: 4':'有氧时长','Unnamed: 5':'力量内容', \
        #                           'Unnamed: 6':'重量','Unnamed: 7':'次数','Unnamed: 8':'教练姓名','Unnamed: 9':'教练评语',},inplace=True)
        # print(infos['教练评语'])
        traing_times=infos['时间'].nunique()
        
        # print('会员训练次数：',traing_times)

        train_dates=infos.groupby(['时间','目标肌群'])
        # print(list(train_dates))
        train_big_type=[]
        for dt,itm in train_dates:
            train_big_type.append(list(dt))
        # print(train_big_type)
        df_train_big_type=pd.DataFrame(train_big_type)
        df_train_big_type.columns=['时间','目标肌群']
        # print(df_train_big_type)
        _sum_train_items=df_train_big_type.groupby(['目标肌群'],as_index=False)
        print(_sum_train_items.count())
        sum_train_items=[]
        for dst_muscle,itm in _sum_train_items:
            print(dst_muscle)
        # print(sum_train_items)
            
        
        # print(list(train_dates))
        # k=datetime.strptime(str(train_dates[1]).split('T')[0],'%Y-%m-%d')  
        # now=datetime.now()
        # interval=now-k

        # print(interval.days)
        # print(t[0])

        # fenzhu_times=infos.groupby(['时间','目标肌群'])
        # print(fenzhu_times)


if __name__=='__main__':
    p=MingHu()
    p.read_cus()

