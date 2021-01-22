import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),'modules'))
import days_cal
import json
import pandas as pd
import numpy as np
from datetime import datetime
import re
from PIL import Image

class MingHu:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        with open (os.path.join(self.dir,'config.dazhi'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line) 

        self.cus_file_dir=config['会员档案文件夹']

    def read_cus(self,cus='MH001韦美霜',start_time='20150101',end_time=''):
        #------------基本情况--------
        out=Vividict()       

        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        xls_name=os.path.join(self.cus_file_dir,cus+'.xlsx')
        
        df_basic=pd.read_excel(xls_name,sheet_name='基本情况')        
        out['nickname']=df_basic['昵称'].tolist()[0] #昵称
        out['sex']=df_basic['性别'].tolist()[0] #性别

        #年龄
        if not df_basic['出生年月'].empty:
            birth=df_basic['出生年月'].tolist()[0]
            if len(str(birth))==4:
                birth=str(birth)+'0101'
                age=days_cal.calculate_age(birth)
                out['age']=age
            elif len(str(birth))==6:
                birth=str(birth)+'01'
                age=days_cal.calculate_age(birth)
                out['age']=age
            elif len(str(birth))==8:
                age=days_cal.calculate_age(str(birth))
                out['age']=age
            else:
                out['age']=''            
        else:
            out['age']='' 
        

        #------------身体数据--------
        df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        df_body=df_body[(df_body['时间']>=start_time) & (df_body['时间']<=end_time)] #根据时间段筛选记录
        df_body=df_body.fillna(0)
        if df_body.empty:
            out['body']=''
        else:
            body_recent=df_body[df_body['时间']==df_body['时间'].max()].values.tolist()[0]       
            out['body']['time']=body_recent[0]
            out['body']['ht']=body_recent[1]
            out['body']['wt']=body_recent[2]
            out['body']['bfr']=body_recent[3]
            out['body']['chest']=body_recent[4]
            out['body']['l_arm']=body_recent[5]
            out['body']['r_arm']=body_recent[6]
            out['body']['waist']=body_recent[7]
            out['body']['hip']=body_recent[8]
            out['body']['l_leg']=body_recent[9]
            out['body']['r_leg']=body_recent[10]
            out['body']['l_calf']=body_recent[11]
            out['body']['r_calf']=body_recent[12]
            

        #------------训练数据--------
        infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos=infos.iloc[:,0:10] #取前10列
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','次数','教练姓名','教练评语']

        infos=infos[(infos['时间']>=start_time) & (infos['时间']<=end_time)] #根据时间段筛选记录

        #起止日期
        out['interval']=[infos['时间'].min(),infos['时间'].max()]  

        #抗阻训练
        train_dates=infos.groupby(['时间','目标肌群'])
        train_big_type=[]
        for dt,itm in train_dates:
            train_big_type.append(list(dt))
        df_train_big_type=pd.DataFrame(train_big_type)    
        if not df_train_big_type.empty:   
            df_train_big_type.columns=['时间','目标肌群']
            _sum_train_items=df_train_big_type.groupby(['目标肌群'],as_index=False)
            sum_train_items=pd.DataFrame(_sum_train_items.count())  
            sum_train_items.dropna(axis=0, how='any', inplace=True)
            sum_train_items=sum_train_items[sum_train_items['目标肌群']!=' '].values
            if len(sum_train_items)>0:
                for itm in sum_train_items:
                    out['train']['muscle'][itm[0]]=itm[1]
            else:
                out['train']['muscle']=''
        else:
            out['train']['muscle']=''

        #有氧训练总时长
        out['train']['oxy_time']=infos['有氧时长'].sum()

        # print(out)
        return out

    def draw(self,cus='MH001韦美霜',start_time='20150101',end_time=''):
        infos=self.read_cus(cus=cus,start_time=start_time,end_time=end_time)
        print(infos)

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    p=MingHu()
    p.draw(cus='MH001韦美霜',start_time='20200901',end_time='20200905')

