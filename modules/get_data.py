import os
import sys
import pandas as pd
import days_cal
from datetime import datetime
import random
from tkinter import simpledialog


class ReadAndExportData:
    def __init__(self,adj_bfr='yes',adj_src='prg',gui=''):
        self.adj_bfr=adj_bfr
        self.adj_src=adj_src
        self.gui=gui


    def read_excel(self,cus_file_dir,cus='MH001韦美霜'):
        xls_name=os.path.join(cus_file_dir,cus+'.xlsx')
        df_basic=pd.read_excel(xls_name,sheet_name='基本情况')    
        df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        df_infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)

        return [df_basic,df_body,df_infos]

    def exp_cus_prd(self,cus_file_dir,cus='MH001韦美霜',start_time='20150101',end_time=''):
        df=self.read_excel(cus_file_dir=cus_file_dir,cus=cus)
        df_basic=df[0] #基本情况
        df_body=df[1] #身体围度
        infos=df[2] #训练情况

        #------------基本情况--------
        out=Vividict()       

        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        
        
        # df_basic=pd.read_excel(xls_name,sheet_name='基本情况')        
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
        # df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        # df_body=df_body[(df_body['时间']>=start_time) & (df_body['时间']<=end_time)] #根据时间段筛选记录
        df_body=df_body[df_body['时间']<=end_time] #根据时间段筛选记录
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
            out['body']['ht_lung']=body_recent[13]
            out['body']['balance']=body_recent[14]
            out['body']['power']=body_recent[15]
            out['body']['flexibility']=body_recent[16]
            out['body']['core']=body_recent[17]

            bfr_data=cals()
            bfr=bfr_data.bfr(age=age,sex=out['sex'],ht=out['body']['ht'],wt=out['body']['wt'],waist=out['body']['waist'],
                adj_bfr=self.adj_bfr,adj_src=self.adj_src,gui=self.gui,formula=1)
            
            out['body']['bfr']=bfr

        #------------训练数据--------
        # infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos=infos.iloc[:,0:12] #取前11列
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','距离','次数','消耗热量','教练姓名','教练评语']
        # print(infos.dropna(how='all'))
        if infos.dropna(how='all').shape[0]!=0:
            infos=infos[(infos['时间']>=start_time) & (infos['时间']<=end_time)] #根据时间段筛选记录      

            # print('168 line:',infos)

            #起止日期
            out['interval']=[infos['时间'].min(),infos['时间'].max()]  
            out['interval_input']=[start_time,end_time]

            #次数
            train_times=infos.groupby(['时间'],as_index=False).nunique()['时间'].nunique()
            out['train_times']=train_times

            #抗阻训练
            train_dates=infos.groupby(['时间','目标肌群'])
            # print(train_dates)
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
            out['train']['oxy_time']=infos['有氧时长'].apply(lambda x:int(x) if isinstance(x,str) else x).sum()
        else:
            out['train']=''

        # print('201 line:',out)
        return out

class ReadAndExportDataNew:
    def __init__(self,adj_bfr='yes',adj_src='prg',gui=''):
        self.adj_bfr=adj_bfr
        self.adj_src=adj_src
        self.gui=gui

    def read_excel(self,cus_file_dir,cus='MH001韦美霜'):
        xls_name=os.path.join(cus_file_dir,cus+'.xlsx')
        df_basic=pd.read_excel(xls_name,sheet_name='基本情况')    
        df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        df_infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)

        return [df_basic,df_body,df_infos]

    def exp_cus_prd(self,cus_file_dir,cus='MH001韦美霜',start_time='20150101',end_time=''):
        df=self.read_excel(cus_file_dir=cus_file_dir,cus=cus)
        df_basic=df[0] #基本情况
        df_body=df[1] #身体围度
        infos=df[2] #训练情况

        #------------基本情况--------
        out=Vividict()       

        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        
        
        # df_basic=pd.read_excel(xls_name,sheet_name='基本情况')        
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
        # df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        # df_body=df_body[(df_body['时间']>=start_time) & (df_body['时间']<=end_time)] #根据时间段筛选记录
        df_body=df_body[df_body['时间']<=end_time] #根据时间段筛选记录
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
            out['body']['ht_lung']=body_recent[13]
            out['body']['balance']=body_recent[14]
            out['body']['power']=body_recent[15]
            out['body']['flexibility']=body_recent[16]
            out['body']['core']=body_recent[17]

            bfr_data=cals()
            # age,sex,ht,wt,waist,adj_que='yes',adj_src='prg',gui='',formula=1
            bfr=bfr_data.bfr(age=age,sex=out['sex'],ht=out['body']['ht'],wt=out['body']['wt'],waist=out['body']['waist'],
                            adj_bfr=self.adj_bfr,adj_src=self.adj_src,gui=self.gui,formula=1)
            out['body']['bfr']=bfr

        #------------训练数据--------
        # infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos=infos.iloc[:,0:12] #取前10列
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','距离','次数','消耗热量','教练姓名','教练评语']
        # print(infos.dropna(how='all'))
        if infos.dropna(how='all').shape[0]!=0:
            infos=infos[(infos['时间']>=start_time) & (infos['时间']<=end_time)] #根据时间段筛选记录      

            # print('168 line:',infos)

            #起止日期
            out['interval']=[infos['时间'].min(),infos['时间'].max()]  
            out['interval_input']=[start_time,end_time]

            #次数
            train_times=infos.groupby(['时间'],as_index=False).nunique()['时间'].nunique()
            out['train_times']=train_times

            #抗阻训练
            #细项
            train_muscle_info=[]
            train_muscle_data=infos.groupby(['力量内容'])
            for mscl_item,mscl_count in train_muscle_data:
                train_muscle_info.append([mscl_item,mscl_count['重量'].sum(),mscl_count['次数'].sum()])
            out['train']['muscle_item']=train_muscle_info
            # print(out['train']['muscle_item'])   

            #大项
            train_dates=infos.groupby(['时间','目标肌群'])
            # print(train_dates)
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
        
            #有氧训练
            # 总时长
            out['train']['oxy_time']=infos['有氧时长'].apply(lambda x:int(x) if isinstance(x,str) else x).sum()
            # 有氧细项
            _df_oxy_data=infos[['有氧项目','有氧时长']]
            df_oxy_data=_df_oxy_data.dropna(axis=0,how='all')
            oxy_time_group_sum=df_oxy_data.groupby(['有氧项目'])
            oxy_time_group=pd.DataFrame(oxy_time_group_sum.sum())
            oxy_data=[]
            for num_oxy,oxy_time in enumerate(oxy_time_group['有氧时长'].apply(lambda x:int(x)).values.tolist()):
                oxy_data.append([oxy_time_group.index.tolist()[num_oxy],oxy_time])
            
            out['train']['oxy_infos']=oxy_data

            #消耗热量
            _calories=infos['消耗热量']
            calories=_calories.dropna(axis=0,how='all')
            burn_cal=calories.sum()
            out['train']['calories']=burn_cal
            


        else:
            out['train']=''

        # print('201 line:',out)
        return out


class cals:
    def bfr(self,age,sex,ht,wt,waist,adj_bfr='yes',adj_src='prg',gui='',formula=1):
            # 女：
            # 参数a=腰围（cm）×0.74
            # 参数b=体重（kg）×0.082+34.89
            # 体脂肪重量（kg）=a－b
            # 体脂率=（身体脂肪总重量÷体重）×100%
            # 男：
            # 参数a=腰围（cm）×0.74
            # 参数b=体重（kg）×0.082+44.74
            # 体脂肪重量（kg）=a－b
            # 体脂率=（身体脂肪总重量÷体重）×100%
        if formula==1:
            if sex=='女':
                k=34.89
            if sex=='男':
                k=44.74
            a=waist*0.74
            b=wt*0.082+k
            fat=a-b

            bfr=fat/wt

        elif formula==2:
            # 1.2×BMI+0.23×年龄-5.4-10.8×性别（男为1，女为0）
            if sex=='女':
                k=0
            if sex=='男':
                k=1

            bmi=wt/((ht/100)*(ht/100))
            bfr=1.2*bmi+0.23*age-5.4-10.8*k

        if adj_bfr=='yes':
            if adj_src=='prg':
                adj_bfr_value=input('\n计算出的体脂率为 {}，如需修改请直接输入体脂率（如：12.46%），不需要修改请直接按回车——\n\n'.format(str('{:.2%}'.format(bfr))))
            elif adj_src=='gui':
                # gui.withdraw()
                # the input dialog
                print('\n计算出的体脂率为 {}，如需修改请直接输入体脂率（如：12.46%），不需要修改请直接按回车——\n\n'.format(str('{:.2%}'.format(bfr))))
                adj_bfr_value = simpledialog.askstring(title="输入体脂率",
                                                prompt="请输入体脂率(仅输入数字):")
                if adj_bfr_value:
                    print('修正体脂率为 {}%'.format(adj_bfr_value))
                
            if adj_bfr_value:
                if '%' in adj_bfr_value:
                    bfr=float(adj_bfr_value[:-1])/100
                else:
                    bfr=float(adj_bfr_value)/100
        

        return bfr

class ReadDiet:
    def __init__(self,fn_diet='D:\\Documents\\WXWork\\1688851376227744\\WeDrive\\铭湖健身工作室\\05-专业资料\\减脂饮食建议表.xlsx'):
        self.fn_diet=fn_diet
    
    def exp_diet_suggests(self):
        df=pd.read_excel(self.fn_diet,sheet_name='饮食建议')
        df.dropna(how='any',inplace=True)
        diet_suggests=df['饮食建议'].apply(lambda x:str(x).strip()).values.tolist()
        # print(diet_suggests) 

        return diet_suggests




class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    p=ReadAndExportDataNew(adj_bfr='yes')
    res=p.exp_cus_prd(cus_file_dir="D:\\Documents\\WXWork\\1688851376227744\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料",cus='MH003吕雅颖',start_time='20210727',end_time='20210727')
    print(res)
    # p=ReadDiet()
    # su=p.exp_diet_suggests()
    # print(random.choice(su))