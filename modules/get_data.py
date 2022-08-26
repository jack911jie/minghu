import os
import sys
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
import days_cal
import numpy as np
from datetime import datetime
from datetime import timedelta
import random
from tkinter import simpledialog
import re
import copy


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
        infos=infos.iloc[:,0:14] #取前13列
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
            infos_muscle=pd.DataFrame(infos,columns=['时间','力量内容','重量','次数'])
            infos_muscle['次数'].fillna(1,inplace=True)
            infos_muscle.dropna(subset=['重量'],inplace=True)
            # print(infos_muscle)
            infos_muscle['合计重量']=infos_muscle['重量']*infos_muscle['次数']
            out['train']['muscle_total_wt']=infos_muscle['合计重量'].sum()
            train_muscle_data=infos.groupby(['力量内容'])
            for mscl_item,mscl_count in train_muscle_data:
                train_muscle_info.append([mscl_item,mscl_count['重量'].sum(),mscl_count['次数'].sum(),mscl_count['距离'].sum()])
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

            #教练评语
            _ins_cmts=infos['教练评语']
            _ins_cmts.dropna(axis=0,how='any',inplace=True)
            ins_cmts=list(_ins_cmts)
            out['train']['ins_cmts']=ins_cmts

            # print(_ins_cmts)
            
            # #训练次数
            # train_types=['常规私教','团课']
            # all_train_amount=0
            # for train_type in train_types:
            #     df_train_amount=infos[infos['课程']==train_type][['时间','节次','课程']]
            #     df_train_amount=df_train_amount.drop_duplicates(['时间','节次']).reset_index(drop=True)
            #     total_train_amount=df_train_amount['时间'].count()
            #     all_train_amount+=total_train_amount
            #     df_train_amount['年']=pd.to_datetime(df_train_amount['时间']).dt.strftime('%Y')
            #     df_train_amount['月']=pd.to_datetime(df_train_amount['时间']).dt.strftime('%m')
            #     train_amt_month=df_train_amount.groupby(['年','月'])['时间'].count().reset_index()
            #     # train_amount_gp=df_train_amount.groupby(['月'])['时间'].count().reset_index()
            #     train_amt_month.columns=['年','月','次数']
            # # print(train_amount_gp)
            #     out['train_stat']['total_train_amt'][train_type]=total_train_amount
            #     out['train_stat']['train_amt_month'][train_type]=train_amt_month

            # # print(df_train_amount['时间'].count())
            # # print(total_train_amount)
            # out['train_stat']['all_train_amount']=all_train_amount
            

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
            if sex=='女' or sex=='f':
                k=34.89
            if sex=='男' or sex=='m':
                k=44.74
            a=waist*0.74
            b=wt*0.082+k
            fat=a-b

            bfr=fat/wt

        elif formula==2:
            # 1.2×BMI+0.23×年龄-5.4-10.8×性别（男为1，女为0）
            if sex=='女' or sex=='f':
                k=0
            if sex=='男' or sex=='m':
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


    def bmr(self,sex='f',ht=161,wt=61,age=35):
        if ht=='' or wt=='' or age=='' or np.isnan(ht) or np.isnan(wt):
            print('身体数据或年龄未填写，请核实。')
            exit(0)
        if sex=='f' or sex=='女':
            bmr=665.1+9.6*wt+1.8*ht-4.7*age
        else:
            bmr=66.5+13.8*wt+5*ht-6.8*age
        return bmr

class ReadCourses:
    def __init__(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室',dl_taken_fn='D:\\铭湖健身工作目录\\教练工作日志\\教练工作日志.xlsx'):
        self.work_dir=work_dir
        self.base_fn=os.path.join(work_dir,'01-会员管理','工作文档','20220531私教会员剩余课程节数.xlsx')
        self.gp_base_fn=os.path.join(work_dir,'01-会员管理','工作文档','20220430限时课程会员剩余课程节数.xlsx')
        self.taken_fn=dl_taken_fn
        self.df_base=pd.read_excel(self.base_fn)
        self.gp_base=pd.read_excel(self.gp_base_fn)
        self.df_base['备注'].fillna('无',inplace=True)

    def read_excel_taken(self,start_time='20220501',nowtime='20220604'):
        fn_shtnames=pd.ExcelFile(self.taken_fn).sheet_names
        df_takens=[]
        for shtname in fn_shtnames:
            if  re.match(r'\d{4}-\d{2}',shtname):
                # print(shtname)
                df_taken=pd.read_excel(self.taken_fn,sheet_name=shtname)
                df_takens.append(df_taken)
        _df_all_taken=pd.concat(df_takens)
        df_all_taken=copy.deepcopy(_df_all_taken)
        # print('llllllllll:',nowtime)
        if nowtime=='':
            df_all_taken=df_all_taken[df_all_taken['日期']>=datetime.strptime(start_time, '%Y%m%d')]
        else:
            df_all_taken=df_all_taken[(df_all_taken['日期']>=datetime.strptime(start_time, '%Y%m%d')) & (df_all_taken['日期']<=datetime.strptime(nowtime, '%Y%m%d'))]
        # print(df_all_taken)
        df_all_taken.columns=["序号","日期","时间","时长","课程类型","会员姓名","教练","是否完成","备注","体验课出单","出单日期"]
        df_all_taken['是否完成'].fillna('否',inplace=True)
        df_all_taken['备注'].fillna('无',inplace=True)
        df_all_taken['体验课出单'].fillna('不适用',inplace=True)
        df_all_taken['出单日期'].fillna('不适用',inplace=True)

        # print(df_all_taken)

        return  df_all_taken
    
    def cus_info(self,cus_name='MH016徐颖丽'):
        df_info=pd.read_excel(os.path.join(self.work_dir,'01-会员管理','会员资料',cus_name+'.xlsx'),sheet_name='基本情况')
        t_name=df_info['姓名'].tolist()[0]
        t_nickname=df_info['昵称'].tolist()[0]
        t_sex=df_info['性别'].tolist()[0]
        t_birth=df_info['出生年月'].tolist()[0]

        if t_nickname==t_name[1:]:
            callit=t_name
        else:
            callit=t_nickname
        
        if t_sex=='女':
            title='女士'
        else:
            title='先生'
        
        return {'name':t_name,'nickname':t_nickname,'sex':t_sex,'birthday':t_birth,'callit':callit,'title':title}


####################常规私教#####################

    def cus_taken(self,cus_name='MH016徐颖丽',crs_types=['常规私教课','初级团课'],start_time='20220501',nowtime='20220604'):
        all_cus_taken=self.read_excel_taken(start_time=start_time,nowtime=nowtime)
        _df_cus_takens=[]
        try:
            for crs_type in crs_types:
                df_cus_taken=all_cus_taken[(all_cus_taken['会员姓名']==cus_name) & (all_cus_taken['课程类型']==crs_type) & (all_cus_taken['是否完成']=='是')]
                _df_cus_takens.append(df_cus_taken)
            df_cus_takens=pd.concat(_df_cus_takens)
            # print(_df_cus_takens)
            gp_taken=df_cus_takens.groupby(['课程类型']).count().reset_index()
            gp_taken=gp_taken[['课程类型','会员姓名']]
            gp_taken.columns=['课程类型','上课次数']
            # gp_taken['上课次数']=gp_taken['上课次数'].apply(lambda x:x*-1)
        #如无上课的记录
        except Exception as e: 
            print(e)
            # gp_taken=''

        # print(gp_taken)
        return gp_taken
    
    def cus_buy(self,cus_name='MH016徐颖丽',crs_types=['常规私教课','团课'],data_fn='客户业务流水数据.xlsx',start_time='20220601'):
        df_gp_buy=pd.read_excel(os.path.join(self.work_dir,'02-运营规划','业务数据管理',data_fn),sheet_name='购课业务流水')
        # df_gp_buy['购课期数'].fillna(0,inplace=True)
        # df_gp_buy['购课节数'].fillna(0,inplace=True)
        df_gp_buy['购课时间'].fillna(0,inplace=True)
        df_gp_buy['备注'].fillna('无',inplace=True)

        df_gp_buy=df_gp_buy[df_gp_buy['购课时间']>=datetime.strptime(start_time,'%Y%m%d')]

        # print('df_gp_buy:',df_gp_buy)
        try:
            _df_cus_buy=df_gp_buy[df_gp_buy['购课编号'].str[:-8]==cus_name]

            _cus_buy=[]
            for crs_type in crs_types:
    
                if crs_type in ['常规私教课']:
                    pre_df_cus_buy_jieshu=_df_cus_buy[['购课编号','购课类型','购课节数','购课时间']]
                    pre_df_cus_buy_jieshu.columns=['购课编号','购课类型','购课数量','购课时间']
                    _cus_buy.append(pre_df_cus_buy_jieshu)
                elif crs_type in ['初级团课','限时私教','踏板团课']:
                    pre_df_cus_buy_qishu=_df_cus_buy[['购课编号','购课类型','购课期数','购课时间']]
                    pre_df_cus_buy_qishu.columns=['购课编号','购课类型','购课数量','购课时间']
                    _cus_buy.append(pre_df_cus_buy_qishu)
                else:
                    print('无效的课程类别')
                
            df_cus_buy=pd.concat(_cus_buy)
            df_cus_buy.dropna(how='any',inplace=True)

            df_cus_buy_cal=df_cus_buy.groupby(['购课类型']).sum().reset_index()
        #如无购课的记录
        except Exception as e:
            print(e)
            # df_cus_buy_cal=''

        # print('test buy cal',df_cus_buy_cal)
        return df_cus_buy_cal

    def ins_info(self,ins='MHINS001陆伟杰'):
        if re.match(r'^MHINS.*',ins):
            df_all_ins=pd.read_excel(os.path.join(self.work_dir,'03-教练管理','教练资料','教练信息.xlsx'))
            df_ins=df_all_ins[df_all_ins['员工编号']==ins[:8]]
        else:
            df_all_ins=pd.read_excel(os.path.join(self.work_dir,'03-教练管理','教练资料','教练信息.xlsx'))
            df_ins=df_all_ins[df_all_ins['姓名']==ins]
        return df_ins

    def cal_crs_remain(self,cus_name='MH016徐颖丽',crs_types=['常规私教课','团课']):
        # 客户的课程节数/期数的底
        _df_cus_takens=[]
        for crs_type in crs_types:
            df_cus_base=self.df_base[(self.df_base['客户名称']==cus_name) & (self.df_base['购买课程类型']==crs_type)]
            _df_cus_takens.append(df_cus_base)
        cus_base=pd.concat(_df_cus_takens)
        cus_base=cus_base[['客户名称','购买课程类型','剩余课时（节）']]
        cus_base.columns=['客户名称','课程类型','剩余课时（节）']
        cus_base.groupby(['课程类型']).sum()

        #客户购课的数
        df_cus_buy=self.cus_buy(cus_name=cus_name,crs_types=crs_types)
        #如客户购课数为空，构建一个空表。
        if df_cus_buy.empty:
            dic_empty_buy=[]
            for crs_type in crs_types:
                _dic_empty_buy={'购课类型':crs_type,'购课数量':0}
                dic_empty_buy.append(_dic_empty_buy)
            
            df_cus_buy=pd.DataFrame(dic_empty_buy)



        #客户上课的数
        #start_time用于确定教练工作日志的起始日期
        #nowtime用于
        df_cus_taken=self.cus_taken(cus_name=cus_name,crs_types=crs_types,start_time='20220601',nowtime='')
        #如客户上课数为空，构建一个空表。
        if df_cus_taken.empty:
            dic_empty_taken=[]
            for crs_type in crs_types:
                _dic_empty_taken={'课程类型':crs_type,'上课次数':0}
                dic_empty_taken.append(_dic_empty_taken)
            df_cus_taken=pd.DataFrame(dic_empty_taken)

        res=pd.merge(cus_base,df_cus_taken,how='outer')
        # print(res)
        res=pd.merge(res,df_cus_buy,left_on='课程类型',right_on='购课类型',how='left')

        res['客户名称']=cus_name
        res['剩余课时（节）'].fillna(0,inplace=True)
        res['本次剩余课时']=res['剩余课时（节）']+res['购课数量']-res['上课次数']

        res=res[['客户名称','课程类型','剩余课时（节）','购课数量','上课次数','本次剩余课时']]

        # print(res)

        return res

        # print(res)
    
    def exp_txt(self,cus_name='MH016徐颖丽',crs_type='常规私教课',crs_date='20220527',crs_time='1000-1100',ins='MHINS001陆伟杰'):
        cus_info=self.cus_info(cus_name=cus_name)
        # print(cus_name,crs_type,crs_date,crs_time,ins)
        crs_types=[crs_type]
        crs_remain=self.cal_crs_remain(cus_name=cus_name,crs_types=crs_types)
        txt_crs_remain=str(int(crs_remain[crs_remain['课程类型']==crs_type]['本次剩余课时'].tolist()[0]))

        talk_template=pd.read_excel(os.path.join(self.work_dir,'01-会员管理','工作文档','预约客户话术模板.xlsx'))
        txt_datetime='\n【'+crs_date[:4]+'年'+crs_date[4:6]+'月'+crs_date[6:]+'日】\n【'+crs_time[:2]+':'+crs_time[2:7]+':'+crs_time[7:]+'】'


        txt_talk=talk_template[talk_template['课程类型']==crs_type]['话术'].tolist()[0]
        txt_talk=txt_talk.replace('cus_name',' '+cus_info['callit']+' '+cus_info['title'])
        txt_talk=txt_talk.replace('time',' '+txt_datetime+' ')
        txt_talk=txt_talk.replace('ins','【'+ins[8:]+'】')
        txt_talk=txt_talk.replace('remain',txt_crs_remain)        
        
        return txt_talk


#####################团课及限时私教################

    def gp_cus_taken(self,cus_name='MH016徐颖丽',crs_types=['常规私教课','初级团课'],start_time='20210501',nowtime='20220604'):
        all_gp_cus_taken=self.read_excel_taken(start_time=start_time,nowtime=nowtime)
        print(all_gp_cus_taken)
        _df_gp_cus_takens=[]
        # print(all_cus_taken)
        try:
            for crs_type in crs_types:
                df_gp_cus_taken=all_gp_cus_taken[(all_gp_cus_taken['会员姓名']==cus_name) & 
                                                (all_gp_cus_taken['课程类型']==crs_type) & (all_gp_cus_taken['是否完成']=='是')]
                # if nowtime=='':
                #     df_gp_cus_taken=all_gp_cus_taken[(all_gp_cus_taken['会员姓名']==cus_name) & 
                #                                 (all_gp_cus_taken['课程类型']==crs_type) & (all_gp_cus_taken['是否完成']=='是') &
                #                                 (all_gp_cus_taken['日期']>=datetime.strptime(start_time,'%Y%m%d')) ]
                # else:
                #     df_gp_cus_taken=all_gp_cus_taken[(all_gp_cus_taken['会员姓名']==cus_name) & 
                #                                 (all_gp_cus_taken['课程类型']==crs_type) & (all_gp_cus_taken['是否完成']=='是') &
                #                                 (all_gp_cus_taken['日期']>=datetime.strptime(start_time,'%Y%m%d')) & (all_gp_cus_taken['日期']<=datetime.strptime(nowtime,'%Y%m%d'))]
                _df_gp_cus_takens.append(df_gp_cus_taken)
            df_gp_cus_takens=pd.concat(_df_gp_cus_takens)
            gp_taken=df_gp_cus_takens.groupby(['课程类型']).count().reset_index()
            gp_taken=gp_taken[['课程类型','会员姓名']]
            gp_taken.columns=['课程类型','上课次数']
            # gp_taken['上课次数']=gp_taken['上课次数'].apply(lambda x:x*-1)
        #如无上课的记录
        except Exception as e: 
            print(e)
            # gp_taken=''

        # print(gp_taken)
        return gp_taken

    def gp_cus_prd(self,cus_name='MH010苏云',start_time='20220501',nowtime='20220604'):
        s_time=datetime.strptime(start_time, '%Y%m%d')
        df_all_cus_prd=pd.read_excel(os.path.join(self.work_dir,'02-运营规划','业务数据管理','限时课程启动及结束记录表.xlsx'))
        res_df_gp_cus_prd=df_all_cus_prd[(df_all_cus_prd['会员姓名及编号']==cus_name) & (df_all_cus_prd['课程开始时间']>s_time) & (df_all_cus_prd['课程开始时间']<=datetime.strptime(nowtime,'%Y%m%d'))]
        

        return res_df_gp_cus_prd

    def gp_cus_buy(self,cus_name='MH010苏云',crs_types=['初级团课'],start_time='20220501',nowtime='20220604',data_fn='客户业务流水数据.xlsx'):
        df_gp_buy=pd.read_excel(os.path.join(self.work_dir,'02-运营规划','业务数据管理',data_fn),sheet_name='购课业务流水')
        # df_gp_buy['购课期数'].fillna(0,inplace=True)
        # df_gp_buy['购课节数'].fillna(0,inplace=True)
        df_gp_buy['购课时间'].fillna(0,inplace=True)
        df_gp_buy['备注'].fillna('无',inplace=True)

        df_gp_buy=df_gp_buy[(df_gp_buy['购课时间']>=datetime.strptime(start_time,'%Y%m%d')) & (df_gp_buy['购课时间']<=datetime.strptime(nowtime,'%Y%m%d'))]

        # print('df_gp_buy:\n',df_gp_buy)
        try:
            _df_gp_cus_buy=df_gp_buy[df_gp_buy['购课编号'].str[:-8]==cus_name]

            _gp_cus_buy=[]
            for crs_type in crs_types:
    
                if crs_type in ['常规私教课']:
                    pre_df_gp_cus_buy_jieshu=_df_gp_cus_buy[['购课编号','购课类型','购课节数','购课时间']]
                    pre_df_gp_cus_buy_jieshu.columns=['购课编号','购课类型','购课数量','购课时间']
                    _gp_cus_buy.append(pre_df_gp_cus_buy_jieshu)
                elif crs_type in ['初级团课','限时私教','踏板团课']:
                    pre_df_gp_cus_buy_qishu=_df_gp_cus_buy[['购课编号','购课类型','购课期数','购课节数','购课时间']]
                    # pre_df_gp_cus_buy_qishu.columns=['购课编号','购课类型','购课数量','购课时间']
                    _gp_cus_buy.append(pre_df_gp_cus_buy_qishu)
                else:
                    print('无效的课程类别')

            df_gp_cus_buy=pd.concat(_gp_cus_buy)
            df_gp_cus_buy.dropna(how='any',inplace=True)
            #读取的excel基数为object格式，转为int格式
            df_gp_cus_buy['购课期数']=df_gp_cus_buy['购课期数'].astype('int')
            # print(df_gp_cus_buy['购课期数'].dtypes)
            df_gp_cus_buy_cal=df_gp_cus_buy.groupby('购课类型').sum().reset_index()
        #如无购课的记录
        except Exception as e:
            print(e)
            # df_cus_buy_cal=''

        # print('test buy cal',df_cus_buy_cal)
        return df_gp_cus_buy_cal

    def gp_cal_crs_remain(self,cus_name='MH010苏云',crs_types=['常规私教课','初级团课'],data_fn='客户业务流水数据.xlsx',start_time='20220501',nowtime='20220521'):
        print(start_time,'   ',nowtime)
        #客户在start_time及nowtime之间的购课记录
        gp_cus_buy=self.gp_cus_buy(cus_name=cus_name,crs_types=crs_types,start_time=start_time,nowtime=nowtime,data_fn=data_fn)

        #客户在start_time及nowtime之间的课程开启及结束记录
        df_gp_cus_prd=self.gp_cus_prd(cus_name=cus_name,start_time=start_time,nowtime=nowtime)
        max_date=df_gp_cus_prd['课程开始时间'].max()
        latest_cus_prd=df_gp_cus_prd[df_gp_cus_prd['课程开始时间']==max_date]

        #客户在start_time及nowtime之间的上课记录,在nowtime小于课程结束时间情况下方有效,计算期间的上课次数。  
        if pd.isnull(latest_cus_prd['课程实际结束时间'].tolist()):
            end_time=latest_cus_prd['课程结束时间'].tolist()[0].strftime('%Y%m%d')
        else:
            end_time=latest_cus_prd['课程实际结束时间'].tolist()[0].strftime('%Y%m%d')
        # if datetime.strptime(nowtime,'%Y%m%d')<=latest_cus_prd['课程结束时间']  
        gp_cus_taken=self.gp_cus_taken(cus_name=cus_name,crs_types=crs_types,start_time=latest_cus_prd['课程开始时间'].tolist()[0].strftime('%Y%m%d'),nowtime=end_time)

        
        print(gp_cus_buy,'\n\n',latest_cus_prd,'\n\n',gp_cus_taken)
        
    

#############批量导出预约信息#########################

    def group_exp_txt(self,y_m='202206',crs_type='常规私教课'):
        df_schedule=pd.read_excel(self.taken_fn,sheet_name=y_m[:4]+'-'+y_m[4:])
        df_schedule.dropna(subset=['日期','工作内容'],how='any',inplace=True)
        # print(df_schedule)
        df_exp_list=df_schedule[(df_schedule['日期']==df_schedule['日期'].max()) & (df_schedule['工作内容']==crs_type)]
        
        all_ins=df_exp_list['教练'].drop_duplicates().tolist()
        crs_date=df_schedule['日期'].max().strftime('%Y%m%d')

        ins_info=pd.read_excel(os.path.join(self.work_dir,'03-教练管理','教练资料','教练信息.xlsx'),sheet_name='教练信息')
        
        txt_out=Vividict()

        all_txt=[]
        for ins in all_ins:  
            ins_code=ins_info[ins_info['姓名']==ins]['员工编号'].tolist()[0]+ins_info[ins_info['姓名']==ins]['姓名'].tolist()[0]
            df_ins_cus=df_schedule[df_schedule['教练']==ins]
            df_ins_cus=df_ins_cus.reset_index()
            #遍历行
            for index,row in df_ins_cus.iterrows():
                rec_datetime=datetime.strptime(row['日期'].strftime('%Y-%m-%d')+' '+row['时间'].strftime('%H:%M:%S'),'%Y-%m-%d %H:%M:%S')                
                end_time=rec_datetime+timedelta(hours=row['时长\n（小时）'])
                time_prd=rec_datetime.strftime('%H%M')+'-'+end_time.strftime('%H%M')  
                try:     
                    txt=self.exp_txt(cus_name=row['会员姓名'],crs_type=crs_type,crs_date=crs_date,crs_time=time_prd,ins=ins_code)
                    # print(txt)
                    txt_out[ins][index]=txt
                except Exception as e:
                    print(e)

        return txt_out

class ReadDiet:
    def __init__(self,fn_diet='D:\\Documents\\WXWork\\1688851376227744\\WeDrive\\铭湖健身工作室\\05-专业资料\\减脂饮食建议表.xlsx'):
        self.fn_diet=fn_diet
    
    def exp_diet_suggests(self):
        df=pd.read_excel(self.fn_diet,sheet_name='饮食建议')
        df.dropna(how='any',inplace=True)
        diet_suggests=df['饮食建议'].apply(lambda x:str(x).strip()).values.tolist()
        # print(diet_suggests) 

        return diet_suggests

class ReadWebData:
    def __init__(self):
        pass
        

    def deal_data(self,df):
        # print(df)        
        cus_name=df['Q4_会员姓名'].tolist()[0]
        train_date=datetime.strftime(df['Q3_训练日期'].tolist()[0],'%Y-%m-%d')
        ins_name=df['Q2_教练姓名'].tolist()[0]
        calories=df['Q7_消耗热量'].tolist()[0]
        ins_cmt=df['Q8_教练评语'].tolist()[0]
        cus_star=df['Q9_为会员本次训练情况打分_选项1'].tolist()[0]


        ########################抗阻####################
        # if pd.isna(df['Q5_抗阻训练内容_训练项目']):
        # print(df['Q5_抗阻训练内容_训练项目'])

        if df['Q5_抗阻训练内容_训练部位'].empty or pd.isna(df['Q5_抗阻训练内容_训练部位'].tolist()[0]):
            print('抗阻内容为空')
            # res_muscle={'内容':'','目标肌群':'','重量（Kg）':'','距离（m）':'','次数':''}
            # df_muscle=pd.DataFrame(res_muscle,index=[0])
            # df_muscle.fillna('',inplace=True)
            df_res_muscle=pd.DataFrame()
        else:
            if '||' in df['Q5_抗阻训练内容_训练部位'].tolist()[0]:
                #多项内容
                train_parts=df['Q5_抗阻训练内容_训练部位'].tolist()[0].split('||')
                try:
                    train_items_upper=pd.DataFrame(df['Q5_抗阻训练内容_● 训练项目（上肢及胸部）'].tolist()[0].split('||'))
                    train_items_upper_detail=pd.DataFrame(df['Q5_抗阻训练内容_---- 上肢及胸部详细内容'].tolist()[0].split('||'))

                    train_items_upper_complete=[]
                    for dt_up_num,item_up in enumerate(train_items_upper[0].tolist()):
                        dt_up=train_items_upper_detail[0].tolist()[dt_up_num]
                        if dt_up!='':
                            up_train=str(item_up)+'（'+str(dt_up)+'）'
                        else:
                            up_train=item_up
                        train_items_upper_complete.append(up_train)

                    train_items_legs=pd.DataFrame(df['Q5_抗阻训练内容_● 训练项目（下肢/背部/核心）'].tolist()[0].split('||'))
                    train_items_legs_detail=pd.DataFrame(df['Q5_抗阻训练内容_---- 下肢/背部/核心详细内容'].tolist()[0].split('||'))
                    train_items_legs_complete=[]
                    for dt_leg_num,item_leg in enumerate(train_items_legs[0].tolist()):
                        dt_leg=train_items_legs_detail[0].tolist()[dt_leg_num]
                        if dt_leg!='':
                            leg_train=str(item_leg)+'（'+str(dt_leg)+'）'
                        else:
                            leg_train=item_leg
                        train_items_legs_complete.append(leg_train)

                    # print(train_items_upper_complete,train_items_legs_complete)

                    df_train_items=pd.concat([pd.DataFrame(train_items_upper_complete),pd.DataFrame(train_items_legs_complete)],axis=1)

                    df_train_items.columns=['upper','down']                    
                    df_train_items['训练项目']=df_train_items.apply(lambda x: x['upper'] if x['down']=='' else x['down'],axis=1)
                    train_items=df_train_items['训练项目'].tolist()
                except Exception as e:
                    print('错误:',e)
                train_wts=df['Q5_抗阻训练内容_重量（Kg）'].tolist()[0].split('||')
                train_diss=df['Q5_抗阻训练内容_距离（m）'].tolist()[0].split('||')
                train_nums=df['Q5_抗阻训练内容_次数'].tolist()[0].split('||')
                train_grps=df['Q5_抗阻训练内容_组数'].tolist()[0].split('||')
                
                train_wts=['0' if x=='' else x for x in train_wts]
                train_diss=['0' if x=='' else x for x in train_diss]
                train_nums=['0' if x=='' else x for x in train_nums]
                train_grps=['0' if x=='' else x for x in train_grps]
                

                muscle_train_part,muscle_train_item,muscle_train_wt,muscle_train_num,muscle_train_dis=[],[],[],[],[]
                for itm_no,grp in enumerate(train_grps):
                    for ct in range(int(grp)):
                        muscle_train_part.append(train_parts[itm_no])
                        muscle_train_item.append(train_items[itm_no])
                        muscle_train_wt.append(train_wts[itm_no])
                        muscle_train_dis.append(train_diss[itm_no])
                        muscle_train_num.append(train_nums[itm_no])
                res_muscle={'目标肌群':muscle_train_part,'内容':muscle_train_item,'重量（Kg）':muscle_train_wt,'距离（m）':muscle_train_dis,'次数':muscle_train_num}
                # print(train_items,train_wts,train_diss,train_nums,train_grps)
                # print(res_muscle)
                df_muscle=pd.DataFrame(res_muscle)
                # df_muscle.replace('',0,inplace=True)        
                
            else:
                #一项内容
                train_part=df['Q5_抗阻训练内容_训练部位'].tolist()[0]
                if train_part in ['上肢肌群','胸部肌群']:
                    train_item=df['Q5_抗阻训练内容_● 训练项目（上肢及胸部）'].tolist()[0]
                    train_item_detail=df['Q5_抗阻训练内容_---- 上肢及胸部详细内容'].tolist()[0]
                    if train_item_detail!='':
                        train_item=str(train_item)+'（'+str(train_item_detail)+'）'
                else:
                    train_item=df['Q5_抗阻训练内容_● 训练项目（下肢/背部/核心）'].tolist()[0]
                    train_item_detail=df['Q5_抗阻训练内容_---- 下肢/背部/核心详细内容'].tolist()[0]
                    if train_item_detail!='':
                        train_item=str(train_item)+str(train_item_detail)

                train_wt=df['Q5_抗阻训练内容_重量（Kg）'].tolist()[0]
                train_dis=df['Q5_抗阻训练内容_距离（m）'].tolist()[0]
                train_num=df['Q5_抗阻训练内容_次数'].tolist()[0]
                train_grp=df['Q5_抗阻训练内容_组数'].tolist()[0]

                train_wts='0' if train_wt=='' else train_wt
                train_dis='0' if train_wt=='' else train_dis
                train_num='0' if train_wt=='' else train_num
                train_grp='0' if train_wt=='' else train_grp
                # print(train_grps)
                # print(muscle_train_item,muscle_train_wt,muscle_train_dis,muscle_train_num)

                muscle_train_part,muscle_train_item,muscle_train_wt,muscle_train_dis,muscle_train_num=[],[],[],[],[]
                for ct in range(int(train_grp)):
                    muscle_train_part.append(train_part)
                    muscle_train_item.append(train_item)
                    muscle_train_wt.append(train_wt)
                    muscle_train_dis.append(train_dis)
                    muscle_train_num.append(train_num)

                    res_muscle={'目标肌群':muscle_train_part,'内容':muscle_train_item,'重量（Kg）':muscle_train_wt,'距离（m）':muscle_train_dis,'次数':muscle_train_num}
                # print(res_muscle)
                df_muscle=pd.DataFrame(res_muscle)
                df_muscle.fillna('',inplace=True)
        
            df_muscle['时间']=train_date
            df_muscle['形式']='抗阻训练'
            df_muscle['消耗热量']=int(calories)
            df_muscle['教练姓名']=ins_name
            df_muscle['教练评语']=ins_cmt
            df_muscle['评分']=int(cus_star)
            df_muscle['有氧项目']=''
            df_muscle['有氧时长']=''

            # print(df_muscle)
            df_res_muscle=df_muscle[['时间','形式','目标肌群','有氧项目','有氧时长','内容','重量（Kg）','距离（m）','次数','消耗热量','教练姓名','教练评语','评分']]
            # print(df_res_muscle)

        ########################有氧################
        # if pd.isna(df['Q6_有氧训练内容_训练项目']):
        if df['Q6_有氧训练内容_训练项目'].empty or pd.isna(df['Q6_有氧训练内容_训练项目'].tolist()[0]):
            print('有氧训练内容为空')
            # res_muscl={'项目':'','时长（s）':''}
            # df_muscle=pd.DataFrame(res_muscle,index=[0])
            # df_muscle.fillna('',inplace=True)
            df_res_oxy=pd.DataFrame()
        else:
            if '||' in df['Q6_有氧训练内容_训练项目'].tolist()[0]:
                #多项内容
                oxy_train_items=df['Q6_有氧训练内容_训练项目'].tolist()[0].split('||')
                # oxy_train_wts=df['Q6_有氧训练内容_重量（Kg）'].tolist()[0].split('||')
                oxy_train_times=df['Q6_有氧训练内容_时间（秒）'].tolist()[0].split('||')
                oxy_train_grps=df['Q6_有氧训练内容_组数'].tolist()[0].split('||')

                oxy_train_items=['0' if x=='' else x for x in oxy_train_items]
                oxy_train_times=['0' if x=='' else x for x in oxy_train_times]

                oxy_train_grps=['1' if  np.isnan(int(x)) else x for x in oxy_train_grps]

                oxy_train_item,oxy_train_wt,oxy_train_time=[],[],[]
                for itm_no,grp in enumerate(oxy_train_grps):
                    for ct in range(int(grp)):
                        oxy_train_item.append(oxy_train_items[itm_no])
                        # oxy_train_wt.append(oxy_train_wts[itm_no])
                        oxy_train_time.append(oxy_train_times[itm_no])
                res_oxy={'项目':oxy_train_item,'时长（s）':oxy_train_time}
                # res_oxy={'项目':oxy_train_item,'重量':oxy_train_wt,'时长（s）':oxy_train_time}

                df_oxy=pd.DataFrame(res_oxy)
                # df_muscle.replace('',0,inplace=True)               

            else:
                #一项内容
                train_item=df['Q6_有氧训练内容_训练项目'].tolist()[0]
                # train_wt=df['Q6_有氧训练内容_重量（Kg）'].tolist()[0]
                train_time=df['Q6_有氧训练内容_时间（秒）'].tolist()[0]
                train_grp=df['Q6_有氧训练内容_组数'].tolist()[0]

                train_item='0' if train_item=='' else train_item
                train_time='0' if train_time=='' else train_time

                if isinstance(train_grp,str):
                    train_grp=int(train_grp)
                train_grp='1' if np.isnan(train_grp) else train_grp

             

                # print('train_grp',np.isnan(train_time),type(train_grp))

                oxy_train_item,oxy_train_wt,oxy_train_time=[],[],[]
                for ct in range(int(train_grp)):
                    oxy_train_item.append(train_item)
                    # oxy_train_wt.append(train_wt)
                    oxy_train_time.append(train_time)
                res_oxy={'项目':oxy_train_item,'时长（s）':oxy_train_time}
                # res_oxy={'项目':oxy_train_item,'重量（Kg）':oxy_train_wt,'时长（s）':oxy_train_time}
                    # print(res_oxy)
                df_oxy=pd.DataFrame(res_oxy)
                df_oxy.fillna('',inplace=True)

            df_oxy['时间']=train_date
            df_oxy['形式']='有氧训练'
            df_oxy['消耗热量']=int(calories)
            df_oxy['教练姓名']=ins_name
            df_oxy['教练评语']=ins_cmt
            df_oxy['评分']=int(cus_star)
            df_oxy['内容']=''
            df_oxy['重量（Kg）']=''
            df_oxy['距离（m）']=''
            df_oxy['次数']=''
            df_oxy['目标肌群']=''

            df_oxy.rename(columns={'项目':'有氧项目','时长（s）':'有氧时长'},inplace=True)
            # df_oxy.rename(columns={'项目':'有氧项目','时长（s）':'有氧时长','重量':'有氧重量'},inplace=True)

            #去年有氧重量项目
            df_res_oxy=df_oxy[['时间','形式','目标肌群','有氧项目','有氧时长','内容','重量（Kg）','距离（m）','次数','消耗热量','教练姓名','教练评语','评分']]
        
            # print(df_res_oxy)

        return {'df_muscle':df_res_muscle,'df_oxy':df_res_oxy}

    def exp_data_one(self,cus_name='MH003吕雅颖',date_input='20220729',fn='e:/temp/minghu/test.xlsx'):
        df_fn=pd.read_excel(fn,parse_dates=['Q3_训练日期'])
        df_match=df_fn[(df_fn['Q4_会员姓名']==cus_name) & (df_fn['Q3_训练日期']==datetime.strptime(date_input,'%Y%m%d'))]
        # df_match.iloc[:,5:]=df_match.iloc[:,5:].astype(str)
        if df_match.empty:
            # print('无数据')
            res=''
        else:
            res=self.deal_data(df_match)
        return res
        # df_row=self.df_fn[self.df_fn['']]
        # df_row=self.df_fn.iloc[0,:]
        # res=self.deal_data(df_match)
        # return res

    def body_data(self,cus_name='MH003吕雅颖',date_input='20220803',webfn='e:/temp/minghu/body.xlsx'):
        df_body=pd.read_excel(webfn,parse_dates=['Q3_日期'])
        df_body=df_body[(df_body['Q3_日期']==datetime.strptime(date_input,'%Y%m%d')) & (df_body['Q1_客户编码及姓名']==cus_name)]
        return df_body

class CusInfo:
    def __init__(self):
        pass

    def get_cus_list(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        cus_list=[]
        for fns in os.listdir(work_dir):
            if re.match(r'MH\d{3}.*.xlsx',fns):
                # print(d[:-5])
                cus_list.append(fns[:-5])
        return cus_list 

    def get_cus_body_data(self,cus_fn='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH003吕雅颖.xlsx'):
        df_body=pd.read_excel(cus_fn,sheet_name='身体数据')
        return df_body

    def get_cus_basic_data(self,cus_fn='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH003吕雅颖.xlsx'):
        df_basic=pd.read_excel(cus_fn,sheet_name='基本情况')
        return df_basic

class InsInfo:
    def __init__(self):
        pass

    def get_info(self,ins_fn='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\03-教练管理\\教练资料\\教练信息.xlsx'):
        df_ins=pd.read_excel(ins_fn,sheet_name='教练信息')
        return df_ins

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    # p=InsInfo()
    # p.get_info()

    p=ReadWebData()
    res=p.exp_data_one(cus_name='MH075黄子严',date_input='20220825',fn='e:/temp/minghu/webdl.xlsx')
    print(res)
    # res=p.body_data(cus_name='MH003吕雅颖',date_input='20220803',webfn='e:/temp/minghu/body.xlsx')
    # print(res)
    

    # print(res['df_muscle'])
    # print(res)


    # p=ReadCourses(work_dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室',dl_taken_fn='e:\\temp\\minghu\\教练工作日志.xlsx')

    # k=p.cus_buy(cus_name='MH016徐颖丽',crs_types=['常规私教课','团课'])
    # print(k)
    # res=p.cus_taken(cus_name='MH010苏云',crs_types=['常规私教课','初级团课'],start_time='20220501',nowtime='20220804')
    # print(res)
    # k=p.cal_crs_remain(cus_name='MH016徐颖丽',crs_types=['常规私教课','初级团课'])
    # print(k)
    # k=p.exp_txt(cus_name='MH016徐颖丽',crs_type='常规私教课',crs_date='20220608',crs_time='1000-1100',ins='MHINS001陆伟杰')
    # print(k)
    # p.cus_info(cus_name='MH016徐颖丽')
    # k=p.group_exp_txt(y_m='202206',crs_type='常规私教课')
    # k=p.gp_cus_taken(cus_name='MH010苏云',crs_types=['常规私教课','初级团课'],start_time='20210501',nowtime='20220804')
    # print(k)
    # p.gp_cal_crs_remain(cus_name='MH010苏云',crs_types=['初级团课'],data_fn='客户业务流水数据.xlsx',start_time='20210501',nowtime='20220531')


    # p=ReadAndExportDataNew(adj_bfr='no')
    # res=p.exp_cus_prd(cus_file_dir='E:\\temp\\minghu\\会员\\会员资料',cus='MH000唐青剑',start_time='20201201',end_time='20220523')
    # print(res['train_stat']['total_train_amt'],res['train_stat']['train_amt_month'])

    # c=cals()
    # bmr=c.bmr(sex='m',ht=170,wt=64,age=41)
    # print(bmr)
    # p=ReadDiet()
    # su=p.exp_diet_suggests()
    # print(random.choice(su))