from ast import parse
import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules'))
import get_data
import write_data
import numpy as np
from datetime import datetime
import openpyxl
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)



class WebData:
    def __init__(self):
        pass

    def deal_train_data(self,dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220729'):
        web_data=get_data.ReadWebData()
        res=web_data.exp_data_one(cus_name=cus_name,date_input=date_input,fn=dl_xlsx)


        if res!='':

            if res['df_muscle'].empty:
                mix_res=res['df_oxy']
            elif res['df_oxy'].empty:
                mix_res=res['df_muscle']
            else:
                mix_res=pd.concat([res['df_muscle'],res['df_oxy']])
            return mix_res
        else:
            return pd.DataFrame()
    

    def append_to_train_sheet(self,target_xlsx,target_sheet,dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220729'):
        webdata=self.deal_train_data(dl_xlsx=dl_xlsx,cus_name=cus_name,date_input=date_input)
        # print(webdata)

        # if not webdata.empty:
        # webdata=webdata[['时间','形式','目标肌群','有氧项目','有氧时长','内容','重量（Kg）','距离（m）','次数','消耗热量','教练姓名','教练评语','评分']]
        

        if webdata.empty:
            # print('无数据，未追加数据。')        
            return '无数据，未追加数据。'
        else:
            webdata=webdata[['时间','形式','目标肌群','有氧项目','有氧时长','内容','重量（Kg）','距离（m）','次数','消耗热量','教练姓名','教练评语']]
            # print(webdata)
            try:
                webdata['有氧时长']=webdata['有氧时长'].replace('',0).astype(float).replace(0,'')
                webdata['重量（Kg）']=webdata['重量（Kg）'].replace('',0).astype(float).replace(0,'')
                webdata['距离（m）']=webdata['距离（m）'].replace('',0).astype(float).replace(0,'')
                webdata['次数']=webdata['次数'].replace('',0).astype(float).replace(0,'')

            except Exception as e:
                print(e)
            new_data=write_data.WriteData()
            put_data_res=new_data.write_to_xlsx(input_dataframe=webdata,output_xlsx=target_xlsx,sheet_name=target_sheet,parse_date_col_name='时间')
            # print(cus_name+' '+put_data_res+'\n')
            return cus_name+' '+put_data_res+'\n'

    def batch_deal_train_data(self,web_file,target_dir,date_input):
        df_web=pd.read_excel(web_file,parse_dates=['Q3_训练日期'])
        df_web=df_web[df_web['Q3_训练日期']==datetime.strptime(date_input,'%Y%m%d')]
        cus_names=df_web['Q4_会员姓名'].tolist()
        if len(cus_names)==0:
            print('该日期下无数据待追加')
        else:
            for cus_name in cus_names:
                print('正在处理 '+cus_name+' 的数据。。。')
                target_xlsx=os.path.join(target_dir,cus_name+'.xlsx')
                deal_res=self.append_to_train_sheet(target_xlsx=target_xlsx,target_sheet='训练情况',dl_xlsx=web_file,cus_name=cus_name,date_input=date_input)
                print(deal_res+'\n')

    def append_to_body_sheet(self,cus_name='MH003吕雅颖',date_input='20220803',webfn='e:/temp/minghu/body.xlsx',target_xlsx='e:/temp/minghu/MH003吕雅颖.xlsx',target_sheet='身体数据'):
        df_web_bodydata=get_data.ReadWebData().body_data(cus_name=cus_name,date_input=date_input,webfn=webfn)
        if df_web_bodydata.empty:
            return '无数据，未追加数据。'
        else:
            df_web_bodydata=df_web_bodydata[['Q3_日期','Q4_身高（cm）','Q5_体重（Kg）','Q6_体脂率（%）','Q7_胸围（cm）','Q8_左臂围（cm）','Q9_右臂围（cm）',
                                                'Q10_腰围（cm）','Q11_臀围（cm）','Q12_左腿围（cm）','Q13_右腿围（cm）','Q14_左小腿围（cm）','Q15_右小腿围（cm）',
                                                'Q16_心肺','Q17_平衡','Q18_力量','Q19_柔韧性','Q20_核心']]

            df_web_bodydata=df_web_bodydata.replace(0,'')
            bodydata=write_data.WriteData()
            res_put_body_data=bodydata.write_to_xlsx(input_dataframe=df_web_bodydata,output_xlsx=target_xlsx,sheet_name=target_sheet,parse_date_col_name='时间')
            return cus_name+' 围度及体适能 '+res_put_body_data

    def batch_deal_bodydata_on_date(self,web_file='e:/temp/minghu/body.xlsx',target_dir='e:/temp/minghu',date_input='20220803'):
        df_body=pd.read_excel(web_file,parse_dates=['Q3_日期'])
        df_ondate=df_body[df_body['Q3_日期']==datetime.strptime(date_input,'%Y%m%d')]
        cus_names=df_ondate['Q1_客户编码及姓名'].tolist()
        if len(cus_names)==0:
            print('该日期下无数据待追加')
        else:
            for cus_name in cus_names:
                print('正在处理 '+cus_name+' 的数据。。。')
                target_xlsx=os.path.join(target_dir,cus_name+'.xlsx')
                deal_body_res=self.append_to_body_sheet(cus_name=cus_name,date_input=date_input,webfn=web_file,target_xlsx=target_xlsx,target_sheet='身体数据')
                print(deal_body_res+'\n')
        


if __name__=='__main__':
    p=WebData()
    # p.deal_data(dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220720')
    # p.append_to_target(target_xlsx='e:/temp/minghu/MH011小韦.xlsx',target_sheet='训练情况',
    #                     dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH011小韦',date_input='20220729')
    # p.batch_deal_train_data(web_file='E:/temp/minghu/20220802.xlsx',target_dir='e:/temp/minghu',date_input='20220802')

    # p.append_to_body_sheet(cus_name='MH003吕雅颖',date_input='20220803',webfn='e:/temp/minghu/body.xlsx')
    p.batch_deal_bodydata_on_date(web_file='e:/temp/minghu/body.xlsx',target_dir='e:/temp/minghu',date_input='20220803')
