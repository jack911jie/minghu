import os
import sys
sys.path.append('d:\\py\\minghu\\modules')
import write_data
import re
import pandas as pd
import numpy as np
import re
import shutil


class XlsxToXlsm:
    def get_list(self,input_dir):
        fn_list=[]
        for fn in os.listdir(input_dir):
            if re.match(r'MH\d{3}.*.xlsx$',fn):
                fn_list.append(fn)
        
        return fn_list
    
    def new_xlsm(self,template,fn_list,out_dir):
        # shts=['基本情况','身体数据','训练情况','购课表','上课记录','限时课程记录']
        for fn in fn_list:
            filename=fn.split('\\')[-1].split('.')[0]
            print('正在生成 {} 的xlsm文件'.format(filename))
            shutil.copy(template,os.path.join(out_dir,filename+'.xlsm'))
        print('完成')

    def copy_data(self,xlsx,xlsm):
        shts=['基本情况','身体数据','训练情况']
        date_dic={
            '基本情况':'',
            '身体数据':'日期',
            '训练情况':'日期'
        }
        for sht in shts:
            try:
                df_src=pd.read_excel(xlsx,sheet_name=sht)
                res=write_data.WriteData().write_to_xlsx(input_dataframe=df_src,output_xlsx=xlsm,
                sheet_name=sht,parse_date_col_name=date_dic[sht])

            except Exception as e:
                # print(e)

                pass

    def batch_copy_data(self,xlsx_dir,xlsm_dir):
        xlsx_list=self.get_list(input_dir=xlsx_dir)
        for xlsx in xlsx_list:
            xlsx_name=os.path.join(xlsx_dir,xlsx)
            xlsm_name=os.path.join(xlsm_dir,xlsx.split('\\')[-1].split('.')[0]+'.xlsm')
            self.copy_data(xlsx=xlsx_name,xlsm=xlsm_name)
        print('done')

    def read_xlsm(self,fn,shtname):
        
        df=pd.read_excel(fn,sheet_name=shtname)
        # df.to_excel('C:\\Users\\admin\\Desktop\\模板2eeee333.xlsm')
        return df

    def write_xlsm(self,df,out_name,sheet_name,parse_date_col_name=''):
        # fn='E:\\temp\\minghu\\01-会员管理\\会员资料\\MH120肖婕.xlsx'
        pass
    

class CopyRecs:
    def __init__(self,buy_fn,taken_fn):
        self.df_buys=self.read_buys(buy_fn=buy_fn)
        self.df_taken=self.read_takens(taken_fn=taken_fn)

    def read_buys(self,buy_fn):
        df_buy=pd.read_excel(buy_fn,sheet_name='购课财务流水')
        df_buy['会员编码及姓名']=df_buy['编码'].apply(lambda x: x[:-8])
        return df_buy

    def read_takens(self,taken_fn):
        df_taken=pd.read_excel(taken_fn)
        return df_taken

    def append_buy_to_fn(self,fn='E:\\temp\\minghu\\xlsm\\MH016徐颖丽.xlsm'):
        #'收款日期','购课编码','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注'
        # df_fn=pd.read_excel(fn,sheet_name='购课表')
        cus_name=fn.split('\\')[-1].split('.')[0]
        df_had=self.df_buys[self.df_buys['会员编码及姓名']==cus_name]
        if df_had.shape[0]>0:
            df_had=df_had.copy()
            df_had.rename(columns={'编码':'购课编码','购课类别':'收入类别'},inplace=True)
            df_had.loc[:,'购课节数']=''
            df_had.loc[:,'购课时长（天）']=''
            df_input=df_had[['收款日期','购课编码','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']]

            append_log=write_data.WriteData().write_to_xlsx(input_dataframe=df_input,output_xlsx=fn,sheet_name='购课表',parse_date_col_name='收款日期')
            # print(append_log)

    def append_taken_to_fn(self,fn='E:\\temp\\minghu\\xlsm\\MH016徐颖丽.xlsm'):
        cus_name=fn.split('\\')[-1].split('.')[0]
        df_cus_taken=self.df_taken[(self.df_taken['会员姓名']==cus_name) & (self.df_taken['是否\n完成']=='是') ]
        if df_cus_taken.shape[0]>0:
            df_input=df_cus_taken.copy()
            df_input.rename(columns={'工作内容':'课程类型'},inplace=True)
            # print(df_input)
            df_input=df_input[['日期','时间','时长\n（小时）','课程类型','教练','备注']]
            apd_log=write_data.WriteData().write_to_xlsx(input_dataframe=df_input,output_xlsx=fn,sheet_name='上课记录',parse_date_col_name='日期')
            print(apd_log)

    def batch_append_buy(self,dir):
        for fn in os.list(dir):
            print('正在添加 {} 的购课记录'.format(fn))
            filename=os.path.join(dir,fn)
            self.append_buy_to_fn(fn=filename)
        print('完成')

    def batch_append_taken(self,dir):
        for fn in os.list(dir):
            print('正在添加 {} 的上课记录'.format(fn))
            filename=os.path.join(dir,fn)
            self.append_taken_to_fn(fn=filename)
        print('完成')

    def batch_transfer_short_date(self,dir):
        shts={'身体数据':['日期'],'训练情况':['日期'],'购课表':['收款日期'],'上课记录':['日期'],
                            '限时课程记录':['限时课程起始日','限时课程结束日','限时课程实际结束日']}

        for fn in os.listdir(dir):
            if re.match(r'MH\d{3}.*.xlsm$',fn):
                filename=os.path.join(dir,fn)
                for key in shts.keys():
                    for itm in shts[key]:
                        style_log=write_data.WriteData().convert_column_to_date_format(file_path=filename,sheet_name=key,column_name=itm)
                

if __name__=='__main__':
    #一、从xlsx生成xlsm至目录
    # p=XlsxToXlsm()
    # fnlist=p.get_list(input_dir='E:\\temp\\minghu\\01-会员管理\\会员资料')
    # p.new_xlsm(template='E:\\temp\\minghu\\01-会员管理\\模板.xlsm',fn_list=fnlist,out_dir='E:\\temp\\minghu\\xlsm')

    #二、批量将xlsx数据copy到新的xlsm中
    # p.batch_copy_data(xlsx_dir='E:\\temp\\minghu\\01-会员管理\\会员资料',xlsm_dir='E:\\temp\\minghu\\xlsm')

    #三、复制既往的上课及购课数据
    q=CopyRecs(buy_fn='E:\\temp\\minghu\\客户业务流水数据.xlsx',taken_fn='E:\\temp\\minghu\\教练工作日志合并.xlsx')
    q.batch_append_buy(dir='E:\\temp\\minghu\\xlsm')
    q.batch_append_taken(dir='E:\\temp\\minghu\\xlsm')

    #四、批量将每个文件的日期格式改为矩日期
    # q.batch_transfer_short_date(dir='E:\\temp\\minghu\\xlsm')