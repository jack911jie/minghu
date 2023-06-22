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
        


if __name__=='__main__':
    p=XlsxToXlsm()
    # fnlist=p.get_list(input_dir='E:\\temp\\minghu\\01-会员管理\\会员资料')
    # p.new_xlsm(template='E:\\temp\\minghu\\01-会员管理\\模板.xlsm',fn_list=fnlist,out_dir='E:\\temp\\minghu\\xlsm')
    # p.copy_data(xlsx='E:\\temp\\minghu\\01-会员管理\\会员资料\\MH016徐颖丽.xlsx',xlsm='E:\\temp\\minghu\\xlsm\\MH016徐颖丽.xlsm')

    # df_src=pd.read_excel('E:\\temp\\minghu\\01-会员管理\\会员资料\\MH016徐颖丽.xlsx',sheet_name='基本情况')
    # res=write_data.WriteData().write_to_xlsx(input_dataframe=df_src,output_xlsx='E:\\temp\\minghu\\xlsm\\MH016徐颖丽.xlsm',
    #             sheet_name='基本情况',parse_date_col_name='')
    # k=pd.read_excel('E:\\temp\\minghu\\xlsm\\MH016徐颖丽.xlsm',sheet_name='基本情况')

    # print(k)
    p.batch_copy_data(xlsx_dir='E:\\temp\\minghu\\01-会员管理\\会员资料',xlsm_dir='E:\\temp\\minghu\\xlsm')