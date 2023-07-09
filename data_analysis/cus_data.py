import os
import sys
import re
from datetime import datetime
from tqdm import tqdm
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐


class CusData:
    def __init__(self):
        pass

    def get_cus_buy(self,fn='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH207杨薇.xlsm'):
        df_cus_buy=pd.read_excel(fn,sheet_name='购课表')
        return df_cus_buy

    def batch_get_cus_buy(self,dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        dfs=[]    
        pbar=tqdm(os.listdir(dir))    
        for fn in pbar:
            pbar.set_description("正在读取客户购课信息")
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                realfn=os.path.join(dir,fn) 
                df=self.get_cus_buy(realfn)
                dfs.append(df)
        
        all_df_buy=pd.concat(dfs)
        
        return all_df_buy

    def exp_all_cus_buy(self,input_dir,output_dir):
        # print('\n正在抽取购课数据……',end='')
        all_df_buy=self.batch_get_cus_buy(dir=input_dir)
        # print('完成')
        if all_df_buy.shape[0]>0:
            print('\n正在提取及合并数据……',end='')
            fn=os.path.join(output_dir,datetime.now().strftime('%Y%m%d%H%M')+'-购课数据.xlsx')
            all_df_buy.dropna(how='any',subset=['购课编码'],inplace=True)
            all_df_buy.to_excel(fn,sheet_name='购课表',index=False)
            print('完成')

        return all_df_buy

    def formal_cls_taken(self,fn='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH120肖婕.xlsm'):
        df_formal_tk=pd.read_excel(fn,sheet_name='上课记录')
        df_formal_tk['会员姓名']=fn.split('\\')[-1].split('.')[0]
        df_formal_tk=df_formal_tk[['日期', '时间', '时长（小时）', '课程类型', '会员姓名', '教练', '备注']]
        return df_formal_tk

    def batch_fomral_cls_taken(self,dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx'):

        dfs_tk=[]
        pbar=tqdm(os.listdir(dir))
        for fn in pbar:
            pbar.set_description("正在读取教练上课记录")
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                df_tk=self.formal_cls_taken(os.path.join(dir,fn))
                dfs_tk.append(df_tk)
        
        df_taken=pd.concat(dfs_tk)
        df_taken.dropna(how='any',subset=['日期'],inplace=True)
        df_taken.to_excel(out_fn,index=False,sheet_name='教练上课记录表')
        df_taken['时长（小时）']=1
        print('完成')
        return df_taken




    def trial_cls(self,fn='E:\\temp\\minghu\\体验课上课记录表-2023.xlsx'):
        df_trial=pd.read_excel(fn,sheet_name='体验课上课记录表')
        return df_trial

    def all_trial_cls(self,dir='E:\\temp\\minghu',old_trial_fn='E:\\temp\minghu\\既往体验课记录.xlsx'):
        trial=[]
        for fn in os.listdir(dir):
            if re.match(r'^体验课上课记录表-\d{4}.xlsx$',fn):
                df_trial=pd.read_excel(os.path.join(dir,fn))
                df_trial=df_trial.iloc[:,1:]
                df_trial.dropna(how='all',inplace=True)
                trial.append(df_trial)
        
        df_old_trial=pd.read_excel(old_trial_fn)
        df_old_trial.dropna(how='all',inplace=True)
        trial.append(df_old_trial)


        df_all_trial=pd.concat(trial)
        return df_all_trial


if __name__=='__main__':
    p=CusData()
    p.batch_fomral_cls_taken(dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx')

    # res=p.get_cus_buy()
    # res=p.batch_get_cus_buy()
    # p.exp_all_cus_buy(input_dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',output_dir='e:\\temp\\minghu')
    # p.formal_cls_taken()
    # res=p.all_trial_cls()
    # print(res)
    # res.to_excel('E:\\temp\\minghu\\所有体验课合并.xlsx',sheet_name='所有体验课数据',index=False)
  