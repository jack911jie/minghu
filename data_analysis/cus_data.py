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
        for fn in tqdm(os.listdir(dir)):
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

        
                



if __name__=='__main__':
    p=CusData()
    # res=p.get_cus_buy()
    # res=p.batch_get_cus_buy()
    p.exp_all_cus_buy(input_dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',output_dir='e:\\temp\\minghu')
