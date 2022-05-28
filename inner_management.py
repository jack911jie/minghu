import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import readconfig
import get_data
import re
import pandas as pd

class MinghuTable:
    def __init__(self,wecomid='1688851376239499'):
        config=readconfig.exp_json(os.path.join(os.path.dirname(__file__),'configs','minghu_dir.config'),wecomid_replace='yes',wecomid_pair=['$wecomid$',wecomid])
        self.cus_dir=config['会员资料文件夹']

    def exp_cus_info(self):
        cus_list=[]
        for fn in os.listdir(self.cus_dir):
            if re.match(r'^MH\d{3}.*.xlsx',fn):
                cus_list.append(os.path.join(self.cus_dir,fn))
        
        infos=[]
        for cus_info_fn in cus_list:
            try:
                # print(cus_info_fn)
                df_info=pd.read_excel(cus_info_fn,sheet_name='基本情况')
                infos.append(df_info)
            except Exception as e:
                print(cus_info_fn,e)
        df_infos=pd.concat(infos)
        df_infos['出生年月'].fillna(0,inplace=True)
        df_infos['出生年月']=df_infos['出生年月'].astype(int)
        save_name=os.path.join(self.cus_dir,'会员信息汇总.xlsx')
        df_infos.to_excel(save_name)
        os.startfile(save_name)

    def exp_cus_train(self,cus_file_dir='E:\\temp\\minghu\\会员\\会员资料',start_time='20211201',end_time='20220523'):
        cus_list=[]
        for fn in os.listdir(self.cus_dir):
            if re.match(r'^MH\d{3}.*.xlsx',fn):
                cus_list.append(os.path.join(self.cus_dir,fn))
        

        for cus_info_fn in cus_list:
            try:
                # print(cus_info_fn)
                res=get_data.ReadAndExportDataNew(adj_bfr='no').exp_cus_prd(cus_file_dir=cus_file_dir,cus=cus_info_fn[:-5],start_time=start_time,end_time=end_time)
                print(res['train_stat']['total_train_amt'])

            except Exception as e:
                print(cus_info_fn,e)


if __name__=='__main__':
    p=MinghuTable(wecomid='1688851376239499')
    # p.exp_cus_info()
    p.exp_cus_train(cus_file_dir='E:\\temp\\minghu\\会员\\会员资料',start_time='20211201',end_time='20220523')
