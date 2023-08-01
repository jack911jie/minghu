import os
import sys
import pandas as pd
import  xlwings as xw
import re
from tqdm import tqdm

class XlsxFiles:
    def add_column(self,fn): 
        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['辅助表']
        rng='I1'
        sht.range(rng).value='未开课的购课编码'

        wb.save(fn)
        wb.close()
        app.quit()

    def batch_deal_add_columns(self,dir):
        pbar=tqdm(os.listdir(dir))
        for fn in pbar:
            pbar.set_description("正在处理 ")
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                dir_fn=os.path.join(dir,fn)
                self.add_column(dir_fn)
            

if __name__=='__main__':
    dir='E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料'
    p=XlsxFiles()
    p.batch_deal_add_columns(dir=dir)
    # p.add_column('E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料\\MH016徐颖丽.xlsm')
    
