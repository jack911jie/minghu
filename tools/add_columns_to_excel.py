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
        sht=wb.sheets['基本情况']
        rng='f1'        
        sht.range(rng).value='客户来源'
        sht.range(rng).api.Font.Bold = True
        sht.range(rng).api.Font.Name = '黑体'
        sht.range(rng).api.Font.ColorIndex=49
        sht.range(rng).api.Interior.ColorIndex = 34

        # 设置字体为黑体、加粗
        cell_range=sht.range('A1:F2')
        # cell_range.api.Font.Bold = True
        # cell_range.api.Font.Name = '黑体'

        # 添加框线
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeTop).LineStyle = xw.constants.LineStyle.xlContinuous
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).LineStyle = xw.constants.LineStyle.xlContinuous
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).LineStyle = xw.constants.LineStyle.xlContinuous
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeRight).LineStyle = xw.constants.LineStyle.xlContinuous

        cell_range=sht.range('A1:F1')
        # 设置边框线宽度为粗线
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
        cell_range.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3

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
    # p.add_column('E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料\\MH207杨薇.xlsm')
    
