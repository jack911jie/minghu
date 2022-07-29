import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules'))
import get_data
import numpy as np
import openpyxl
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)



class WebData:
    def __init__(self):
        pass

    def deal_data(self,dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220729'):
        web_data=get_data.ReadWebData(fn=dl_xlsx)
        res=web_data.exp_data_one(cus_name=cus_name,date_input=date_input)

        if res!='':
            mix_res=pd.concat([res['df_muscle'],res['df_oxy']])
            return mix_res
        else:
            return pd.DataFrame()
    

    def append_to_target(self,target_xlsx,dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220729'):
        webdata=self.deal_data(dl_xlsx=dl_xlsx,cus_name=cus_name,date_input=date_input)
        # print(webdata)
        try:
            webdata['有氧时长']=webdata['有氧时长'].replace('',0).astype(int).replace(0,'')
            webdata['重量（Kg）']=webdata['重量（Kg）'].replace('',0).astype(int).replace(0,'')
            webdata['距离（m）']=webdata['距离（m）'].replace('',0).astype(int).replace(0,'')
            webdata['次数']=webdata['次数'].replace('',0).astype(int).replace(0,'')

        except Exception as e:
            pass

        if webdata.empty:
            print('无数据，未追加数据。')            
        else:
            old=pd.read_excel(target_xlsx,sheet_name='训练情况')
            book=openpyxl.load_workbook(target_xlsx)
            writer=pd.ExcelWriter(target_xlsx,engine='openpyxl')
            writer.book=book
            writer.sheets=dict((ws.title,ws) for ws in book.worksheets)            

            old_rows=old.shape[0]
            webdata.to_excel(writer,sheet_name='训练情况',startrow=old_rows+1,index=False,header=False)
            writer.save()
            print('完成')


            


if __name__=='__main__':
    p=WebData()
    # p.deal_data(dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220720')
    p.append_to_target(target_xlsx='e:/temp/minghu/MH003吕雅颖.xlsx',dl_xlsx='e:/temp/minghu/test.xlsx',cus_name='MH003吕雅颖',date_input='20220723')

