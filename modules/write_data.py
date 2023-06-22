import pandas as pd
import numpy as np
import openpyxl

class WriteData:
    def write_to_xlsx(self,input_dataframe,output_xlsx,sheet_name,parse_date_col_name='时间'):
        if input_dataframe.shape[0]==0:
            return '无数据追加'
        else:
            if parse_date_col_name:
                old=pd.read_excel(output_xlsx,sheet_name=sheet_name,parse_dates=[parse_date_col_name])
            else:
                old=pd.read_excel(output_xlsx,sheet_name=sheet_name)
            book=openpyxl.load_workbook(output_xlsx,keep_vba=True)
            writer=pd.ExcelWriter(output_xlsx,engine='openpyxl')
            writer.book=book
            writer.sheets=dict((ws.title,ws) for ws in book.worksheets)         
            old_rows=old.shape[0]
            input_dataframe.to_excel(writer,sheet_name=sheet_name,startrow=old_rows+1,index=False,header=False)
            writer.save()
            log_txt=str(input_dataframe.shape[0])+' 条数据追加完成，行号：'+str(old_rows+2)+'-'+str(old_rows+input_dataframe.shape[0]+1)

            return log_txt
        

    def verify_data(self,df_old,df_new,cols):
        # print(df_old,'\n',df_new,'\n')
        df_diff=pd.concat([df_new,df_old,df_old]).drop_duplicates(subset=cols,keep=False,inplace=False)
        # print(df_diff)
        
        return df_diff