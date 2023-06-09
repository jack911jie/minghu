import os
import sys
import pandas as pd
import  openpyxl


class WriteData:
    def write_to_xlsx(self,input_dataframe,output_xlsx,sheet_name,parse_date_col_name='时间'):
        old=pd.read_excel(output_xlsx,sheet_name=sheet_name,parse_dates=[parse_date_col_name])
        book=openpyxl.load_workbook(output_xlsx)
        writer=pd.ExcelWriter(output_xlsx,engine='openpyxl')
        writer.book=book
        writer.sheets=dict((ws.title,ws) for ws in book.worksheets)         
        old_rows=old.shape[0]
        input_dataframe.to_excel(writer,sheet_name=sheet_name,startrow=old_rows+1,index=False,header=False)
        writer.save()
        log_txt=str(input_dataframe.shape[0])+' 条数据追加完成，行号：'+str(old_rows+2)+'-'+str(old_rows+input_dataframe.shape[0]+1)

        return log_txt
 
class Xlsx(WriteData):
    def __init__(self):
        pass

    def del_sht(self,sht_name='上课记录',xls='e:/temp/minghu/000.xlsx'):
        try:
            # 删除空表头Sheet1
            # 载入工作簿
            workbook = openpyxl.load_workbook(xls)
            #
            # 删除目标Sheet
            worksheet = workbook[sht_name]
            workbook.remove(worksheet)

            # 保存已做删除处理的工作簿
            workbook.save(xls)
            print('原 {} 已删除'.format(sht_name))
        except:
            print('无 {} 表',format(sht_name))

    def add_sht(self,cls=['日期','时间','时长（小时）','工作内容','会员姓名','教练'],sht_name='上课记录',xls='e:/temp/minghu/000.xlsx'):
        df_empty=pd.DataFrame(columns=cls)
        writer = pd.ExcelWriter(xls, mode='a',engine="openpyxl")
        df_empty.to_excel(writer, sheet_name=sht_name,index=False)
        writer.save()
        print('增加 {} 表完成'.format(sht_name))

class ReadData:
    def __init__(self):
        pass

    def read_cus_buy(self,xls='E:\\temp\\minghu\\10-工作室管理文件\\财务流水\\客户业务流水数据.xlsx',sht_name='购课财务流水'):
        df_cus_buy=pd.read_excel(xls,sheet_name=sht_name)
        df_cus_buy['客户编码及姓名']=df_cus_buy['编码'].apply(lambda x: x[:len(x)-8])

        # 日期	购课类型	购课节数	购课时长（天）	限时课程起始日	限时课程结束日	限时课程实际结束日	购课金额	备注

        df_new=pd.DataFrame()
        df_new['客户编码及姓名']=df_cus_buy['客户编码及姓名']
        df_new['收款编码']=df_cus_buy['编码']
        df_new['日期']=df_cus_buy['收款日期']
        df_new['购课类型']=df_cus_buy['购课类型']
        df_new['购课节数']=''
        df_new['购课时长（天）']=''
        df_new['限时课程起始日']=''
        df_new['限时课程结束日']=''
        df_new['限时课程实际结束日']=''
        df_new['购课金额']=df_cus_buy['实收金额']
        df_new['收款人']=df_cus_buy['收款人']
        df_new['备注']=df_cus_buy['备注']


        print(df_new)




if __name__=='__main__':
    p=ReadData()
    p.read_cus_buy()




    # xls='e:/temp/minghu/000.xlsx'
    # cls=['日期','时间','时长（小时）','工作内容','会员姓名','教练']
    # sht_name='上课记录'

    # p=Xlsx()
    # p.del_sht(sht_name=sht_name)    
    # p.add_sht(cls=cls,sht_name=sht_name,xls=xls)


    # cls=['已上课时数']
    # sht_name='修正参数'
    # p.del_sht(sht_name=sht_name)    
    # p.add_sht(cls=cls,sht_name=sht_name,xls=xls)