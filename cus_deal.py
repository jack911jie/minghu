import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import write_data
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐

from datetime import datetime
import  openpyxl
import re


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
            print('无 {} 表'.format(sht_name))

    def add_sht(self,cls=['日期','时间','时长（小时）','工作内容','会员姓名','教练'],sht_name='上课记录',xls='e:/temp/minghu/000.xlsx'):
        df_empty=pd.DataFrame(columns=cls)
        writer = pd.ExcelWriter(xls, mode='a',engine="openpyxl")
        df_empty.to_excel(writer, sheet_name=sht_name,index=False)
        writer.save()
        print('增加 {} 表完成'.format(sht_name))

    def xlsx_deal(self,xls='e:/temp/minghu/000.xlsx'):        
        cls=['日期','时间','时长（小时）','课程类型','教练','备注']
        sht_name='上课记录'
        
        self.del_sht(sht_name=sht_name,xls=xls)    
        self.add_sht(cls=cls,sht_name=sht_name,xls=xls)

        cls=['收款日期','收款编码','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']
        sht_name='购课表'
        self.del_sht(sht_name=sht_name,xls=xls)
        self.add_sht(cls=cls,sht_name=sht_name,xls=xls)


        cls=['已上课时数','剩余天数','参数3','参数4']
        sht_name='修正参数'
        self.del_sht(sht_name=sht_name,xls=xls)    
        self.add_sht(cls=cls,sht_name=sht_name,xls=xls)

    def batch_dir_deal(self,dir):
        for fn in os.listdir(dir):
            if re.match(r'MH\d{3}.*.xlsx$',fn):
                print('\n',fn)
                fn=os.path.join(dir,fn)
                self.xlsx_deal(xls=fn)
    


class DealData:
    def __init__(self):
        self.each_cus_buy=self.read_cus_buy()
        self.each_cus_crs=self.read_cus_crs()

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
        df_new['应收金额']=df_cus_buy['应收金额']
        df_new['实收金额']=df_cus_buy['实收金额']
        df_new['收款人']=df_cus_buy['收款人']
        df_new['备注']=df_cus_buy['备注']


        df_gp_cusname=df_new.groupby('客户编码及姓名')
        lst=list(df_gp_cusname)
        dic_cus={}
        for cusname in lst:
            dic_cus[cusname[0]]=cusname[1]
        
        return dic_cus


    def write_cus_buy_xlsx(self,xls='e:/temp/minghu/MH017李俊娴.xlsx'):
        dat=self.each_cus_buy

        cus_name=xls.split('\\')[-1].split('.')[0]
        cus_dat=dat[cus_name]

        df_input=pd.DataFrame()
        df_input['收款日期']=cus_dat['日期']
        df_input['收款编码']=cus_dat['收款编码']
        df_input['购课类型']=cus_dat['购课类型']
        df_input['购课节数']=''
        df_input['购课时长（天）']=''
        df_input['应收金额']=cus_dat['应收金额']
        df_input['实收金额']=cus_dat['实收金额']
        df_input['收款人']=cus_dat['收款人']
        df_input['收入类别']=''
        df_input['客户来源']=''
        df_input['备注']=''

        #删除重复
        df_old=pd.read_excel(xls,sheet_name='购课表').dropna(how='any',subset=['收款编码'])
        df_new=write_data.WriteData().verify_data(df_old=df_old,df_new=df_input,cols=['收款编码','收款日期'])
        #写入不重复的数据
        res=write_data.WriteData().write_to_xlsx(input_dataframe=df_new,output_xlsx=xls,sheet_name='购课表',parse_date_col_name='收款日期')
        print(res)

    def batch_write_cus_buy_dir(self,dir):
        for fn in os.listdir(dir):
            if re.match(r'MH\d{3}.*.xlsx$',fn):
                try:
                    print('\n',fn)
                    fn=os.path.join(dir,fn)

                    self.write_cus_buy_xlsx(xls=fn)
                except Exception as err:
                    print('错误',err)

    def read_cus_crs(self,xls='E:\\temp\\minghu\\教练工作日志合并.xlsx'):
        df_cus_crs=pd.read_excel(xls)

        df_new_crs=pd.DataFrame()
        df_new_crs['日期']=df_cus_crs['日期']

        df_gp_crs=df_cus_crs.groupby('会员姓名')
        lst_cusname=list(df_gp_crs)
        dic_crs={}
        for cusname in lst_cusname:
            dic_crs[cusname[0]]=cusname[1]

        return dic_crs

    def write_cus_crs(self,xls='e:\\temp\\minghu\\MH017李俊娴.xlsx'):
        df_crs=self.each_cus_crs
        cus_name=xls.split('\\')[-1].split('.')[0]
        df_cus_crs=df_crs[cus_name]
        df_cus_crs_yes=df_cus_crs[df_cus_crs['是否\n完成']=='是']

        df_new=pd.DataFrame()
        df_new['日期']=df_cus_crs_yes['日期']
        df_new['日期']=df_new['日期'].apply(lambda x: int(datetime.strftime(x,'%Y%m%d')))
        df_new['时间']=df_cus_crs_yes['时间']
        df_new['时长（小时）']=df_cus_crs_yes['时长\n（小时）']
        df_new['课程类型']=df_cus_crs_yes['工作内容']
        df_new['教练']=df_cus_crs_yes['教练']
        df_new['备注']=df_cus_crs_yes['备注']


         #删除重复
        df_crs_old=pd.read_excel(xls,sheet_name='上课记录').dropna(how='any',subset=['日期'])
        df_crs_new=write_data.WriteData().verify_data(df_old=df_crs_old,df_new=df_new,cols=['日期'])

        #写入不重复的数据
        res=write_data.WriteData().write_to_xlsx(input_dataframe=df_crs_new,output_xlsx=xls,sheet_name='上课记录',parse_date_col_name='时间')
        print(res)


    def batch_write_cus_crs_dir(self,dir):
        for fn in os.listdir(dir):
            if re.match(r'MH\d{3}.*.xlsx$',fn):
                try:
                    print('\n',fn)
                    fn=os.path.join(dir,fn)

                    self.write_cus_crs(xls=fn)
                except Exception as err:
                    print('错误',err)

# class WriteData:
#     def write_to_xlsx(self,input_dataframe,output_xlsx,sheet_name,parse_date_col_name='时间'):

if __name__=='__main__':

    #批处理修改个人档案中的上课记录、购课表、修正参数等子表

    # p=Xlsx()
    # # p.xlsx_deal(xls='E:\\temp\\minghu\\01-会员管理\\会员资料\\MH003吕雅颖.xlsx')
    # p.batch_dir_deal(dir='E:\\temp\\minghu\\01-会员管理\\会员资料')

    #------------------------------------------------------------------

    #将【财务流水】中的购课信息写入个人档案中
    p=DealData()
    # p.read_cus_buy()
    # p.write_cus_buy_xlsx(xls='E:\\temp\\minghu\\01-会员管理\\会员资料\\MH017李俊娴.xlsx')
    # p.write_cus_buy_xlsx(xls='E:\\temp\\minghu\\MH017李俊娴.xlsx')
    # p.batch_write_cus_buy_dir(dir='E:\\temp\\minghu\\01-会员管理\\会员资料')


    # p.read_cus_crs()
    # p.write_cus_crs(xls='e:\\temp\\minghu\\MH017李俊娴.xlsx')
    # p.batch_write_cus_crs_dir(dir='E:\\temp\\minghu\\01-会员管理\\会员资料')

    
    


    