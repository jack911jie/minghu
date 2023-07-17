import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'WeCom'))
import agenda
import re
from datetime import datetime
from tqdm import tqdm
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列


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

    # def cus_lmt_cls_rec(self,fn='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH120肖婕.xlsm'):
    def cus_cls_rec(self,fn='E:\\temp\\minghu\\MH017李俊娴.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课']):

        df_tkn=pd.read_excel(fn,sheet_name='上课记录')
        df_buy=pd.read_excel(fn,sheet_name='购课表')
        df_ltm_prd=pd.read_excel(fn,sheet_name='限时课程记录')
        #排序
        # df_tkn.sort_values(by=['日期'],ascending=[True],inplace=True)
        # df_buy.sort_values(by=['收款日期'],ascending=[True],inplace=True)
        df_ltm_prd.sort_values(by=['限时课程起始日'],ascending=[True],inplace=True)

        cus_name=fn.split('\\')[-1].split('.')[0]

        try:
        #上课次数
            tkn_num=df_tkn['日期'].count()     
            tkn_nums={}
            for cls_type in cls_types:
                tkn_nums['上课次数-'+cls_type]=df_tkn[df_tkn['课程类型']==cls_type]['日期'].count()
            df_tkn_nums=pd.DataFrame(data=tkn_nums,index=[0])
        except:
            tkn_num=0

        try:
        #上课总天数
            interval=df_tkn['日期'].max()-df_tkn['日期'].min()
            interval=interval.days
        except:
            interval=0

        try:
        #上课频率
            tkn_frqc=interval/tkn_num
        except:
            tkn_frqc=0
        


        try:
        #到期日        
            latest_ltm=df_ltm_prd.iloc[-1]
            if pd.isna(latest_ltm['限时课程实际结束日']):
                end_date=latest_ltm['限时课程结束日']
            else:
                end_date=''        
        except:
            end_date=''

        try:
        #购课次数
            buy_nums={}
            for cls_type in cls_types:
                buy_nums['购课次数-'+cls_type]=df_buy[df_buy['购课类型']==cls_type]['购课编码'].nunique()
            df_buy_nums=pd.DataFrame(data=buy_nums,index=[0])
        #消费总金额
            buy_pays={}
            for cls_type in cls_types:
                buy_pays['消费金额-'+cls_type]=df_buy[df_buy['购课类型']==cls_type]['实收金额'].sum()
            buy_pays=pd.DataFrame(data=buy_pays,index=[0])

            buy_num=df_buy['购课编码'].nunique()
            total_pay=df_buy['实收金额'].sum()
        #平均每单消费金额
            avg_pay=total_pay/buy_num

            #续课资料
            if buy_num>0:
                ctn_buy_num=buy_num-1
            else:
                ctn_buy_num=0
        except Exception as err:
            print('购课表统计错误：',err)
            buy_num=0
            total_pay=0
            avg_pay=0
            ctn_buy_num=0

  
        try:
            df_out=pd.DataFrame(data={'会员编码及姓名':cus_name,'限时课程到期日':end_date,'总消费金额':total_pay,'平均每单消费金额':avg_pay,
                                '开始上课日期':df_tkn['日期'].min(),'最后一次上课日期':df_tkn['日期'].max(),'上课总天数':interval,'上课总次数':tkn_num,
                                '上课频率':tkn_frqc},index=[0])
            df_out=pd.concat([df_out,df_tkn_nums,df_buy_nums,buy_pays],axis=1)
        except Exception as err:
            df_out=pd.DataFrame()
            print(cus_name,'：',err)


        return df_out

    def all_cus_cls_rec(self,dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',cls_types=['常规私教课','限时私教课','常规团课','限时团课']):
        dfs=[]
        pbar=tqdm(os.listdir(dir))
        for fn in pbar:
            pbar.set_description('正在读取所有客户信息 ')
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                df=self.cus_cls_rec(fn=os.path.join(dir,fn),cls_types=cls_types)
                if not df.empty:
                    dfs.append(df)
        dfs_out=pd.concat(dfs)

        # print(dfs_out)
        return dfs_out

    def due_cus(self,df_data,month,userids=['AXiao'],input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx'):
        if input_data_type=='xlsx':
            df_data=pd.read_excel(xlsx,sheet_name='客户上课及购课信息表') 
        if not month:
            month=int(datetime.now().month)
        
        if not df_data.empty:
            df_due=df_data[df_data['限时课程到期日'].dt.month==month]
            
            if not df_due.empty:
                df_cusname=df_due[['会员编码及姓名','限时课程到期日']]
                df_by_date=df_cusname.groupby(['限时课程到期日'])['会员编码及姓名'].apply(list).reset_index()
                df_by_date['日程文本']=df_by_date['会员编码及姓名'].apply(lambda x: '今日有下列会员课程到期：\n'+'\n'.join(x).strip())
                df_by_date['写入日程日期']=pd.to_datetime(df_by_date["限时课程到期日"]) + pd.Timedelta("8 hours")
                df_by_date['写入日程日期及文本']=df_by_date['写入日程日期'].astype(str)+'|'+df_by_date['日程文本']
                # print(df_by_date)
                date_and_txts=df_by_date['写入日程日期及文本'].tolist()

                writer=agenda.WeCom()
                for d_and_t in date_and_txts:       
                    start_time,agenda_txt=d_and_t.split('|')
                    s_date=start_time.split(' ')[0]
                    print(f'\n正在写入 {s_date} 的记录……',end='')
                    end_time=datetime.strptime(start_time.split(' ')[0]+' 23:00:00','%Y-%m-%d %H:%M:%S')
                    writer.create_schedule(userids, 
                                    desp=agenda_txt, 
                                    start_time=start_time,
                                    end_time=end_time,                      
                                    access_token_fn='e:\\temp\\minghu\\access_token\\access_token.txt')

                return df_by_date
            else:
                print(f'{month} 月份没有到期的客户')
                return
        else:
            print('客户上课及购课信息为空')
            return


if __name__=='__main__':
    p=CusData()

    p.due_cus(df_data='',month='',userids=['AXiao','hal','WoShiXinMeiMei','WeiYueQi','likw'],input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx')
    #input_data_type参数：可以为xlsx或dataframe，如果为dataframe，则df_data参数需输入一个dataframe，如为xlsx，则xlsx参数需输入一个xlsx表格。

    # p.batch_fomral_cls_taken(dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx')
    # res=p.cus_cls_rec(fn='E:\\temp\\minghu\\MH017李俊娴.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # print(res)
    # res=p.all_cus_cls_rec(dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # res=p.all_cus_cls_rec(dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # res.to_excel('e:\\temp\\minghu\\客户上课及购课信息.xlsx',sheet_name='客户上课及购课信息表',index=False)

    # res=p.get_cus_buy()
    # res=p.batch_get_cus_buy()
    # p.exp_all_cus_buy(input_dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',output_dir='e:\\temp\\minghu')
    # p.formal_cls_taken()
    # res=p.all_trial_cls()
    # print(res)
    # res.to_excel('E:\\temp\\minghu\\所有体验课合并.xlsx',sheet_name='所有体验课数据',index=False)
  