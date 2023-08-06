import os
import sys
sys.path.extend([os.path.join(os.path.dirname(os.path.dirname(__file__)),'WeCom'),os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules')])
import readconfig
import agenda
import get_data
from dateutil.relativedelta import relativedelta
import re
from datetime import datetime
from tqdm import tqdm
from flask import jsonify
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列


class CusData:
    def __init__(self):
        pass

    def get_cus_buy(self,fn='E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料\\MH207杨薇.xlsm',year_month=''):
        df_cus_buy=pd.read_excel(fn,sheet_name='购课表')
        if year_month:
            try:
                df_cus_buy=df_cus_buy[(df_cus_buy['收款日期'].dt.year==int(str(year_month)[:4])) & (df_cus_buy['收款日期'].dt.month==int(str(year_month)[4:])) ]
            except Exception as e:
                print(fn.split('\\')[-1],'  错误：',e)

        return df_cus_buy
        

    def batch_get_cus_buy(self,dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',year_month=''):
        dfs=[]    
        pbar=tqdm(os.listdir(dir))    
        for fn in pbar:
            pbar.set_description("正在读取客户购课信息")
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                realfn=os.path.join(dir,fn) 
                df=self.get_cus_buy(realfn,year_month)
                dfs.append(df)
        
        all_df_buy=pd.concat(dfs)
        
        return all_df_buy

    def exp_all_cus_buy(self,input_dir,output_dir,year_month,out_put_format='xlsx'):
        # print('\n正在抽取购课数据……',end='')
        all_df_buy=self.batch_get_cus_buy(dir=input_dir,year_month=year_month)
        # print('完成')
        if all_df_buy.shape[0]>0:
            print('\n正在提取及合并数据……',end='')
            fn=os.path.join(output_dir,datetime.now().strftime('%Y%m%d%H%M')+'-购课数据.xlsx')
            all_df_buy.dropna(how='any',subset=['购课编码'],inplace=True)
            if out_put_format=='xlsx':
                all_df_buy.to_excel(fn,sheet_name='购课表',index=False)

        print('完成')
        return all_df_buy

    def merge_old_this_month_cus_buy(self,year_month,input_dir,output_dir,old_xlsx='E:\\temp\\minghu\\会员购课数据.xlsx'):
        df_old=pd.read_excel(old_xlsx,sheet_name='购课表')
        df_this_month=self.exp_all_cus_buy(input_dir=input_dir,output_dir=output_dir,year_month=year_month,out_put_format='xlsx')
        # df_this_month=df_this_month[['收款日期','编码','购课类型','应收金额','实收金额','收款人','收入类别','备注']]
        # df_this_month.rename(columns={'购课编码':'编码','收入类别':'购课类别'},inplace=True)
        
        df_out=pd.concat([df_old,df_this_month])
        df_out.to_excel(os.path.join(output_dir,'会员购课数据-'+str(year_month)+'月.xlsx'),index=False)
       
        print('完成')
        return df_out

    def merge_old_this_month_trial_class(self,month,output_dir,new_trial_table='E:\\temp\\minghu\\体验课上课记录.xlsx',old_xlsx='E:\\temp\\minghu\\所有体验课合并.xlsx'):
        df_old=pd.read_excel(old_xlsx,sheet_name='体验课上课记录表')
        df_new=pd.read_excel(new_trial_table,sheet_name='体验课上课记录表')    
        df_this_month=df_new[df_new['体验课日期'].dt.month==int(month)]
        df_out=pd.concat([df_old,df_this_month])
        df_out.to_excel(os.path.join(output_dir,'所有体验课合并-'+str(month)+'月.xlsx'),index=False)
       
        print('完成')
        return df_out

    def this_month_formal_cls_taken(self,year_month,fn='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH120肖婕.xlsm'):
        df_formal_tk=pd.read_excel(fn,sheet_name='上课记录')
        df_formal_tk['会员姓名']=fn.split('\\')[-1].split('.')[0]
     

        df_formal_tk=df_formal_tk[['日期', '时间', '时长（小时）', '课程类型', '会员姓名', '教练', '备注']]
        try:
            df_formal_tk_this_month=df_formal_tk[(df_formal_tk['日期'].dt.year==int(str(year_month)[:4])) & (df_formal_tk['日期'].dt.month==int(str(year_month)[4:])) ]
        except Exception as e:
            df_formal_tk_this_month=pd.DataFrame()
            print(fn.split('\\')[-1].split('.')[0],e)
        
        return df_formal_tk_this_month

    def batch_fomral_cls_taken(self,year_month,dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx',out_format='xlsx'):

        dfs_tk=[]
        pbar=tqdm(os.listdir(dir))
        for fn in pbar:
            pbar.set_description("正在读取教练上课记录")
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                df_tk=self.this_month_formal_cls_taken(year_month,os.path.join(dir,fn))
                dfs_tk.append(df_tk)
        
        df_taken=pd.concat(dfs_tk)
        df_taken.dropna(how='any',subset=['日期'],inplace=True)
        df_taken['时长（小时）']=1
        if out_format=='xlsx':
            df_taken.to_excel(out_fn,index=False,sheet_name='教练上课记录表')
        
        print('完成')
        return df_taken

    def merge_this_month_taken_to_old(self,year_month,old_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx',input_dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_dir='E:\\temp\\minghu'):
        
        df_new=self.batch_fomral_cls_taken(year_month=year_month,dir=input_dir,out_fn=os.path.join(out_dir,'当月教练上课记录-'+str(year_month)+'.xlsx'),out_format='dataframe')
        df_old=pd.read_excel(old_fn,sheet_name='教练上课记录表')
        df_merge=pd.concat([df_old,df_new])
        out_fn=os.path.join(out_dir,'教练上课记录合并'+str(year_month)+'.xlsx')
        df_merge.to_excel(out_fn,sheet_name='教练上课记录表',index=False)

        print('完成')
        return df_merge

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
    def cus_cls_rec(self,fn='E:\\temp\\minghu\\MH017李俊娴.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课'],not_lmt_types=['常规私教课','常规团课']):
        df_basic=pd.read_excel(fn,sheet_name='基本情况')
        df_tkn=pd.read_excel(fn,sheet_name='上课记录')
        df_buy=pd.read_excel(fn,sheet_name='购课表')
        df_ltm_prd=pd.read_excel(fn,sheet_name='限时课程记录')
        df_body_msr=pd.read_excel(fn,sheet_name='身体数据')
        #排序
        # df_tkn.sort_values(by=['日期'],ascending=[True],inplace=True)
        # df_buy.sort_values(by=['收款日期'],ascending=[True],inplace=True)
        df_ltm_prd.sort_values(by=['限时课程起始日'],ascending=[True],inplace=True)
        # print(df_ltm_prd.iloc[-1])

        cus_name=fn.split('\\')[-1].split('.')[0]

        try:
        #上课次数
            tkn_num=df_tkn['日期'].count()     
            tkn_nums={}
            for cls_type in cls_types:
                tkn_nums['上课次数-'+cls_type]=df_tkn[df_tkn['课程类型']==cls_type]['日期'].count()
            df_tkn_nums=pd.DataFrame(data=tkn_nums,index=[0])
        except Exception as e:
            print('上课次数错误',e)
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
        except Exception as e:
            print('上课频率错误',e)
            tkn_frqc=0
        


        try:
        #到期日        
            latest_ltm=df_ltm_prd.iloc[-1]
            if pd.isna(latest_ltm['限时课程实际结束日']):
                end_date=latest_ltm['限时课程结束日']
            else:
                end_date=latest_ltm['限时课程实际结束日']     
        except Exception as e:
            print('限时课程到期日错误在 cus_cls_rec：',e)
            end_date=''

        try:
        #购课次数
            buy_nums={}
            for cls_type in cls_types:
                buy_nums['购课次数-'+cls_type]=df_buy[df_buy['购课类型']==cls_type]['购课编码'].nunique()
            df_buy_nums=pd.DataFrame(data=buy_nums,index=[0])
        
        #非限时课程的购课节数
            #读取校正节数
            df_adj=pd.read_excel(fn,sheet_name='修正参数')
            if df_adj.empty:
                adj_tkn=0
            else:
                adj_tkn=df_adj['已上课时数'].tolist()[0]

            for cls_type in cls_types:
                if cls_type in not_lmt_types:
                    df_buy_drop_dup=df_buy.drop_duplicates(subset=['购课编码'],keep='first')
                    buy_nums['购课节数-'+cls_type]=df_buy_drop_dup[df_buy_drop_dup['购课类型']==cls_type]['购课节数'].sum()
                    buy_nums['剩余节数-'+cls_type]=buy_nums['购课节数-'+cls_type]-df_tkn_nums['上课次数-'+cls_type]-adj_tkn
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
        
        #最后一次购课时间
            latest_buy_date=df_buy['收款日期'].max()

            #续课次数
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

        #围度测量数据
        try:
            body_msr=Vividict()
            body_msr['lst_msr_date']=df_body_msr['日期'].max().strftime('%Y-%m-%d')
            body_msr['msr_num']=df_body_msr['日期'].count()
            #将体测的日期拼接成文本
            body_msr['msr_dates']='\n'.join([x.strftime('%Y-%m-%d') for x in df_body_msr['日期'].tolist()])


            sex=df_basic['性别'].tolist()[0]
            latest_body_data=df_body_msr[df_body_msr['日期']==df_body_msr['日期'].max()]
            birthday=df_basic['出生年月'].tolist()[0]
            lst_ht=latest_body_data['身高（cm）'].tolist()[0]
            lst_wt=latest_body_data['体重（Kg）'].tolist()[0]
            lst_waist=latest_body_data['腰围'].tolist()[0]
            latest_body_data['腰围'].tolist()[0]
        except Exception as e:
            birthday=''
            age=''
            lst_ht=''
            lst_wt='' 
            bfr=0
            latest_body_data=pd.DataFrame()
            print('围度测量数据错误:',e)

        
        bfr_test=get_data.cals()
        if birthday:
            try:
                if re.match(r'\d{4}',str(birthday)) and 1900<int(birthday)<2999:
                    birthday=datetime.strptime(str(birthday)+'0101','%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=lst_ht,wt=lst_wt,waist=lst_waist,adj_bfr='no',adj_src='prg',formula=1)
                elif re.match(r'\d{6}',str(birthday)) and datetime.strptime(str(birthday)+'01','%Y%m%d'):
                    birthday=datetime.strptime(str(birthday)+'01','%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=lst_ht,wt=lst_wt,waist=lst_waist,adj_bfr='no',adj_src='prg',formula=1)
                elif re.match(r'\d{8}',str(birthday)) and datetime.strptime(birthday,'%Y%m%d'):
                    birthday=datetime.strptime(str(birthday)+str('01'),'%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=lst_ht,wt=lst_wt,waist=lst_waist,adj_bfr='no',adj_src='prg',formula=1)
            except Exception as e:
                bfr=0
                age=''
                print('bfr计算错误:',e)
        else:
            bfr=0
        # 体脂率计算def bfr(self,age,sex,ht,wt,waist,adj_bfr='yes',adj_src='prg',gui='',formula=1):
        # body_msr['lst_msr']replace
        try:
            if not latest_body_data.empty:
                body_msr['bfr']=bfr
                body_msr['age']=age
                body_msr['ht']=lst_ht
                body_msr['wt']=lst_wt
                body_msr['waist']=lst_waist
                body_msr['chest']=latest_body_data['胸围'].tolist()[0]
                body_msr['l_arm']=latest_body_data['左臂围'].tolist()[0]
                body_msr['r_arm']=latest_body_data['右臂围'].tolist()[0]
                body_msr['hip']=latest_body_data['臀围'].tolist()[0]
                body_msr['l_leg']=latest_body_data['左腿围'].tolist()[0]
                body_msr['r_leg']=latest_body_data['右腿围'].tolist()[0]
                body_msr['l_calf']=latest_body_data['左小腿围'].tolist()[0]
                body_msr['r_calf']=latest_body_data['右小腿围'].tolist()[0]
                body_msr['heart']=latest_body_data['心肺'].tolist()[0]
                body_msr['balance']=latest_body_data['平衡'].tolist()[0]
                body_msr['power']=latest_body_data['力量'].tolist()[0]
                body_msr['flex']=latest_body_data['柔韧性'].tolist()[0]
                body_msr['core']=latest_body_data['核心'].tolist()[0]
            else:
                body_msr['bfr']=''
                body_msr['age']=''
                body_msr['ht']=''
                body_msr['wt']=''
                body_msr['waist']=''
                body_msr['chest']=''
                body_msr['l_arm']=''
                body_msr['r_arm']=''
                body_msr['hip']=''
                body_msr['l_leg']=''
                body_msr['r_leg']=''
                body_msr['l_calf']=''
                body_msr['r_calf']=''
                body_msr['heart']=''
                body_msr['balance']=''
                body_msr['power']=''
                body_msr['flex']=''
                body_msr['core']=''
        except Exception as e:
            print('写入身体数据错误')

        df_msr=pd.DataFrame(data=body_msr,index=[0])
  
        try:
            df_out=pd.DataFrame(data={'会员编码及姓名':cus_name,'限时课程到期日':end_date,'总消费金额':total_pay,'平均每单消费金额':avg_pay,'最后一次购课日期':latest_buy_date,
                                '开始上课日期':df_tkn['日期'].min(),'最后一次上课日期':df_tkn['日期'].max(),'上课总天数':interval,'上课总次数':tkn_num,
                                '上课频率':tkn_frqc},index=[0])
            df_out=pd.concat([df_out,df_tkn_nums,df_buy_nums,buy_pays,df_msr],axis=1)
        except Exception as err:
            df_out=pd.DataFrame()
            print('生成既往课程及围度测量结果错误',cus_name,'：',err)
 


        return df_out

    def cal_bfr(self,birthday,sex,ht,wt,waist):
        bfr_test=get_data.cals()
        try:
            if re.match(r'\d{4}',birthday) and 1900<int(birthday)<2999:
                birthday=datetime.strptime(str(birthday)+'0101','%Y%m%d')
                age=relativedelta(datetime.now(),birthday).years
                bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
            elif re.match(r'\d{6}',birthday) and datetime.strptime(str(birthday)+'01','%Y%m%d'):
                birthday=datetime.strptime(str(birthday)+'01','%Y%m%d')
                age=relativedelta(datetime.now(),birthday).years
                bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
            elif re.match(r'\d{8}',birthday) and datetime.strptime(birthday,'%Y%m%d'):
                birthday=datetime.strptime(str(birthday)+str('01'),'%Y%m%d')
                age=relativedelta(datetime.now(),birthday).years
                bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
            return bfr
        except Exception as e:
            print('bfr计算错误:',e)
            return 'err cal bfr:'+e


    def cus_cls_rec_toweb(self,fn='E:\\temp\\minghu\\MH017李俊娴.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课'],not_lmt_types=['常规私教课','常规团课']):
        df_web=self.cus_cls_rec(fn=fn,cls_types=cls_types,not_lmt_types=not_lmt_types)
        try:
            if df_web['限时课程到期日'].tolist()[0]>=datetime.now():
                df_web['限时课程是否有效']='是'
            else:
                df_web['限时课程是否有效']='否'
        except Exception as e:
            print('计算限时课程错误在cus_cls_rec_toweb中：',e)
            df_web['限时课程是否有效']='否'
            df_web['限时课程到期日']='-'

        df_web=df_web.fillna('-')
        return df_web

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

    def due_cus(self,df_data,month,input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx'):
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
                return df_by_date
            else:
                print(f'{month} 月份没有到期的客户')
                return
        else:
            print('客户上课及购课信息为空')
            return

    def send_agenda_due_cus(self,df_data,month,userids=['AXiao'],input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx'):
        df_by_date=self.due_cus(df_data=df_data,month=month,input_data_type=input_data_type,xlsx=xlsx)
                # print(df_by_date)
        if not df_by_date.empty:
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
        else:
            print('客户上课及购课信息为空')
            return                      

    def zombie_cus(self,df_data,today,input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx'):
        #读取设置
        zombie_cfg=readconfig.exp_json2(os.path.join(os.path.dirname(__file__),'config','zombie.config'))
        #排除一些客户
        exp_list=readconfig.txt_to_list(os.path.join(os.path.dirname(__file__),'config','zombie_except_list.config'))
        zombie_lstbuy_days=zombie_cfg['zombie_lstbuy_days']
        zombie_lsttkn_days=zombie_cfg['zombie_lsttkn_days']
        if input_data_type=='xlsx':
            df_data=pd.read_excel(xlsx,sheet_name='客户上课及购课信息表') 
        if today:
            today=datetime.strptime(today,'%Y-%m-%d')
        else:
            today=datetime.now()
        
        
        if not df_data.empty:
            #判断标准：最后一次购课日期与今天相比大于180日，& 最后一次上课日期与今天相比大于 180日
            df_zombie=df_data[(df_data['最后一次购课日期']<=today-pd.Timedelta(days=zombie_lstbuy_days)) & (df_data['最后一次上课日期']<=today-pd.Timedelta(days=zombie_lsttkn_days))]
            
            #剔除排除名单里面的客户
            if exp_list:
                df_zombie=df_zombie[~df_zombie['会员编码及姓名'].isin(exp_list)]

            df_zombie.reset_index(inplace=True,drop=True)
            return df_zombie
        else:
            print('输入的表格或dataframe为空')
            return
class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    p=CusData()

    #僵尸客户
    # res=p.zombie_cus(df_data='',today='',input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx')
    # df_data：输入的dataframe
    # today：与“今天”比较，可输入，格式为“2023-1-1”，如为空，则按今天的日期。
    # input_data_type: xlsx或dataframe， 如为xlsx则后面的xlsx参数必须有，如为dataframe，则前面的df_data必须有。
    
    # res.to_excel('e:\\temp\\minghu\\zombie.xlsx',sheet_name='僵尸客户名单',index=False)
    # print(res)
    # p.send_agenda_due_cus(df_data='',month='',userids=['AXiao','hal','WoShiXinMeiMei','WeiYueQi','likw'],input_data_type='xlsx',xlsx='e:\\temp\\minghu\\客户上课及购课信息.xlsx')
    #input_data_type参数：可以为xlsx或dataframe，如果为dataframe，则df_data参数需输入一个dataframe，如为xlsx，则xlsx参数需输入一个xlsx表格。

    # p.batch_fomral_cls_taken(dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',out_fn='E:\\temp\\minghu\\教练上课记录合并.xlsx')
    # res=p.cus_cls_rec(fn='E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料\\MH041陈智翀.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # res=p.cus_cls_rec_toweb(fn='E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料\\MH041陈智翀.xlsm',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # res=res.fillna('-')
    # print(res)
    # res=p.all_cus_cls_rec(dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    
    # 客户上课及购课信息
    # res=p.all_cus_cls_rec(dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',cls_types=['常规私教课','限时私教课','常规团课','限时团课'])
    # res.to_excel('e:\\temp\\minghu\\客户上课及购课信息.xlsx',sheet_name='客户上课及购课信息表',index=False)

    # res=p.get_cus_buy(month='8')
    res=p.merge_old_this_month_cus_buy(month=7,input_dir='E:\\temp\\minghu\\铭湖健身工作室\\01-会员管理\\会员资料',output_dir='E:\\temp\\minghu',old_xlsx='E:\\temp\\minghu\\客户业务流水数据.xlsx')
    # res=p.batch_get_cus_buy()
    # p.exp_all_cus_buy(input_dir='E:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料',output_dir='e:\\temp\\minghu')
    # p.formal_cls_taken()
    # res=p.all_trial_cls()
    print(res)
    # res.to_excel('E:\\temp\\minghu\\所有体验课合并.xlsx',sheet_name='所有体验课数据',index=False)

  