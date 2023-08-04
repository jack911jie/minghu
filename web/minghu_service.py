import os
import sys
sys.path.extend([os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'data_analysis'),os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'modules')])
# print(sys.path)
import readconfig
import cus_data
import get_data
from datetime import datetime,date
import time
from dateutil.relativedelta import relativedelta
import re
import xlwings as xw
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
from flask import Flask, request, jsonify,render_template

class MinghuService(Flask):

    def log_time_and_function_name(func):
        def wrapper(*args, **kwargs):
            dt=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            func
            print(f"在 {dt} 执行 '{func.__name__}()' ")
            return 
        return wrapper



    def __init__(self,*args,**kwargs):
        super(MinghuService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.join(os.path.dirname(__file__),'config','minghu_service.config'))
        self.config_mh=readconfig.exp_json2(config_fn)

        #路由
        #渲染页面
        #首页
        self.add_url_rule('/',view_func=self.index)
        #获取客户信息页面
        self.add_url_rule('/cus_infos',view_func=self.cus_infos)
        #欢迎页面
        self.add_url_rule('/welcome',view_func=self.welcome)
        #上课记录页面
        self.add_url_rule('/cus_cls_input',view_func=self.cus_cls_input)
        #新生成客户页面
        self.add_url_rule('/new_cus', view_func=self.new_cus,methods=['GET','POST'])
        #购课页面
        self.add_url_rule('/input_buy', view_func=self.input_buy,methods=['GET','POST'])
        #成功页面
        self.add_url_rule('/success', view_func=self.success,methods=['GET','POST'])
        #围度测量页面
        self.add_url_rule('/input_body', view_func=self.input_body,methods=['GET','POST'])
        #限时课开课页面
        self.add_url_rule('/start_limit_class', view_func=self.start_limit_class,methods=['GET','POST'])
        # 体验课上课页面
        self.add_url_rule('/trial_class', view_func=self.trial_class,methods=['GET','POST'])
        
               

        #功能
        #从模板.xlsm获取基本信息，如教练姓名、课程种类等
        self.add_url_rule('/get_template_info', view_func=self.get_template_info,methods=['GET','POST'])
        #遍历会员资料文件夹获取所有 客户列表
        self.add_url_rule('/get_cus_list', view_func=self.get_cus_list,methods=['GET','POST'])
        #从,'03-教练管理','教练资料','教练信息.xlsx'中获取教练列表
        self.add_url_rule('/get_ins_list', view_func=self.get_ins_list,methods=['GET','POST'])
        #获取客户信息，包括既往购课、上课信息，剩余课时信息，限时课程信息，围度测量信息
        self.add_url_rule('/get_cus_info', view_func=self.get_cus_info,methods=['GET','POST'])
        #打开客户的xlsm文件
        self.add_url_rule('/open_cus_fn', view_func=self.open_cus_fn,methods=['GET','POST'])
        #遍历会员资料文件夹，生成新的客户ID号
        self.add_url_rule('/check_new', view_func=self.check_new,methods=['GET','POST'])
        #通过copy模板.xlsm生成新会员的文件 
        self.add_url_rule('/generate_new', view_func=self.generate_new,methods=['GET','POST'])
        #写入购课记录
        self.add_url_rule('/write_buy', view_func=self.write_buy,methods=['GET','POST'])
        #获取客户既往购课记录,并整理合并
        self.add_url_rule('/get_cus_buy', view_func=self.get_cus_buy,methods=['GET','POST'])
        #获取客户既往购课记录表，不整理
        self.add_url_rule('/get_cus_buy_list', view_func=self.get_cus_buy_list,methods=['GET','POST'])
        #获取客户未开课的购课编码
        self.add_url_rule('/deal_start_limit_page', view_func=self.deal_start_limit_page,methods=['GET','POST'])
        #通过'05-专业资料','训练项目.xlsx'获取训练项目的名称及分类
        self.add_url_rule('/get_train_list', view_func=self.get_train_list,methods=['GET','POST'])        
        #执行写入上课记录、写入训练记录
        self.add_url_rule('/deal_cls', view_func=self.deal_cls,methods=['GET','POST'])
        #获取既往体测记录
        self.add_url_rule('/get_body_history', view_func=self.get_body_history,methods=['GET','POST'])
        #写入体测记录
        self.add_url_rule('/write_body', view_func=self.write_body,methods=['GET','POST'])
        #写入体测记录
        self.add_url_rule('/deal_start_class_page', view_func=self.deal_start_class_page,methods=['GET','POST'])
        # 获取体验课记录
        self.add_url_rule('/get_trial_list', view_func=self.get_trial_list,methods=['GET','POST'])
        # 写入体验课上课记录
        self.add_url_rule('/write_trial_rec', view_func=self.write_trial_rec,methods=['GET','POST'])


    def write_trial_rec(self):
        try:
            data=request.json
            fn=os.path.join(self.config_mh['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
            df_old=pd.read_excel(fn,sheet_name='体验课上课记录表')
            df_old.dropna(subset=['体验客户姓名'],inplace=True)

            app=xw.App(visible=False)
            wb=app.books.open(fn)
            sht=wb.sheets['体验课上课记录表']
            row = df_old.shape[0]+2
            sht.range(f'B{row}').value=data['dateString']
            sht.range(f'C{row}').value=data['timeString']
            sht.range(f'D{row}').value=1
            sht.range(f'E{row}').value=data['cusNameInput']
            sht.range(f'F{row}').value=data['mobilePhone']
            sht.range(f'G{row}').value=data['ins']
            sht.range(f'H{row}').value='是'
            sht.range(f'I{row}').value=data['cusSource']
            sht.range(f'J{row}').value=data['comment']

            wb.save(fn)
            wb.close()
            app.quit()
        except Exception as e:
            print('写入体验课表错误：',e)
            return '写入体验课表成功'+e

        return '写入体验课表成功'
        


    def trial_class(self):
        return render_template('./trial_class.html')
         
    def date_to_string(self,obj,format='date'):
        try:
            if format=='date':
                return obj.strftime('%Y-%m-%d')
            elif format=='time':
                return obj.strftime('%H:%M')+':00'
        except:
            return None

        raise TypeError(f"Object of type {obj.__class__.__name__} is not JSON serializable")

    def get_trial_list(self,):
        
        df_trial=pd.read_excel(os.path.join(self.config_mh['work_dir'],'03-教练管理','体验课上课记录表.xlsx'),sheet_name='体验课上课记录表')
        df_trial.dropna(subset=['体验客户姓名'],inplace=True)
        df_trial.fillna('',inplace=True)
        df_trial['体验课日期']=df_trial['体验课日期'].apply(lambda x:self.date_to_string(x,'date'))
        df_trial['体验课时间']=df_trial['体验课时间'].apply(lambda x:self.date_to_string(x,'time'))
        df_trial['出单日期']=df_trial['出单日期'].apply(lambda x:self.date_to_string(x,'date'))

        dic_trial=df_trial.to_dict()
        num_keys=len(dic_trial[list(dic_trial.keys())[0]])
        result={}
        for i in range(num_keys):
            result[i] = {key: dic_trial[key][i] for key in dic_trial}
        # print(result)
        
        return jsonify({'all_trial':result})


    def deal_start_class_page(self):
        data=request.json
        
        try:
            self.add_rec_to_start_class_table(dic=data)
            self.delete_rec_in_aux_table(dic=data)
            return jsonify({'result':'写入限时课程表及辅助表成功'})
        except Exception as e:
            print('写入限时课程表或辅助表错误：',e)
            return jsonify({'result':'写入限时课程表及辅助表成功错误'+e})

    
    def add_rec_to_start_class_table(self,dic):
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',dic['cusName'].strip()+'.xlsm')

        df=pd.DataFrame(dic,index=[0])
        df=df[['buyCode','startDate','endDate']]
        df_old=pd.read_excel(fn,sheet_name='限时课程记录')

        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['限时课程记录']
        row = df_old.shape[0]+2
        rng='A'+str(row)+':C'+str(row)
        sht.range(rng).value=df.iloc[0].tolist()

        wb.save(fn)
        wb.close()
        app.quit()

    
    def add_rec_in_aux_table(self,buy_code):
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',buy_code[:-8].strip()+'.xlsm')
        df_old=pd.read_excel(fn,sheet_name='辅助表')
        df_old=df_old[['未开课的购课编码']]
        df_old.dropna(how='any',inplace=True)

        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['辅助表']
        row = df_old.shape[0]+2
        sht.range('I'+str(row)).value=buy_code

        wb.save(fn)
        wb.close()
        app.quit()
        print('写入辅助表成功')
        return '限时课程写入辅助表成功'

    def delete_rec_in_aux_table(self,dic):
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',dic['cusName'].strip()+'.xlsm')
        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['辅助表']
        column = sht.range('I2:I20')
        value_to_delete=dic['buyCode']
    
        # 在列中查找指定值并删除它
        for cell in column:
            if cell.value == value_to_delete:
                cell.delete('up')

        wb.save(fn)
        wb.close()
        app.quit()
    
        
    def index(self):
        return render_template('index.html')

    def start_limit_class(self):
        return render_template('./start_limit_class.html')

    
    def write_body(self):
        try:
            data=request.json
            # print('写入身体数据：',data)
            fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',data['cusName']+'.xlsm')
            df_old=pd.read_excel(os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',data['cusName']+'.xlsm'),sheet_name='身体数据')
            df_new=pd.DataFrame(data,index=[0])
            #第一列为姓名，去除
            df_write=df_new.iloc[:,1:]

            # print(df_write)

            app=xw.App(visible=False)
            wb=app.books.open(fn)
            sht=wb.sheets['身体数据']
            row = df_old.shape[0]+2
            rng='A'+str(row)+':R'+str(row)
            sht.range(rng).value=df_write.iloc[0].tolist()

            wb.save(fn)
            wb.close()
            app.quit()

            return '后端：写入身体数据表成功'
        except Exception as e:
            print('后端写入身体数据错误：',e)
            return '后端写入身体数据错误：'+e


    def bfr(self,sex,birthday,ht,wt,waist):
        bfr_test=get_data.cals()
        if birthday:
            try:
                if re.match(r'\d{4}',str(birthday)) and 1900<int(birthday)<2999:
                    birthday=datetime.strptime(str(birthday)+'0101','%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
                elif re.match(r'\d{6}',str(birthday)) and datetime.strptime(str(birthday)+'01','%Y%m%d'):
                    birthday=datetime.strptime(str(birthday)+'01','%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
                elif re.match(r'\d{8}',str(birthday)) and datetime.strptime(birthday,'%Y%m%d'):
                    birthday=datetime.strptime(str(birthday)+str('01'),'%Y%m%d')
                    age=relativedelta(datetime.now(),birthday).years
                    bfr=bfr_test.bfr(age=age,sex=sex,ht=ht,wt=wt,waist=waist,adj_bfr='no',adj_src='prg',formula=1)
            except Exception as e:
                print('bfr计算错误:',e)
        else:
            bfr=0
        return bfr


    #将dic的数据整理为按0，1，2的形式以输出给前端
    def dic_format(self,dic,order_name):
        formatted_data = {}
        # 遍历原始数据，并根据需要构建新的格式化数据

        for i in range(len(dic[order_name])):
            items={}
            for key,value in dic.items():
                items[key]=dic[key][i]                
            formatted_data[str(i)]=items
        return formatted_data

    
    def get_body_history(self):
        cus_name=request.data.decode('utf-8')
        print(f'\n{cus_name}')
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name+'.xlsm')
        df_body=pd.read_excel(fn,sheet_name='身体数据')
        df_basic=pd.read_excel(fn,sheet_name='基本情况')

        sex=df_basic['性别'].tolist()[0]
        birthday=df_basic['出生年月'].tolist()[0]

        # print(df_body)
        if df_body.empty:
            empty_body={'0': {'日期': '', '身高（cm）': '', '体重（Kg）': '', '体脂率': '', 
                        '胸围': '', '左臂围': '', '右臂围': '', '腰围': '', '臀围': '', '左腿围': '', 
                        '右腿围': '', '左小腿围': '', '右小腿围': '', '心肺': '', '平衡': '', 
                        '力量': '', '柔韧性': '', '核心': ''}}
            return jsonify(empty_body)
        else:
            df_body.dropna(subset=['日期'],inplace=True)
            df_body.fillna(0,inplace=True)
            
            dic_body=df_body.to_dict()

            formatted_data=self.dic_format(dic=dic_body,order_name='日期')

        # 计算bfr
        try:
            for key,item in formatted_data.items():
                item['体脂率']=self.bfr(sex,birthday,item['身高（cm）'],item['体重（Kg）'],item['腰围'])
        except Exception as e:
            print('bfr计算错误',e)
            item['体脂率']='-'

        print(formatted_data)
        # print(dic_body)
        return jsonify(formatted_data)
    
    def input_body(self):
        return render_template('./input_body.html')

    def deal_cls(self):
        cls_data=request.json

        cls_rec=cls_data['cls_tkn']
        # train_rec=cls_data['train_rec']
        # 写入上课记录
        cls_tkn_res=self.write_cls_tkn(cls_rec)

        # 写入训练情况
        train_rec_res=self.write_train_rec(cls_data)

        # print(cls_data)
        return cls_tkn_res+'\n'+train_rec_res

    def train_info(self,action_name,dic):
    # print('162----',action_name,dic)
        if action_name:
            return dic['by_action_name'][action_name]
        else:
            return ['','','']

    
    def write_cls_tkn(self,dic_tkn):
        # print(dic_tkn)
        try:
            fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',dic_tkn['cus_name']+'.xlsm')
            df_tkn=pd.DataFrame(dic_tkn,index=[0])
            df_tkn=df_tkn[['cls_tkn_date','cls_tkn_time','cls_long','cls_type','ins_name','basic_cls_comment']]
            df_old=pd.read_excel(fn,sheet_name='上课记录')


            app=xw.App(visible=False)
            wb=app.books.open(fn)
            sht=wb.sheets['上课记录']
            row = df_old.shape[0]+2
            rng='A'+str(row)+':F'+str(row)
            sht.range(rng).value=df_tkn.iloc[0].tolist()

            wb.save(fn)
            wb.close()
            app.quit()
            return '\n写入上课表成功'
        except Exception as e:
            return  f'写入上课表错误： {e}'
    
        
    def write_train_rec(self,dic):
        # dic=request.json
        trainlist=self.get_train_dic()
        # print(trainlist)
        basic_cls_info=dic['cls_tkn']
        print(dic['train_rec']['train_recs'])
        cus_name,ins_name=basic_cls_info['cus_name'],basic_cls_info['ins_name']
        # print(dic)
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name+'.xlsm')
        # print(dic['train_rec'])
    
        
        df_train_rec=pd.DataFrame(dic['train_rec']['train_recs'])
        # print(df_train_rec)
        #'nonOxyWt','nonOxyDis','nonOxyNum','nonOxyGroup','oxyTime','oxyGroup','calories'
        df_train_rec['nonOxyGroup']=df_train_rec['nonOxyGroup'].replace('','1')
        df_train_rec['oxyGroup']=df_train_rec['oxyGroup'].replace('','1')
        df_train_rec['oxyTime']=df_train_rec['oxyTime'].replace('','0')
        df_train_rec['nonOxyWt']=df_train_rec['nonOxyWt'].replace('','0')
        df_train_rec['nonOxyDis']=df_train_rec['nonOxyDis'].replace('','0')
        df_train_rec['nonOxyNum']=df_train_rec['nonOxyNum'].replace('','0')
        # df_train_rec=df_train_rec[['cls_tkn_date','cls_tkn_time','cls_long','cls_type','ins_name','basic_cls_comment']]
        # df_old=pd.read_excel(fn,sheet_name='训练情况')
        try:
            df_copy_rows_nonOxy= df_train_rec.loc[df_train_rec.index.repeat(df_train_rec['nonOxyGroup'].astype(int))]
        except Exception as e:
            print(f'\n写入训练记录表错误：{e}')
            return f'\n写入训练记录表错误：{e}'
        try:
            df_copy_rows= df_copy_rows_nonOxy.loc[df_train_rec.index.repeat(df_train_rec['oxyGroup'].astype(int))]
        except Exception as e:
            print(f'\n写入训练记录表错误：{e}')
            return f'\n写入训练记录表错误：{e}'
        
        df_copy_rows.reset_index(drop=True, inplace=True)

        df_copy_rows['date']=dic['train_rec']['trainDate']
        df_copy_rows['calories']=dic['train_rec']['calories']
        df_copy_rows['train_comment']=dic['train_rec']['trainComment']
        df_copy_rows['search_name']=df_copy_rows['nonOxyName']+df_copy_rows['oxyName']

        # print(df_copy_rows)

        try:
            df_copy_rows['big_type']=df_copy_rows['search_name'].apply(lambda x: self.train_info(x,trainlist)[0])
            df_copy_rows['muscle']=df_copy_rows['search_name'].apply(lambda x: self.train_info(x,trainlist)[1]+'肌群')
        except Exception as e:
            print('从动作名称获取动作大类及肌肉部位错误：',e)
            return '从动作名称获取动作大类及肌肉部位错误：:'+e
        df_copy_rows.loc[df_copy_rows['big_type'] == '有氧训练', 'muscle'] = ''
        # print(df_copy_rows)

        to_int_list=['nonOxyNum','nonOxyGroup','oxyGroup']
        to_float_list=['oxyTime','nonOxyWt','nonOxyDis','calories']
        for itm in to_int_list:
            try:
                df_copy_rows[itm]=df_copy_rows[itm].fillna(0)
                df_copy_rows[itm]=df_copy_rows[itm].astype(int)
            except Exception as e:
                return 'err:'+e
        for itm in to_float_list:
            try:
                df_copy_rows[itm]=df_copy_rows[itm].fillna(0)
                df_copy_rows[itm]=df_copy_rows[itm].astype(float)
            except Exception as e:
                return 'err:'+e
        try:
            df_copy_rows['oxyTime']=df_copy_rows['oxyTime'].apply(lambda x: int(x)*60)
        except Exception as e:
            return 'err:'+e

        df_copy_rows.replace(0,'',inplace=True)
        df_train_recs=df_copy_rows[['date','big_type','muscle','oxyName','oxyTime','nonOxyName','nonOxyWt','nonOxyDis','nonOxyNum','calories','train_comment']]
        df_train_recs_copy=df_train_recs.copy()
        df_train_recs_copy['ins_name']=ins_name
        df_write_train_rec=df_train_recs_copy[['date','big_type','muscle','oxyName','oxyTime','nonOxyName','nonOxyWt','nonOxyDis','nonOxyNum','calories','ins_name','train_comment']]
        # print(df_write_train_rec)

        df_old=pd.read_excel(fn,sheet_name='训练情况')
        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['训练情况']
        row_start = df_old.shape[0]+2
        row_end=row_start+df_write_train_rec.shape[0]-1
        # rng='A'+str(row_start)+':J'+str(row_end)
        start_cell='A'+str(row_start)
        # sht.range(rng).value=df_old.tolist()
        sht.range(start_cell).options(index=False, header=False).value=df_write_train_rec

        wb.save(fn)
        wb.close()
        app.quit()

        return f'\n写入训练记录表成功'

    def wecom_dir(self):
        # fn=os.path.join(os.path.dirname(__file__),'config','wecom_dir.config')        
        res=os.path.join(self.config_mh['work_dir'].strip(),'01-会员管理','会员资料')
        return res

    def get_cus_buy(self):
        cus_name=request.data.decode('utf-8')
        print(cus_name)
        
        # print(fn)
        try:
            fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
            df=pd.read_excel(fn,sheet_name='购课表')
            df['收款次数']=1
            # df_gp=df.groupby('购课编码')
            df_res = df.groupby('购课编码').agg({
            '应收金额': 'mean',
            '实收金额': 'sum',
            '购课类型': 'first',
            '收款日期': lambda x: '\n'.join(x.dt.strftime('%Y/%m/%d')),
            '收款次数': 'count'
        }).reset_index()
            df_res['未收金额']=df_res['应收金额']-df_res['实收金额']

            df_res=df_res[['购课编码','购课类型','应收金额','实收金额','未收金额','收款次数','收款日期']]
            # df_res.reset_index(drop=True,inplace=True)
            j_data={}
            for key ,value in df_res.to_dict(orient='index').items():
                j_data[key]=list(value.values())
            # result = {str(i): {key: value[i] for key, value in df_res.items()} for i in range(len(df_res['购课编码']))}
            return jsonify(j_data)
            # return jsonify(j_data)

        except Exception as err:
            return {'dat':'获取客户购课错误：','error':err}      
   
    def get_cus_buy_list(self,cus_name):
        # cus_name=request.data.decode('utf-8')
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
        df=pd.read_excel(fn,sheet_name='购课表')
        # dic_buy=df.to_dict()
        # dic_res=self.dic_format(dic=dic_buy,order_name='收款日期')
        # return jsonify(dic_res)
        return df

    def  get_limit_class_records(self,cus_name):
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
        df=pd.read_excel(fn,sheet_name='限时课程记录')
        return df

    def get_not_start_lmt_list(self,cus_name):
        # cus_name=request.data.decode('utf-8')
        fn=os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
        df=pd.read_excel(fn,sheet_name='辅助表')
        df_not_start=df['未开课的购课编码']
        df_not_start.dropna(inplace=True)
        df_not_start=df_not_start.to_frame()
        # dic_not_start=df_not_start.to_dict()
        # print(df_not_start,dic_not_start)
        # return jsonify(dic_not_start)
        return df_not_start

    def deal_start_limit_page(self):
        cus_name_input=request.data.decode('utf-8')
        print('deal_start_limit_page:',cus_name_input)
        try:
            # 获取购课表
            df_buy=self.get_cus_buy_list(cus_name=cus_name_input)
            # print(df_buy)
            if df_buy.empty:
                dic_buy=''
            else:
                df_buy.fillna('',inplace=True)
                dic_buy=df_buy.to_dict()
                dic_buy=self.dic_format(dic=dic_buy,order_name='收款日期')
            

            # 获取限时课程记录、目前生效的限时课程记录
            df_limit_cls_recs=self.get_limit_class_records(cus_name=cus_name_input)
            # print("df_limit_cls_recs",df_limit_cls_recs)
            if df_limit_cls_recs.empty:
                dic_limit_cls_recs=''
                dic_limit_maxdate_rec={'购课编码':'','限时课程起始日':'','限时课程结束日':'','限时课程实际结束日':''}
                dic_limit_maxdate_rec={'0':dic_limit_maxdate_rec}
            else:
                df_limit_cls_recs.fillna('',inplace=True)
                df_limit_maxdate_rec=df_limit_cls_recs[df_limit_cls_recs['限时课程结束日']==df_limit_cls_recs['限时课程结束日'].max()]
                df_limit_maxdate_rec.reset_index(inplace=True)
                # print('df_limit_maxdate_rec:',df_limit_maxdate_rec)

                dic_limit_cls_recs=df_limit_cls_recs.to_dict()
                dic_limit_cls_recs=self.dic_format(dic=dic_limit_cls_recs,order_name='购课编码')


                dic_limit_maxdate_rec=df_limit_maxdate_rec.to_dict()
                dic_limit_maxdate_rec=self.dic_format(dic=dic_limit_maxdate_rec,order_name='购课编码')
                
            print(dic_limit_maxdate_rec)

            # 获取并计算未开课的限时课程表，包括限时私教和团课
            df_buy_limit = df_buy[df_buy['购课类型'].isin(['限时私教课', '限时团课'])]
            df_not_start=self.get_not_start_lmt_list(cus_name=cus_name_input)
            df_not_start.fillna('',inplace=True)
            # print(type(df_not_start),df_not_start)
            df_not_start.rename(columns={'未开课的购课编码':'购课编码'},inplace=True)
            if df_not_start.empty:            
                df_merge=pd.DataFrame(data={'收款日期':'','购课编码':'','购课类型':'','购课节数':'','购课时长（天）':'',
                                            '应收金额':'','实收金额':'','收款人':'','收入类别':'','备注':''},index=[0])
                # df_merge=df_buy_limit                      
            else:
                df_merge=pd.merge(df_buy_limit,df_not_start,on='购课编码',how='inner')  
    
            # print('\n471 line',df_merge)
            df_merge.reset_index(inplace=True)

            if df_merge.empty:
                dic_not_start=''
            else:      
                
                df_merge.sort_values(by='收款日期',ascending=True,inplace=True)
                dic_not_start=df_merge.to_dict()
                # print('\n479 line dic_not_start',dic_not_start)
                dic_not_start=self.dic_format(dic=dic_not_start,order_name='收款日期')
            
            return jsonify({'not_start_list':dic_not_start,'buy_list':dic_buy,'limit_cls_recs':dic_limit_cls_recs,'maxdate_limit_class_rec':dic_limit_maxdate_rec})
        except Exception as e:
            return jsonify({'error':e,'not_start_list':dic_not_start,'buy_list':dic_buy,'limit_cls_recs':dic_limit_cls_recs,'maxdate_limit_class_rec':dic_limit_maxdate_rec})


    def read_template(self):
        df=pd.read_excel(os.path.join(self.config_mh['work_dir'],'01-会员管理','模板.xlsm'),sheet_name='辅助表')
        df_cls_types=df[['购课类型']].copy().dropna()
        cls_types=df_cls_types['购课类型'].tolist()
        
        df_cashier=df[['收款人']].copy().dropna()
        cashier=df_cashier['收款人'].tolist()

        df_income_types=df[['收入类别']].copy().dropna()
        income_types=df_income_types['收入类别'].tolist()

        data={'cls_types':cls_types,'cashiers':cashier,'income_types':income_types}
        return data

    def write_buy_date_to_trial_table(self,formal_cus_name,first_buy_date):
        print('写入体验表：',formal_cus_name,first_buy_date)
        try:
            fn=os.path.join(self.config_mh['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
            app=xw.App(visible=False)
            wb=app.books.open(fn)
            sht=wb.sheets['体验课上课记录表']
            m_column_data=sht.range('M:M').value
            # 从M列找到 出单后正式会员编码及姓名
            for row_index, value in enumerate(m_column_data, start=1):
                if value == formal_cus_name:
                    print(value,first_buy_date)
                    # 根据条件设置 M 列的值为 'AA'
                    if sht.range(f'L{row_index}').value is None:
                        sht.range(f'L{row_index}').value = first_buy_date
                        sht.range(f'K{row_index}').value = '是'

            wb.save(fn)
            wb.close()
            app.quit()
        except Exception as e:
            print('写入体验课表错误：',e)
            return '写入体验课表错误：'+e
    
    def write_buy(self,):
        wk_dir=self.config_mh['work_dir']
        dat=request.json
        for key,value in dat.items():
            try:
                dat[key]=int(value)
            except:
                pass
        fn=os.path.join(wk_dir,'01-会员管理','会员资料',dat['客户编码及姓名'].strip()+'.xlsm')

        df=pd.DataFrame(dat,index=[0])
        df=df[['收款日期','客户购课编号','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']]
        df_old=pd.read_excel(fn,sheet_name='购课表')

        app=xw.App(visible=False)
        wb=app.books.open(fn)
        sht=wb.sheets['购课表']
        row = df_old.shape[0]+2
        rng='A'+str(row)+':J'+str(row)
        sht.range(rng).value=df.iloc[0].tolist()

        wb.save(fn)
        wb.close()
        app.quit()
        

        if dat['购课类型'].strip() in ['限时私教课','限时团课']:
            aux_res=self.add_rec_in_aux_table(dat['客户购课编号'])
        else:
            aux_res='非限时课程'


        # 写入体检课表的出单日期
        print(df['客户购课编号'].tolist()[0][:-8],df['收款日期'].tolist()[0])
        self.write_buy_date_to_trial_table(formal_cus_name=df['客户购课编号'].tolist()[0][:-8],first_buy_date=df['收款日期'].tolist()[0])

        return f'写入成功, 行号：{row}, {aux_res}'


    def cus_list(self):
        dic_li=[]
        for fn in os.listdir(os.path.join(self.config_mh['work_dir'],'01-会员管理','会员资料')):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                dic_li.append(fn.split('.')[0])
        return dic_li


    def ins_list(self):
        fn=os.path.join(self.config_mh['work_dir'],'03-教练管理','教练资料','教练信息.xlsx')
        df=pd.read_excel(fn,sheet_name='教练信息')
        ins_li=df['姓名'].tolist()

        return ins_li

    def  get_train_dic(self):
        fn=os.path.join(self.config_mh['work_dir'],'05-专业资料','训练项目.xlsx')
        df=pd.read_excel(fn,sheet_name='训练项目')
        df.fillna('',inplace=True)
        
        train_data={}
    
        train_data_by_action_name={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if action_name not in train_data_by_action_name:
                train_data_by_action_name[action_name] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_action_name[action_name].append(form)
            train_data_by_action_name[action_name].append(muscle)
            train_data_by_action_name[action_name].append(category)

        train_data_by_muscle={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if muscle not in train_data_by_muscle:
                train_data_by_muscle[muscle] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_muscle[muscle].append([form,category,action_name])   
            # train_data_by_muscle[muscle].append(form)
            # train_data_by_muscle[muscle].append(category)
            # train_data_by_muscle[muscle].append(action_name) 

        train_data_by_category={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if category not in train_data_by_category:
                train_data_by_category[category] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_category[category].append([form,muscle,action_name])    
            # train_data_by_category[category].append(form)
            # train_data_by_category[category].append(muscle)
            # train_data_by_category[category].append(action_name)

        train_data['by_action_name']=train_data_by_action_name
        train_data['by_muscle']=train_data_by_muscle
        train_data['by_category']=train_data_by_category
        return train_data


    def  get_train_list(self):
        fn=os.path.join(self.config_mh['work_dir'],'05-专业资料','训练项目.xlsx')
        df=pd.read_excel(fn,sheet_name='训练项目')
        df.fillna('',inplace=True)
        
        train_data={}
    
        train_data_by_action_name={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if action_name not in train_data_by_action_name:
                train_data_by_action_name[action_name] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_action_name[action_name].append(form)
            train_data_by_action_name[action_name].append(muscle)
            train_data_by_action_name[action_name].append(category)

        train_data_by_muscle={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if muscle not in train_data_by_muscle:
                train_data_by_muscle[muscle] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_muscle[muscle].append([form,category,action_name])   
            # train_data_by_muscle[muscle].append(form)
            # train_data_by_muscle[muscle].append(category)
            # train_data_by_muscle[muscle].append(action_name) 

        train_data_by_category={}
        for index,row in df.iterrows():
            action_name=row['动作名称']
            form=row['形式']
            muscle = row['肌肉部位']
            category = row['动作大类']
            
            # 判断动作名称是否已经在字典中，若不在则添加
            if category not in train_data_by_category:
                train_data_by_category[category] = []
            
            # 将相关信息添加到对应动作名称的列表中
            train_data_by_category[category].append([form,muscle,action_name])    
            # train_data_by_category[category].append(form)
            # train_data_by_category[category].append(muscle)
            # train_data_by_category[category].append(action_name)

        train_data['by_action_name']=train_data_by_action_name
        train_data['by_muscle']=train_data_by_muscle
        train_data['by_category']=train_data_by_category
        
        # return train_data

        return jsonify(train_data)

##--------------------------------------------

    def get_template_info(self):
        fromWeb=request.data
        infos=self.read_template()
        return jsonify(infos)

    #遍历会员资料生成名字
    def get_cus_list(self):
        dic_li=self.cus_list()
        return jsonify(dic_li)

    #获取教练信息
    def get_ins_list(self):
        ins_li=self.ins_list()
        return jsonify(ins_li)


    # 定义前端页面路由
   
    def cus_infos(self):
        return render_template('cus_infos.html')

    def input_buy(self):
        return render_template('input_buy.html')

    def new_cus(self):
        return render_template('new_cus.html')

    def cus_cls_input(self):
        return render_template('cus_cls_input.html')

    def success(self):
        return render_template('success.html')


    def get_cus_info(self):
        cus_name = request.json.get('selected_name')
        work_dir=self.wecom_dir()
        fn=os.path.join(work_dir,cus_name+'.xlsm')
        p=cus_data.CusData()
        res=p.cus_cls_rec_toweb(fn=fn,cls_types=self.config_mh['all_cls_types'],not_lmt_types=self.config_mh['not_lmt_cls_types'])
        res.fillna(0)
        data=res.iloc[0].to_dict()
        return jsonify(data)


    def open_cus_fn(self):
        cus_name=request.data.decode('utf-8')
        cus_li=self.cus_list()
        if cus_name and cus_name in cus_li:
            work_dir=self.wecom_dir()
            fn=os.path.join(work_dir,cus_name+'.xlsm')
            # os.startfile(fn)
            return f'正在打开 {cus_name} 的会员档案'
        else:
            return '会员编码及编码为空/无此会员档案'


    def check_new(self):
        dat=request.data
        cus_li=self.cus_list()
        cus_num=[int(x[2:5]) for x in cus_li]
        max_num=max(cus_num)
        new_num=max_num+1
        txt_num=str(new_num).zfill(3)
        # new_name='MH'+new_num.zfill(3)+cus_name+'.xlsm'
        # new_name=os.path.join(wecom_dir,new_name)
        print('新客户自动编号：',txt_num)
        return txt_num

    
    def generate_new(self):
        try:
            fn_in=request.data
            fn='MH'+fn_in.decode('utf-8')
            print(fn)
            fn,trial_cus_name,sex,birthMonth,cusSource,dvc=fn.split('|')
            work_dir=self.wecom_dir()
            tplt_dir=os.path.dirname(work_dir)
            new_fn=os.path.join(work_dir,fn+'.xlsm')

            app=xw.App(visible=False)
            wb=app.books.open(os.path.join(tplt_dir,'模板.xlsm'))
            sht=wb.sheets['基本情况']
            sht['A2'].value=fn[0:5]
            sht['B2'].value=fn[5:]
            sht['D2'].value=sex
            sht['E2'].value=birthMonth
            sht['F2'].value=cusSource
            if len(fn[5:])>1:
                sht['C2'].value=fn[5:][1:]
            else:
                sht['C2'].value=fn[5:]

            wb.save(new_fn)
            wb.close()
            app.quit()

            if trial_cus_name:
                self.write_deal_cus_name_to_trial_table(formal_cus_name=fn,trial_cus_name=trial_cus_name)

            # os.startfile(work_dir)
            # if dvc=='pc':
            #     os.startfile(new_fn)

            return new_fn
        except Exception as e:
            return e
    
    def write_deal_cus_name_to_trial_table(self,formal_cus_name,trial_cus_name):
        try:
            fn=os.path.join(self.config_mh['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
            app=xw.App(visible=False)
            wb=app.books.open(fn)
            sht=wb.sheets['体验课上课记录表']
            e_column_data=sht.range('e1').expand('down').value

            for row_index, value in enumerate(e_column_data, start=1):
                if value == trial_cus_name:
                    # 根据条件设置 M 列的值为 'AA'
                    sht.range(f'M{row_index}').value = formal_cus_name

            wb.save(fn)
            wb.close()
            app.quit()
        except Exception as e:
            print('写入体验课表错误：',e)
            return '写入体验课表错误：'+e



    def welcome(self):
        return '关于我们页面'

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__ == '__main__':
    app = MinghuService(__name__)
    if len(sys.argv)>1:
        print(f'服务器为：{sys.argv[1]}:5000')
        app.run(debug=True,host=sys.argv[1],port=5000)
    else:
        app.run(debug=True)
    # app.run(debug=True,host='192.168.158.71',port=5000)
    # app.run(debug=True,host='192.168.10.2',port=5000)
    # app.run(debug=True,host='192.168.1.41',port=5000)
    # app.run(debug=True,host='192.168.1.149',port=5000)
    # res=wecom_dir()
    # print(res)
