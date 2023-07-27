import os
import sys
sys.path.extend([os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'data_analysis'),os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))),'modules')])
# print(sys.path)
import readconfig
import cus_data
import get_data
import re
import xlwings as xw
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
from flask import Flask, request, jsonify,render_template

class MinghuService(Flask):
    def __init__(self,*args,**kwargs):
        super(MinghuService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.join(os.path.dirname(__file__),'config','minghu_service.config'))
        self.config_mh=readconfig.exp_json2(config_fn)

        #路由
        self.add_url_rule('/',view_func=self.index)
        self.add_url_rule('/cus_infos',view_func=self.cus_infos)
        self.add_url_rule('/welcome',view_func=self.welcome)
        self.add_url_rule('/cus_cls_input',view_func=self.cus_cls_input)
        self.add_url_rule('/get_cus_list', view_func=self.get_cus_list,methods=['GET','POST'])
        self.add_url_rule('/get_ins_list', view_func=self.get_ins_list,methods=['GET','POST'])
        self.add_url_rule('/get_cus_info', view_func=self.get_cus_info,methods=['GET','POST'])
        self.add_url_rule('/open_cus_fn', view_func=self.open_cus_fn,methods=['GET','POST'])
        self.add_url_rule('/check_new', view_func=self.check_new,methods=['GET','POST'])
        self.add_url_rule('/generate_new', view_func=self.generate_new,methods=['GET','POST'])
        self.add_url_rule('/new_cus', view_func=self.new_cus,methods=['GET','POST'])
        self.add_url_rule('/get_template_info', view_func=self.get_template_info,methods=['GET','POST'])
        self.add_url_rule('/input_buy', view_func=self.input_buy,methods=['GET','POST'])
        self.add_url_rule('/write_buy', view_func=self.write_buy,methods=['GET','POST'])
        self.add_url_rule('/success', view_func=self.success,methods=['GET','POST'])
        self.add_url_rule('/get_cus_buy', view_func=self.get_cus_buy,methods=['GET','POST'])
        self.add_url_rule('/get_train_list', view_func=self.get_train_list,methods=['GET','POST'])
        self.add_url_rule('/deal_cls', view_func=self.deal_cls,methods=['GET','POST'])
        
    def index(self):
        return render_template('index.html')

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
        print(dic_tkn)
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
            return f'\n写入训练记录表错误：{e}'
        try:
            df_copy_rows= df_copy_rows_nonOxy.loc[df_train_rec.index.repeat(df_train_rec['oxyGroup'].astype(int))]
        except Exception as e:
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
            print('err',e)
            return 'err:'+e
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
            return {'dat':'get_cus_buy error','error':err}      
   

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

        return f'写入成功, 行号：{row}'


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
        return txt_num


    def generate_new(self):
        try:
            fn_in=request.data
            fn='MH'+fn_in.decode('utf-8')
            fn,dvc=fn.split('|')
            work_dir=self.wecom_dir()
            tplt_dir=os.path.dirname(work_dir)
            new_fn=os.path.join(work_dir,fn+'.xlsm')

            app=xw.App(visible=False)
            wb=app.books.open(os.path.join(tplt_dir,'模板.xlsm'))
            sht=wb.sheets['基本情况']
            sht['A2'].value=fn[0:5]
            sht['B2'].value=fn[5:]
            if len(fn[5:])>1:
                sht['C2'].value=fn[5:][1:]
            else:
                sht['C2'].value=fn[5:]

            wb.save(new_fn)
            wb.close()
            app.quit()

            # os.startfile(work_dir)
            if dvc=='pc':
                os.startfile(new_fn)

            return new_fn
        except Exception as e:
            return e


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
