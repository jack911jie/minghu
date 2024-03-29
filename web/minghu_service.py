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
import threading
from multiprocessing import Value
import xlwings as xw
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
from flask import Flask, request, jsonify,render_template,session,redirect,url_for,send_file
import pymysql
import datetime
from decimal import Decimal
import json
import random
import hashlib
import base64


class MinghuService(Flask):

    def __init__(self,*args,**kwargs):
        super(MinghuService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.join(os.path.dirname(__file__),'config','lz_service.config'))
        # self.config_lz=readconfig.exp_json2(config_fn)
        with open(config_fn,'r',encoding='utf-8') as f:
            self.config_lz=json.load(f)
        # print(self.config_lz)
        self.app_lock=threading.Lock()
        self.cus_id_prefix=self.config_lz['cus_id_prefix']
        

        #路由
        #渲染页面
        #首页/登录页
        self.add_url_rule('/',view_func=self.login)
        
        #菜单页
        self.add_url_rule('/index',view_func=self.index)     
        #修改密码页
        self.add_url_rule('/reedit_pwd',view_func=self.reedit_pwd_page)      
        #获取客户信息页面
        self.add_url_rule('/cus_infos',view_func=self.cus_infos)
        #欢迎页面
        self.add_url_rule('/welcome',view_func=self.welcome)
        #欢迎页面
        self.add_url_rule('/test_copy',view_func=self.test_copy)
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
        # 上课预约页面
        self.add_url_rule('/ins_book', view_func=self.ins_book_page,methods=['GET','POST'])
        
        
               

        #功能
        #登出处理
        self.add_url_rule('/logout',view_func=self.logout)
        #修改密码处理
        self.add_url_rule('/deal_reedit_pwd',view_func=self.deal_reedit_pwd_db,methods=['GET','POST'])
        #从模板.xlsm获取基本信息，如教练姓名、课程种类等
        self.add_url_rule('/get_template_info', view_func=self.get_template_info_db,methods=['GET','POST'])
        #遍历会员资料文件夹获取所有 客户列表
        self.add_url_rule('/get_cus_list', view_func=self.get_cus_list_db,methods=['GET','POST'])
        #从,'03-教练管理','教练资料','教练信息.xlsx'中获取教练列表
        self.add_url_rule('/get_ins_list', view_func=self.get_ins_list_db,methods=['GET','POST'])
        #获取客户信息，包括既往购课、上课信息，剩余课时信息，限时课程信息，围度测量信息
        self.add_url_rule('/get_cus_info', view_func=self.get_cus_info_db,methods=['GET','POST'])
        #获取客户未开课的购课编码
        self.add_url_rule('/deal_start_limit_page', view_func=self.deal_start_limit_page_db,methods=['GET','POST'])
        #通过'05-专业资料','训练项目.xlsx'获取训练项目的名称及分类
        self.add_url_rule('/get_train_list', view_func=self.get_train_list_db,methods=['GET','POST'])    
        # 获取体验课记录
        self.add_url_rule('/get_trial_list', view_func=self.get_trial_list_db,methods=['GET','POST']) 
        #获取客户既往购课记录,并整理合并
        # 获取有效的卡
        self.add_url_rule('/get_cards', view_func=self.get_cards,methods=['GET','POST']) 
        #获取客户既往购课记录,并整理合并
        #按不同的类型获取有效的卡号返回给购卡页面
        self.add_url_rule('/get_cus_cards_by_type', view_func=self.get_cus_cards_by_type,methods=['GET','POST'])
        #获取限时卡记录返回给限时课程开课页面
        self.add_url_rule('/send_to_start_lmt_page_data',view_func=self.send_to_start_lmt_page_data,methods=['GET','POST'])
        self.add_url_rule('/get_cus_buy', view_func=self.get_cus_buy_db,methods=['GET','POST'])
        #获取既往体测记录
        self.add_url_rule('/get_body_history', view_func=self.get_body_history_db,methods=['GET','POST'])


        #打开客户的xlsm文件
        self.add_url_rule('/open_cus_fn', view_func=self.open_cus_fn,methods=['GET','POST'])
        #遍历会员资料文件夹，生成新的客户ID号
        # self.add_url_rule('/check_new', view_func=self.check_new,methods=['GET','POST'])

        #生成新会员
        self.add_url_rule('/generate_new', view_func=self.generate_new_db,methods=['GET','POST'])
        #写入购课记录
        self.add_url_rule('/write_buy', view_func=self.write_buy_db,methods=['GET','POST'])
        
        #获取客户既往购课记录表，不整理
        self.add_url_rule('/get_cus_buy_list', view_func=self.get_cus_buy_list,methods=['GET','POST'])        
           
        #执行写入上课记录、写入训练记录
        self.add_url_rule('/deal_cls', view_func=self.deal_cls_db,methods=['GET','POST'])
        
        #写入体测记录
        self.add_url_rule('/write_body', view_func=self.write_body_db,methods=['GET','POST'])
        #处理开课记录
        self.add_url_rule('/deal_start_class_page', view_func=self.deal_start_class_page_db,methods=['GET','POST'])
        
        # 写入体验课上课记录
        self.add_url_rule('/write_trial_rec', view_func=self.write_trial_rec_db,methods=['GET','POST'])
        #查询约课记录
        self.add_url_rule('/get_book_data',view_func=self.get_book_data_db,methods=['GET','POST'])
        #写入约课记录表
        self.add_url_rule('/write_ins_book',view_func=self.write_ins_book_db,methods=['GET','POST'])
        #处理login
        self.add_url_rule('/deal_login',view_func=self.deal_login,methods=['GET','POST'])
        #返回小程序首页显示教练的控制
        self.add_url_rule('/show_who_ins',view_func=self.show_who_ins,methods=['GET','POST'])
        #返回微信小程序约课页面下方广告图片
        self.add_url_rule('/get_wxprg_ads_pic',view_func=self.get_wxprg_ads_pic,methods=['GET','POST'])
        self.add_url_rule('/send_ads_pic',view_func=self.send_ads_pic,methods=['GET','POST'])
        

    def connect_mysql(self):
        with open(os.path.join(os.path.dirname(os.path.realpath((__file__))),'config','db.config'),'r',encoding='utf-8') as f:
            cfg=json.load(f)
        # 连接数据库
        conn = pymysql.connect(
            host=cfg['host'],       # 数据库主机地址
            user=cfg['user'],     # 数据库用户名
            password=cfg['password'], # 数据库密码
            database=cfg['database'],  # 要连接的数据库名称
            port=cfg['port']
        )

        return conn

    def show_who_ins(self):
        print('sending show who ins...')
        config_show_who_fn=os.path.join(os.path.dirname(os.path.realpath(__file__)),'config','show_who.config')
        with open(config_show_who_fn,'r',encoding='utf-8') as f:
            config_show_who=json.load(f)
        return jsonify(config_show_who)
    
    def reedit_pwd_page(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return  render_template('./reedit_pwd.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def deal_reedit_pwd_db(self):
        print('deal reedit pwd')
        data=request.json

        user_id=data['user_id']
        edited_pwd=data['edited_pwd']
        salt=str(random.randint(10000000,99999999))

        pwd_salt=hashlib.sha256((edited_pwd + salt).encode()).hexdigest()
        conn=self.connect_mysql()
        cursor=conn.cursor()
        try: 
            #如果没有记录，写入salt值，如有，更新。
            sql=f'select  ins_id from salt_table where ins_id=%s'
            cursor.execute(sql,(user_id))
            res=cursor.fetchone()
            if res:
                sql=f'update salt_table set salt=%s where ins_id=%s'
                cursor.execute(sql,(salt,user_id))
            else:
                sql=f'insert into salt_table (ins_id,salt) values (%s,%s)'
                cursor.execute(sql,(user_id,salt))                

            sql=f'update ins_table set pwd=%s, first_login=0 where ins_id=%s'
            cursor.execute(sql,(pwd_salt,user_id))

            conn.commit()
            cursor.close()
            conn.close()
            return jsonify({'res':'ok'})            
        except Exception as e:
            print(e)
            cursor.close()
            conn.close()
            return jsonify({'res':'failed'})
            
    def get_wxprg_ads_pic(self):    
        with open (os.path.join(os.path.dirname(__file__),'config','wxprg_ads.config'),'r',encoding='utf-8') as f:
            ads_url=json.load(f)
        
        return {'ads_urls':ads_url}

    def send_ads_pic(self):
        return render_template('./ads.html')

    def login(self):
        return render_template('./login.html')

    def deal_login(self):
        data=request.json
        mobile,pwd=data['user'],data['pwd']
        conn=self.connect_mysql()



        with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
            print('matching login information...')
            try:
                #通过手机号码读取ID
                sql='select ins_id,ins_name from ins_table where mobile=%s'
                cursor.execute(sql,(mobile))
                res=cursor.fetchone()
                print(res)
                if res:
                    ins_id=str(res['ins_id'])
                else:
                    raise ValueError('没有手机号对应的ID')

                #读取salt
                sql='select ins_id,salt from salt_table where ins_id=%s'
                
                cursor.execute(sql,(ins_id))
                res=cursor.fetchone()
                # print('226',res)
                if res:
                    salt=str(res['salt'])
                    pwd=hashlib.sha256((pwd + salt).encode()).hexdigest()
                    # hashed_password = hashlib.sha256((password + salt).encode()).hexdigest()

                sql='select ins_name,role,first_login from ins_table where ins_id=%s and pwd=%s'            
                cursor.execute(sql,(ins_id,pwd))
                res=cursor.fetchone()
                print(res)
                session['user']=res['ins_name']
                session['role']=res['role']
                session['ins_id']=ins_id

                # print(res['ins_name'],res['role'])
                if int(res['first_login'])==1:
                    return jsonify({'res':'ok','url':url_for('reedit_pwd_page'),'ins_id':ins_id,'reedit_pwd':'yes'})
                else:
                    return jsonify({'res':'ok','url':url_for('index'),'ins_id':ins_id,'reedit_pwd':'no'})
            except Exception as e:
                print(e)
                return jsonify({'res':'failed'})

    def get_cards(self):
        print('get valid cards')
        data=request.json
        cus_id=data['cus_id_name'][:7]

        conn=self.connect_mysql()
        cursor=conn.cursor()
        #获取该会员上课次数
        sql='select count(card_id) as tkn_qty from cls_tkn_rec_table where cus_id=%s'
        cursor.execute(sql,cus_id)
        qty=cursor.fetchone()[0]
        
        today=datetime.datetime.now().strftime('%Y-%m-%d')


        sql='''
        select card_id,card_type,card_name,cls_qty,card_start_time,prd,end_time,cmt from
        cards_table 
        where card_id in (select card_id from cardholder_card_table where cus_id=%s) and cls_qty>=%s and end_time>=%s;

        '''
        try:
            cursor.execute(sql,(cus_id,qty,today))
            res=cursor.fetchall()
            print(res)

            return {'res':'ok','valid_cards':res}
        except:
            print('get cards id failed')
            return {'res':'failed'}



    def logout(self):
        session.pop('user',None)
        session.pop('role',None)
        session.pop('ins_id',None)
        return redirect(url_for('login'))

    def write_ins_book_db(self):
        try:
            data=request.json
            conn=self.connect_mysql()
            cursor=conn.cursor()

            # print(data)

            #通过教练ins_id获取ins_name
            
            sql=f'''
                SELECT ins_id,ins_name FROM ins_table WHERE ins_id=%s ;
            '''
            cursor.execute(sql,data['insId'])
            data['insName']=cursor.fetchone()[1]
            
            col_names=['date','insId','insName','0600','0630','0700','0730','0800','0830','0900','0930','1000','1030','1100','1130','1200','1230','1300','1330','1400','1430','1500','1530','1600','1630','1700','1730','1800','1830','1900','1930','2000','2030','2100','2130','comment']
            sorted_data={key:data[key] for key in col_names}
            values=list(sorted_data.values())
            # print(values)
            #查询是否已有数据
            sql=f'''
                SELECT ins_id,ins_name FROM ins_book_table WHERE ins_id=%s and date=%s;
            '''
            cursor.execute(sql,[data['insId'],data['date']])
            res=cursor.fetchall()
            #如已有教练、日期 数据，更新
            if res:
                #update的时候，加上这些参数
                values.extend([data['insId'],data['insName'],data['date']])
                # print(values)
                print('old data,updating')     
                sql=f'''
                        update ins_book_table
                        set 
                        date=%s,ins_id=%s,ins_name=%s,`0600`=%s,`0630`=%s,`0700`=%s,`0730`=%s,`0800`=%s,`0830`=%s,`0900`=%s,`0930`=%s,`1000`=%s,`1030`=%s,`1100`=%s,`1130`=%s,`1200`=%s,`1230`=%s,`1300`=%s,`1330`=%s,`1400`=%s,`1430`=%s,`1500`=%s,`1530`=%s,`1600`=%s,`1630`=%s,`1700`=%s,`1730`=%s,`1800`=%s,`1830`=%s,`1900`=%s,`1930`=%s,`2000`=%s,`2030`=%s,`2100`=%s,`2130`=%s,comment=%s
                        where
                        ins_id=%s and ins_name=%s and date=%s
                    '''
                cursor.execute(sql,values)
                conn.commit()
                cursor.close()
                conn.close()     
                return  jsonify({'res':'old data:update in booking table successfully.'})  
            #如无教练、日期 数据，插入
            else:           
                print('new data,inserting')     
                
                sql=f'''
                        insert into ins_book_table
                        (date,ins_id,ins_name,`0600`,`0630`,`0700`,`0730`,`0800`,`0830`,`0900`,`0930`,`1000`,`1030`,`1100`,`1130`,`1200`,`1230`,`1300`,`1330`,`1400`,`1430`,`1500`,`1530`,`1600`,`1630`,`1700`,`1730`,`1800`,`1830`,`1900`,`1930`,`2000`,`2030`,`2100`,`2130`,comment)
                        values
                        (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    '''
                cursor.execute(sql,values)
                conn.commit()
                cursor.close()
                conn.close()     
                return  jsonify({'res':'new data:write in booking table successfully.'})      
        except Exception as e:
            print('write into ins book table error: ',e)
            return  jsonify({'res':'error'+e})
        
       

    def get_book_data_db(self):
        print('from minghu database,get ins book data')
        data=request.json
        # print(data)
        conn=self.connect_mysql()
        cursor=conn.cursor()
        sql=f'''
            SELECT ins_name FROM ins_table WHERE ins_id=%s ;
        '''
        cursor.execute(sql,data['ins_id'])
        res_ins_name=cursor.fetchone()[0]

        sql=f'''
            SELECT * FROM ins_book_table WHERE ins_id=%s and ins_name=%s and date=%s;
        '''

        cursor.execute(sql,[data['ins_id'],res_ins_name,data['date']])
        # print([data['ins_name'],res_ins_id,data['date']])
        ins_book_data=cursor.fetchall()
        # print(ins_book_data)
        if ins_book_data:
            cols=['id','date','ins_id','ins_name','0600','0630','0700','0730','0800','0830','0900','0930','1000','1030','1100','1130','1200','1230','1300','1330','1400','1430','1500','1530','1600','1630','1700','1730','1800','1830','1900','1930','2000','2030','2100','2130','comment']
            ins_book_dic={}
            for n,col in enumerate(cols):
                ins_book_dic[col]=ins_book_data[0][n]
        else:
            ins_book_dic=''

        cursor.close()
        conn.close()
        # print(ins_book_dic)
        return jsonify({'ins_book_data':ins_book_dic})

    def ins_book_page(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('./ins_book.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def test_copy(self):
        return render_template('./test_copy.html')
  

    def write_trial_rec(self):
        try:
            data=request.json
            fn=os.path.join(self.config_lz['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
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
        
    def write_trial_rec_db(self):
        try:
            data=request.json
            conn=self.connect_mysql()
            cursor=conn.cursor()
            sql=f'select ins_id,ins_name from ins_table where ins_id=%s'
            cursor.execute(sql,(data['insId']))
            ins_name_res=cursor.fetchone()
            data['insName']=ins_name_res[1]

            data['trial_cls_long']=1
            data['datetime']=data['dateString']+' '+data['timeString']
            data['finish_yn']='是'
            data['deal_yn']='否'
            data['deal_date']=None
            data['formal_cus_id_name']=None
            del data['dateString']
            del data['timeString']
            data_cols=['datetime','trial_cls_long','cusNameInput','mobilePhone','insName','finish_yn','cusSource','comment','deal_yn','deal_date','formal_cus_id_name']
            sorted_data={key:data[key] for key in data_cols}
            values=tuple(sorted_data.values())
            # values = ', '.join(f'"{data[key]}"' if sorted_data[key] is not None else 'NULL' for key in data_cols)
            # print(values)            
            
            sql=f'''
                    insert into trial_cls_table (trial_datetime,trial_cls_long,trial_cus_name,trial_cus_mobile,ins_name,finish_yn,trial_cus_source,comment,deal_yn,deal_date,formal_cus_id_name)
                    values 
                    (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                '''
            cursor.execute(sql,values)
            conn.commit()
            cursor.close()
            conn.close()

        except Exception as e:
            print('写入体验课表错误：',e)
            return '写入体验课表成功'+e

        return '写入体验课表成功'


    def trial_class(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('./trial_class.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)
         
    def date_to_string(self,obj,format='date'):
        try:
            if format=='date':
                return obj.strftime('%Y-%m-%d')
            elif format=='time':
                return obj.strftime('%H:%M')+':00'
        except:
            return None

        raise TypeError(f"Object of type {obj.__class__.__name__} is not JSON serializable")

    def get_trial_list(self):
        
        df_trial=pd.read_excel(os.path.join(self.config_lz['work_dir'],'03-教练管理','体验课上课记录表.xlsx'),sheet_name='体验课上课记录表')
        df_trial.dropna(subset=['体验客户姓名'],inplace=True)
        df_trial.fillna('',inplace=True)
        df_trial['体验课日期']=df_trial['体验课日期'].apply(lambda x:self.date_to_string(x,'date'))
        df_trial['体验课时间']=df_trial['体验课时间'].apply(lambda x:self.date_to_string(x,'time'))
        df_trial['出单日期']=df_trial['出单日期'].apply(lambda x:self.date_to_string(x,'date'))
        print(df_trial.dtypes)
        dic_trial=df_trial.to_dict()
        num_keys=len(dic_trial[list(dic_trial.keys())[0]])
        result={}
        for i in range(num_keys):
            result[i] = {key: dic_trial[key][i] for key in dic_trial}
        # print(result)
        print(type(result))

        # dic_trial=df_trial.T.to_dict()

        return jsonify({'all_trial':result})

    def get_trial_list_db(self):
        conn=self.connect_mysql()
        cursor = conn.cursor()

        # 执行SQL查询
        # sql = "SELECT * FROM `id_table`"
        sql="select * from trial_cls_table"
        cursor.execute(sql)

        # # 获取查询结果
        results = cursor.fetchall()
        # data=((1, datetime.datetime(2022, 5, 14, 9, 0), 1, 'TS0501唐家源', None, '韦越棋', '是', '小红书', None, '是', datetime.date(2022, 5, 22), None), (2, datetime.datetime(2022, 5, 23, 19, 30), 1, 'TS0502毛圆', None, '韦越棋', '是', '小红书', None, None, None, None))
        df=pd.DataFrame(results)
        df.columns=['序号','体验课日期时间','时长','体验客户姓名','体验客户手机','教练','是否完成','客户来源','备注','是否出单','出单日期','出单后正式会员编码及姓名']
        df['体验课日期']=df['体验课日期时间'].dt.strftime('%Y-%m-%d')
        df['体验课时间']=df['体验课日期时间'].dt.strftime('%H:%M:%S')
        df=df[['序号','体验课日期','体验课时间','时长','体验客户姓名','体验客户手机','教练','是否完成','客户来源','备注','是否出单','出单日期','出单后正式会员编码及姓名']]

        dic_trial_list=df.T.to_dict()

        # 关闭游标和连接
        cursor.close()
        conn.close()

        return jsonify({'all_trial':dic_trial_list})
           

    def deal_start_class_page(self):
        data=request.json
        
        try:
            self.add_rec_to_start_class_table(dic=data)
            self.delete_rec_in_aux_table(dic=data)
            return jsonify({'result':'写入限时课程表及辅助表成功'})
        except Exception as e:
            print('写入限时课程表或辅助表错误：',e)
            return jsonify({'result':'写入限时课程表及辅助表成功错误'+e})

    def deal_start_class_page_db(self):    
        try:
            data=request.json
            print(data)
            conn=self.connect_mysql()
            cursor=conn.cursor()

            data['cus_id']=data['cus_name_input'][:7].strip()
            data['cus_name']=data['cus_name_input'][7:].strip()
            del data['cus_name_input']
            data_cols=['cus_id','cus_name','card_id','start_date','end_date','ins_id','opr_id','opr_type','prd','cls_qty']
            sorted_data={key: data[key] for key in data_cols}
            # print(sorted_data)
            values=tuple(sorted_data.values())
            
            # sql='''
            #     insert into lmt_cls_rec_table
            #     (cus_id,cus_name,buy_code,start_date,end_date,ins_id,operator_id)
            #     values
            #     (%s,%s,%s,%s,%s,%s,%s)
            # '''
            current_time=datetime.datetime.now()
            if sorted_data['opr_type']=='新课':
                sql='''
                update cards_table 
                set card_start_time=%s,end_time=%s,opr_time=%s
                where card_id=%s
                '''

                cursor.execute(sql,(sorted_data['start_date'],sorted_data['end_date'],current_time,sorted_data['card_id']))
                conn.commit()
            else: #延长
                sql='''
                insert into cards_table 
                (card_id,card_type,card_name,cls_qty,card_start_time,prd,end_time,cmt,opr_id,opr_time)
                values
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    
                '''
                values=(sorted_data['card_id'],'limit_prd','私教',sorted_data['cls_qty'],
                        sorted_data['start_date'],sorted_data['prd'],sorted_data['end_date'],'',
                        sorted_data['opr_id'],current_time)
                cursor.execute(sql,(values))
                conn.commit()


            # #从未开课表中删除对应的购课记录，如无，mysql也不会报错
            # sql=f'''
            #     delete from  not_start_lmt_table
            #     where
            #     buy_code=%s
            # '''
            # cursor.execute(sql,(sorted_data['buyCode']))

            # conn.commit()
            # cursor.close()
            # conn.close()



            cursor.close()
            conn.close()
            return jsonify({'res':'ok'})
        except Exception as e:
            print('写入限时课程表或辅助表错误：',e)
            return jsonify({'res':'failed','info':'写入限时课程表及删除未开课表错误'+e})
        

   
    def add_rec_to_start_class_table(self,dic):
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',dic['cusName'].strip()+'.xlsm')

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
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',buy_code[:-8].strip()+'.xlsm')
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
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',dic['cusName'].strip()+'.xlsm')
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
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('index.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def start_limit_class(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('./start_limit_class.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    
    def write_body(self):
        try:
            data=request.json
            # print('写入身体数据：',data)
            fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',data['cusName']+'.xlsm')
            df_old=pd.read_excel(os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',data['cusName']+'.xlsm'),sheet_name='身体数据')
            df_old.dropna(subset=['日期'],inplace=True)
            df_new=pd.DataFrame(data,index=[0])

            #第一列为姓名，去除
            df_write=df_new.iloc[:,1:]
            df_write['bfr']=''
            df_write=df_write[['date','ht','wt','bfr','chest','l_arm','r_arm','waist','hip',
                                'l_leg','r_leg','l_calf','r_calf',
                                'heart','balance','power','flex','core']]
            # df_write

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
            print('后端：写入身体数据成功')
            return '后端：写入身体数据表成功'
        except Exception as e:
            print('后端写入身体数据错误：',e)
            return '后端写入身体数据错误：'+e

    def write_body_db(self):
        try:
            data=request.json
            conn=self.connect_mysql()
            cursor=conn.cursor()
            cus_id,cus_name=data['cusName'][:7],data['cusName'][7:]
            sql=f"select sex,birthday from cus_info_table where cus_id='{cus_id}' and cus_name='{cus_name}'"
            cursor.execute(sql)
            res=cursor.fetchall()
            sex=res[0][0]
            birthday=res[0][1].strftime('%Y%m%d')
            bfr=self.bfr(sex,birthday,float(data['ht']),float(data['wt']),float(data['waist']))
            # print(sex,birthday,bfr)
            data['bfr']=bfr
            data['cus_id']=cus_id
            data['cus_name']=cus_name
            del data['cusName']
            sorted_cols=['cus_id','cus_name','date','ht','wt','bfr','chest','l_arm','r_arm','waist','hip','l_leg','r_leg','l_calf','r_calf','heart','balance','power','flex','core']
            sorted_data={key:data[key] for key in sorted_cols}
            for key,value in sorted_data.items():
                try:
                    sorted_data[key]=float(value)
                except:
                    pass
            
            value=tuple(sorted_data.values())
            sql=f"insert into body_msr_table (cus_id,cus_name,msr_date,ht,wt,bfr,chest,l_arm,r_arm,waist,hip,l_leg,r_leg,l_calf,r_calf,heart,balance,power,flex,core) values {value}"
            # print(value,sql)
            cursor.execute(sql)
            conn.commit()
            cursor.close()
            conn.close()
            #返回字符必须包含“成功”二字供前端
            return '写入体测数据成功' 

        except Exception as e:
            print('write body error in write_body_db():',e)
            return  'error'

    def bfr(self,sex,birthday,ht,wt,waist):
        bfr_test=get_data.cals()       
        
        if birthday:
           # 可能的日期格式列表
            date_formats = ['%Y', '%Y%m', '%Y%m%d']
            for date_format in date_formats:
                try:
                    birthday = datetime.datetime.strptime(str(birthday), date_format)
                    age = relativedelta(datetime.datetime.now(), birthday).years
                    bfr = bfr_test.bfr(age=age, sex=sex, ht=ht, wt=wt, waist=waist, adj_bfr='no', adj_src='prg', formula=1)
                    return bfr
                except ValueError as e:
                    # print('bfr计算错误 in bfr():',e)
                    continue
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
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name+'.xlsm')
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
            print('bfr计算错误 in get_body_history()',e)
            item['体脂率']='-'

        # print(formatted_data)
        # print(dic_body)
        return jsonify(formatted_data)
    
    def get_body_history_db(self):
        # cus_id_name='MH00008苏云'
        cus_id_name=request.data.decode('utf-8')
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]

        conn=self.connect_mysql()
        cursor=conn.cursor()
        sql=f"select * from body_msr_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        body_history=cursor.fetchall()
        body_history_str=[]
        if body_history:
            for r,row in enumerate(body_history):
                row_body_history=[]
                for c,value in enumerate(row):
                    if isinstance(value,datetime.date):
                        row_body_history.append(value.strftime('%Y-%m-%d'))
                    elif isinstance(value,Decimal):
                        row_body_history.append(float(value))
                    else:
                        row_body_history.append(value)
                body_history_str.append(row_body_history)        
            body_history_cols=['id','cus-id','cus_name','日期','身高（cm）','体重（Kg）','体脂率','胸围','左臂围','右臂围','腰围','臀围','左腿围','右腿围','左小腿围','右小腿围','心肺','平衡','力量','柔韧性','核心']
            dic_body_history=self.mysql_list_data_to_dic(data=body_history_str,mysql_cols=body_history_cols)
        else:
            dic_body_history={'0': {'id':'','cus-id':'','cus_name':'','日期': '', '身高（cm）': '', '体重（Kg）': '', '体脂率': '', 
                        '胸围': '', '左臂围': '', '右臂围': '', '腰围': '', '臀围': '', '左腿围': '', 
                        '右腿围': '', '左小腿围': '', '右小腿围': '', '心肺': '', '平衡': '', 
                        '力量': '', '柔韧性': '', '核心': ''}}
        

        cursor.close()
        conn.close()

        return jsonify(dic_body_history)
    
    def input_body(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('./input_body.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

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

    def deal_cls_db(self):
        try:
            data=request.json
            conn=self.connect_mysql()
            cursor=conn.cursor()

            #教练上课记录
            cls_tkn_rec=data['cls_tkn']
            cls_tkn_rec['cus_id']=cls_tkn_rec['cus_name'][:7]
            cls_tkn_rec['cus_name']=cls_tkn_rec['cus_name'][7:]
            cls_tkn_rec['cls_datetime']=cls_tkn_rec['cls_tkn_date']+' '+cls_tkn_rec['cls_tkn_time']
            cus_id=cls_tkn_rec['cus_id']
            cus_name=cls_tkn_rec['cus_name']
            card_id=cls_tkn_rec['card_id']
            ins_id=cls_tkn_rec['ins_id']
            ins_name=cls_tkn_rec['ins_name']
            basic_cls_comment=cls_tkn_rec['basic_cls_comment']
            train_datetime=cls_tkn_rec['cls_datetime']

            del cls_tkn_rec['cls_tkn_date']
            del cls_tkn_rec['cls_tkn_time']

            cls_tkn_data_cols=['cus_id','cus_name','cls_datetime','cls_long','cls_type','card_id','ins_id','ins_name','basic_cls_comment','operator_id','operate_time']
            sorted_cls_tkn_data={key: cls_tkn_rec[key] for key in cls_tkn_data_cols} 
            values_cls_tkn=tuple(sorted_cls_tkn_data.values())
            
            
            sql_cls_tkn=f'''
                insert into cls_tkn_rec_table
                (cus_id,cus_name,cls_datetime,cls_long,cls_type,card_id,ins_id,ins_name,comment,operator_id,operate_time) 
                values 
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
            cursor.execute(sql_cls_tkn,values_cls_tkn)

            #训练记录
            train_data=data['train_rec']
            #替换''为'0'
           
            

            calories=train_data['calories']
            train_comment=train_data['trainComment']
            train_items=train_data['train_recs']

            
           
            oxy_items=[]
            non_oxy_items=[]
            for train_item in train_items:
                # 抗阻运动
                non_oxy_row=[]
                oxy_row=[]
                if train_item['nonOxyName']:   
                    action_name=train_item['nonOxyName']
                    sql=f'select muscle from train_item_table where action_name="{action_name}"'
                    cursor.execute(sql)
                    muscle=cursor.fetchall()[0][0]
                    non_oxy_row.extend([train_datetime,cus_id,cus_name,muscle,action_name,train_item['nonOxyWt'],
                                            train_item['nonOxyDis'],train_item['nonOxyNum'],train_item['nonOxyGroup'],
                                            calories,card_id,ins_id,ins_name,train_comment])
                if non_oxy_row:
                    non_oxy_row[-1]=non_oxy_row[-1] if non_oxy_row[-1] else None
                    non_oxy_items.append(non_oxy_row)

                if train_item['oxyName']:                   
                    oxy_row.extend([train_datetime,cus_id,cus_name,train_item['oxyName'],train_item['oxyTime'],
                                    train_item['oxyGroup'],calories,card_id,ins_id,ins_name,train_comment])
                if oxy_row:
                    oxy_row[-1]=oxy_row[-1] if oxy_row[-1] else None
                    oxy_items.append(oxy_row)

            #将''替换为'0',再将数字转换为浮点数      
            oxy_items = [[item if item != '' else '0' for item in sublist] for sublist in oxy_items]     
            non_oxy_items = [[item if item != '' else '0' for item in sublist] for sublist in non_oxy_items]    
            # #评论为空                  
            # oxy_items[-1]=''        
            # non_oxy_items[-1]=''
            try:
                oxy_items = [[float(item) if item.replace('.', '', 1).isdigit() else item for item in sublist] for sublist in oxy_items]
                non_oxy_items = [[float(item) if item.replace('.', '', 1).isdigit() else item for item in sublist] for sublist in non_oxy_items]
            except:
                pass
            
            # print(oxy_items,non_oxy_items)

            #无氧训练写入
            non_oxy_train_sql=f'''
                insert into train_nonoxy_rec_table 
                (train_datetime,cus_id,cus_name,muscle,non_oxy_name,non_oxy_wt,non_oxy_dis,non_oxy_num,non_oxy_group,calories,card_id,ins_id,ins_name,comment) 
                values 
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);
            '''
            cursor.executemany(non_oxy_train_sql,non_oxy_items)
            
            #有氧训练写入
            oxy_train_sql=f'''
                insert into train_oxy_rec_table 
                (train_datetime,cus_id,cus_name,oxy_name,oxy_time,oxy_group,calories,card_id,ins_id,ins_name,comment) 
                values 
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);

            '''
            cursor.executemany(oxy_train_sql,oxy_items)

            conn.commit()
            cursor.close()
            conn.close()


            return '写入上课记录及训练记录成功'
        except Exception as e:
            print('deal_cls_db() ERROR',e)
            return  '写入上课记录及训练记录错误'



    def train_info(self,action_name,dic):
    # print('162----',action_name,dic)
        if action_name:
            return dic['by_action_name'][action_name]
        else:
            return ['','','']

    
    def write_cls_tkn(self,dic_tkn):
        # print(dic_tkn)
        try:
            fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',dic_tkn['cus_name']+'.xlsm')
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
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name+'.xlsm')
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
        res=os.path.join(self.config_lz['work_dir'].strip(),'01-会员管理','会员资料')
        return res

    def get_cus_buy(self):
        cus_name=request.data.decode('utf-8')
        print(cus_name)
        
        # print(fn)
        try:
            fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
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
   
    def get_cus_buy_db(self):
        # cus_name=request.data.decode('utf-8')
        print('from minghu database,get_cus_buy_statistics')
        cus_id_name=request.data.decode('utf-8')
        # cus_id_name='MH00113肖婕'
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]

        conn=self.connect_mysql()
        cursor=conn.cursor()
        sql=f'''SELECT 
                    buy_code,
                    AVG(pay) AS 平均应收金额,
                    SUM(real_pay) AS 总实收金额,
                    MIN(buy_type) AS 购课类型,
                    GROUP_CONCAT(DISTINCT DATE_FORMAT(buy_date, '%Y/%m/%d') ORDER BY buy_date ASC SEPARATOR '\n') AS 收款日期列表,
                    COUNT(*) AS 收款次数
                FROM
                    buy_rec_table
                WHERE
                    cus_name='{cus_name}' and cus_id='{cus_id}'
                GROUP BY
                    buy_code;
        '''
        cursor.execute(sql)
        buy_stat=cursor.fetchall()
        if buy_stat:
            df=pd.DataFrame(buy_stat)
            df.columns=['购课编码','应收金额','实收金额','购课类型','收款日期','收款次数']
            df['应收金额']=df['应收金额'].astype(float)
            df['实收金额']=df['实收金额'].astype(float)
            df['未收金额']=df['应收金额']-df['实收金额']
            df=df[['购课编码','购课类型','应收金额','实收金额','未收金额','收款次数','收款日期']]

            buy_stat_list=[row.tolist() for row in df.values]        
            return jsonify({'buy_stat_list':buy_stat_list})
        else:
            return jsonify({'buy_stat_list':None})

    def get_cus_buy_list(self,cus_name):
        # cus_name=request.data.decode('utf-8')
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
        df=pd.read_excel(fn,sheet_name='购课表')
        # dic_buy=df.to_dict()
        # dic_res=self.dic_format(dic=dic_buy,order_name='收款日期')
        # return jsonify(dic_res)
        return df

    def  get_limit_class_records(self,cus_name):
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
        df=pd.read_excel(fn,sheet_name='限时课程记录')

        df.fillna('',inplace=True)
        return df

    def get_not_start_lmt_list(self,cus_name):
        # cus_name=request.data.decode('utf-8')
        fn=os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料',cus_name.strip()+'.xlsm')
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
            # pd无法直接将NaT替换为空值，将其替换为0后，再转为''
            df_limit_cls_recs.fillna(0, inplace=True)
            df_limit_cls_recs.replace(0,'',inplace=True)     

            if df_limit_cls_recs.empty:
                dic_limit_cls_recs=''
                dic_limit_maxdate_rec={'购课编码':'','限时课程起始日':'','限时课程结束日':'','限时课程实际结束日':''}
                dic_limit_maxdate_rec={'0':dic_limit_maxdate_rec}
            else:   
                df_limit_maxdate_rec=df_limit_cls_recs[df_limit_cls_recs['限时课程结束日']==df_limit_cls_recs['限时课程结束日'].max()]
                # df_limit_maxdate_rec.fillna('',inplace=True)
                df_limit_maxdate_rec.reset_index(inplace=True)
                # print('df_limit_maxdate_rec:',df_limit_maxdate_rec)

                dic_limit_cls_recs=df_limit_cls_recs.to_dict()
                dic_limit_cls_recs=self.dic_format(dic=dic_limit_cls_recs,order_name='购课编码')


                dic_limit_maxdate_rec=df_limit_maxdate_rec.to_dict()
                dic_limit_maxdate_rec=self.dic_format(dic=dic_limit_maxdate_rec,order_name='购课编码')
            # print('df limit cls recs, df limit maxdate rec: \n',df_limit_cls_recs,df_limit_maxdate_rec) 
            # print('dic_limit_maxdate_rec:',dic_limit_cls_recs)

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
            # print('deal_start_limit_page()',df_merge)

            if df_merge.empty:
                dic_not_start=''
            else:      
                df_merge.fillna('',inplace=True)
                df_merge.sort_values(by='收款日期',ascending=True,inplace=True)
                dic_not_start=df_merge.to_dict()
                # print('\n479 line dic_not_start',dic_not_start)
                dic_not_start=self.dic_format(dic=dic_not_start,order_name='收款日期')
            
            return jsonify({'not_start_list':dic_not_start,'buy_list':dic_buy,'limit_cls_recs':dic_limit_cls_recs,'maxdate_limit_class_rec':dic_limit_maxdate_rec})
        except Exception as e:
            return jsonify({'error':e,'not_start_list':dic_not_start,'buy_list':dic_buy,'limit_cls_recs':dic_limit_cls_recs,'maxdate_limit_class_rec':dic_limit_maxdate_rec})

    def send_to_start_lmt_page_data(self):
        data=request.json
        cus_id_name=data['cus_name']
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]
        conn=self.connect_mysql() 
        cursor=conn.cursor()
        try:
            #未开启限时私教课信息
            sql='''
                select DISTINCT card_id,card_start_time,prd,end_time,cls_qty from cards_table where card_id in 
                (select card_id from cardholder_card_table where cus_id=%s)
                and card_type=%s 
                and card_name=%s
                and card_start_time is NULL
                and end_time is NULL
            '''
            cursor.execute(sql,(cus_id,'limit_prd','私教'))
            not_start_limit_recs=cursor.fetchall()
            # print(not_start_limit_recs)

            #当天有效的限时私教课信息
            current_day=datetime.datetime.now().strftime('%Y-%m-%d')
            # sql='''
            #     select DISTINCT card_id,card_start_time,prd,end_time from cards_table where card_id in 
            #     (select card_id from cardholder_card_table where cus_id=%s)
            #     and card_type=%s 
            #     and card_name=%s
            #     and card_start_time<=%s
            #     and end_time>=%s
            # '''

            # 延长课程时，为insert，对于同一个card_id，查询opr_time为最大的那条记录    
            sql='''
                SELECT DISTINCT ct.card_id, ct.card_start_time, ct.prd, ct.end_time,ct.cls_qty
                FROM cards_table ct
                JOIN (
                    SELECT cht.card_id, MAX(ct.opr_time) AS max_opr_time
                    FROM cardholder_card_table cht
                    JOIN cards_table ct ON cht.card_id = ct.card_id
                    WHERE cht.cus_id = %s
                    GROUP BY cht.card_id
                    ) AS subq ON ct.card_id = subq.card_id AND ct.opr_time = subq.max_opr_time
                WHERE ct.card_id IN (
                    SELECT card_id
                    FROM cardholder_card_table
                    WHERE cus_id = %s
                    )
                AND ct.card_type = %s 
                AND ct.card_name = %s
                AND ct.card_start_time <= %s
                AND ct.end_time >= %s
            '''
            cursor.execute(sql,(cus_id,cus_id,'limit_prd','私教',current_day,current_day))
            current_valid_limit_rec=cursor.fetchall()


            return jsonify({'res':'ok',
                            'not_start_limit_recs':not_start_limit_recs,
                            'current_valid_limit_rec':current_valid_limit_rec})
        except Exception as e:
            print('get limit recs error')
            return jsonify({'res':'failed'})
        finally:
            cursor.close()
            conn.close()


    def deal_start_limit_page_db(self):
        print('get buy history via deal_start_limit_page_db()')
        data=request.json
        cus_id_name=data['cus_name']
        cls_tkn_time=datetime.datetime.strptime(data['cls_tkn_time'],'%Y-%m-%dT%H:%M')
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]
        conn=self.connect_mysql() 
        cursor=conn.cursor()

        # not_start_list
        sql='''
            select * from buy_rec_table 
                where card_id in 
                (SELECT card_id from not_start_lmt_table 
                WHERE cus_id='{cus_id}' and cus_name='{cus_name}')
            '''
        cursor.execute(sql)
        not_start_list=cursor.fetchall()
        if not_start_list:
            not_start_list=self.convert_mysql_data_to_string(not_start_list)
            # buy_rec_cols=['id','cus_id','cus_name','buy_date','buy_code','buy_type','buy_num','buy_cls_days','pay','real_pay','cashier_name','income_type','comment']
            buy_rec_cols=['index','cus_id','cus_name','收款日期','购课编码','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']
            not_start_list=self.mysql_list_data_to_dic(data=not_start_list,mysql_cols=buy_rec_cols)  
        else:
            not_start_list={'0':{'index':'','cus_id':'','cus_name':'','收款日期':'','购课编码':'','购课类型':'','购课节数':'','购课时长（天）':'',
                                            '应收金额':'','实收金额':'','收款人':'','收入类别':'','备注':''}}    

        # buy_list
        sql=f"select * from buy_rec_table where cus_id='{cus_id}' and cus_name='{cus_name}'"
        cursor.execute(sql)
        buy_list=cursor.fetchall()
        if buy_list:
            buy_list=self.convert_mysql_data_to_string(buy_list)
            # buy_rec_cols=['id','cus_id','cus_name','buy_date','buy_code','buy_type','buy_num','buy_cls_days','pay','real_pay','cashier_name','income_type','comment']
            buy_list_cols=['index','cus_id','cus_name','收款日期','购课编码','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']
            buy_list=self.mysql_list_data_to_dic(data=buy_list,mysql_cols=buy_list_cols)  
        else:
            buy_list={'index':'','cus_id':'','cus_name':'','收款日期':'','购课编码':'','购课类型':'','购课节数':'','购课时长（天）':'',
                                            '应收金额':'','实收金额':'','收款人':'','收入类别':'','备注':''}      

        # limit_cls_recs
        sql=f"SELECT * FROM lmt_cls_rec_table where cus_id='{cus_id}' and cus_name='{cus_name}'"
        cursor.execute(sql)
        limit_cls_recs=cursor.fetchall()
        if limit_cls_recs:
            limit_cls_recs=self.convert_mysql_data_to_string(limit_cls_recs)
            # limit_cls_recs_cols=['id','cus_id','cus_name','buy_code','start_date','end_date']
            limit_cls_recs_cols=['id','cus_id','cus_name','购课编码','限时课程起始日','限时课程结束日']
            limit_cls_recs=self.mysql_list_data_to_dic(data=limit_cls_recs,mysql_cols=limit_cls_recs_cols)  
        else:
            limit_cls_recs={'id':'','cus_id':'','cus_name':'','购课编码':'','限时课程起始日':'','限时课程结束日':''}      


        # maxdate_limit_cls_rec ##old
        # sql=f"SELECT * FROM lmt_cls_rec_table WHERE cus_id='{cus_id}' and cus_name='{cus_name}' ORDER BY end_date desc limit 1"
        # cursor.execute(sql)
        # maxdate_limit_cls_rec=cursor.fetchall()
        # if maxdate_limit_cls_rec:
        #     maxdate_limit_cls_rec=self.convert_mysql_data_to_string(maxdate_limit_cls_rec)
        #     # maxdate_limit_cls_rec=['id','cus_id','cus_name','buy_code','start_date','end_date']
        #     maxdate_limit_cls_rec_cols=['id','cus_id','cus_name','购课编码','限时课程起始日','限时课程结束日']
        #     maxdate_limit_cls_rec=self.mysql_list_data_to_dic(data=maxdate_limit_cls_rec,mysql_cols=maxdate_limit_cls_rec_cols)  
        # else:
        #     maxdate_limit_cls_rec={'0':{'id':'','cus_id':'','cus_name':'','购课编码':'','限时课程起始日':'','限时课程结束日':''}}   

        # maxdate_limit_cls_rec限时私教课
        # sql='''
        # SELECT card_id,max(end_time) FROM cards_table 
        # where card_id in (select card_id from cardholder_card_table where cus_id=%s)
        # and card_type='limit_prd'
        # and card_name='私教' 
        # and end_time>=%s
        # group by card_id;
        # '''


        sql='''
        SELECT 
            filtered_a.card_id, max(cards_table.end_time) as end_time,
            sum(cards_table.cls_qty) - IFNULL(b_count.times, 0) as remain_qty
        FROM (
            SELECT DISTINCT card_id 
            FROM cards_table
            WHERE card_id IN (
                SELECT card_id
                FROM cardholder_card_table
                WHERE cus_id=%s
            )
            AND card_type=%s
            AND card_name=%s
            AND end_time>=%s
        ) AS filtered_a
        LEFT JOIN (
            SELECT card_id, COUNT(card_id) as times
            FROM cls_tkn_rec_table
            GROUP BY card_id
        ) AS b_count ON filtered_a.card_id = b_count.card_id 
        LEFT JOIN cards_table ON filtered_a.card_id = cards_table.card_id
        GROUP BY filtered_a.card_id
        '''
        
        try:
            cursor.execute(sql,(cus_id,'limit_prd','私教',cls_tkn_time))
            card_id_lmt_sj,maxdate_limit_class_sj,lmt_sj_cls_remain=cursor.fetchone()
        except:
             card_id_lmt_sj=maxdate_limit_class_sj=lmt_sj_cls_remain=None

        # maxdate_limit_cls_rec限时团课
        sql='''
        SELECT card_id,max(end_time) FROM cards_table 
        where card_id in (select card_id from cardholder_card_table where cus_id=%s)
        and card_type='limit_prd'
        and card_name='团课'
        and end_time>=%s
        group by card_id;
        '''
        cursor.execute(sql,(cus_id,cls_tkn_time))
        try:
            card_id_lmt_grp,maxdate_limit_class_grp=cursor.fetchone()
        except:
            card_id_lmt_grp=maxdate_limit_class_grp=None



        # normal card id 普通卡
        # 普通卡卡号
        # sql='''
        # SELECT DISTINCT card_id FROM cards_table 
        # where card_id in (select card_id from cardholder_card_table where cus_id=%s)
        # and card_type='long_prd'
        # and card_name='私教'
        # and end_time>=%s
        # '''

        sql='''
            SELECT 
                filtered_a.card_id, max(cards_table.end_time) as end_time,
                sum(cards_table.cls_qty) - IFNULL(b_count.times, 0) as remain_qty
            FROM (
                SELECT DISTINCT card_id 
                FROM cards_table
                WHERE card_id IN (
                    SELECT card_id
                    FROM cardholder_card_table
                    WHERE cus_id=%s
                )
                AND card_type=%s
                AND card_name=%s
                AND end_time>=%s
            ) AS filtered_a
            LEFT JOIN (
                SELECT card_id, COUNT(card_id) as times
                FROM cls_tkn_rec_table
                GROUP BY card_id
            ) AS b_count ON filtered_a.card_id = b_count.card_id 
            LEFT JOIN cards_table ON filtered_a.card_id = cards_table.card_id
            GROUP BY filtered_a.card_id
        '''
        cursor.execute(sql,(cus_id,'long_prd','私教',cls_tkn_time))
        cards_id_cgsj=cursor.fetchall()     

        return jsonify({'not_start_list':not_start_list,
                        'buy_list':buy_list,
                        'limit_cls_recs':limit_cls_recs,
                        'maxdate_limit_class_sj':maxdate_limit_class_sj,
                        'limit_prd_card_id_sj':card_id_lmt_sj,
                        'maxdate_limit_class_grp':maxdate_limit_class_grp,
                        'limit_prd_card_id_grp':card_id_lmt_grp,
                        'lmt_sj_cls_remain':lmt_sj_cls_remain,
                        'long_prd_sj_card_ids':cards_id_cgsj})
        

    def convert_mysql_data_to_string(self,data,method=1):        
        if method==1:
            converted_data = []
            for item in data:
                if isinstance(item, tuple) or isinstance(item, list):
                    converted_data.append(self.convert_mysql_data_to_string(item))
                elif isinstance(item, datetime.date):
                    converted_data.append(item.strftime('%Y-%m-%d'))
                elif isinstance(item, Decimal):
                    converted_data.append(float(item))
                else:
                    converted_data.append(item)
        elif method==2:
            # 用于既往购课次数、未收金额的统计,输入的是字典
            converted_data = {}
            for key,item in data.items():
                if isinstance(item, tuple) or isinstance(item, list) or isinstance(item, dict):
                    converted_data[key]=self.convert_mysql_data_to_string(item,method=2)
                elif isinstance(item, datetime.date):
                    converted_data[key]=item.strftime('%Y-%m-%d')
                elif isinstance(item, Decimal):
                    converted_data[key]=float(item)
                else:
                    converted_data[key]=item
        
        return converted_data

    def mysql_list_data_to_dic(self,data,mysql_cols):
        result={}
        for key,item in enumerate(data):
            result[key]=self.list_to_dict(item,column_list=mysql_cols)

        return result

    def list_to_dict(self,data,column_list):
        result={}
        for key,item in enumerate(column_list):
            result[item]=data[key]
        
        return result

    def read_template(self):
        df=pd.read_excel(os.path.join(self.config_lz['work_dir'],'01-会员管理','模板.xlsm'),sheet_name='辅助表')
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
            fn=os.path.join(self.config_lz['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
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
    
    def write_buy(self):
        wk_dir=self.config_lz['work_dir']
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

    def write_buy_db(self):
        try:
            wk_dir=self.config_lz['work_dir']
            dat=request.json
            dat['cus_id']=dat['客户编码及姓名'][:7].strip()
            dat['cus_name']=dat['客户编码及姓名'][7:].strip()
            
            del dat['客户编码及姓名']
            data_cols=['cus_id','cus_name','收款日期', '购课卡号','购课类型','购课节数', '购课时长（天）', '应收金额', '实收金额', '收款人', '收入类别', '备注','operatorId','operateTime']
            sorted_data={key: dat[key] for key in data_cols}
            values=tuple(sorted_data.values())

            conn=self.connect_mysql()
            cursor=conn.cursor()
            
            conn.begin()

            #写入购课表
            sql='''
                insert into buy_rec_table
                (cus_id,cus_name,buy_date,card_id,buy_type,buy_num,buy_cls_days,pay,real_pay,cashier_name,income_type,comment,operator_id,operate_time)
                values
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''
            cursor.execute(sql,values)
            
            #写入卡表
            sql='''
                insert into cards_table
                (card_id,card_type,card_name,cls_qty,card_start_time,prd,end_time,cmt,opr_id,opr_time)
                values
                (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            '''

            if dat['购课类型']=='限时私教课' or dat['购课类型']=='限时团课':
                s_time=None
                end_time=None
            elif dat['购课类型']=='常规私教课' or dat['购课类型']=='常规团课':
                s_time=dat['收款日期']
                e_time=datetime.datetime.strptime(s_time,'%Y-%m-%d')+datetime.timedelta(days=int(dat['购课时长（天）']))
                end_time=e_time.strftime('%Y-%m-%d')

            values=(dat['购课卡号'],self.config_lz['cls_type_config'][dat['购课类型']]['type'],
                        self.config_lz['cls_type_config'][dat['购课类型']]['name'],dat['购课节数'],
                        s_time,dat['购课时长（天）'],end_time,dat['备注'],dat['operatorId'],dat['operateTime'])
            cursor.execute(sql,values)

            #写入持卡人-卡表
            sql='''
                select card_id from cardholder_card_table
                where cus_id=%s
            '''
            cursor.execute(sql,dat['cus_id'])
            hasIds=cursor.fetchall()
            has_ids=[]
            for id in hasIds:
                if id[0] not in has_ids:
                    has_ids.append(id[0])
            print('hasIds:',has_ids)
            if dat['cus_id'] in has_ids:
                pass
            else:
                try:
                    values=(dat['cus_id'],dat['购课卡号'],datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'')
                    sql='''
                        insert into cardholder_card_table 
                        (cus_id,card_id,relation_time,cmt)
                        values
                        (%s,%s,%s,%s)
                    '''
                    cursor.execute(sql,values)
                except:
                    print('insert into cardholder_card_table error')
                    raise FERROR 


            conn.commit()

            
            cus_name=sorted_data['cus_name']
            res_txt=f'{cus_name} 增加一条购课记录'

            # if sorted_data['购课类型'] in ['限时私教课','限时团课']:
            #     try:
            #         sql=f'''
            #             insert into not_start_lmt_table
            #             (cus_id,cus_name,buy_code)
            #             values
            #             (%s,%s,%s)
            #         '''
            #         cursor.execute(sql,[sorted_data['cus_id'],sorted_data['cus_name'],sorted_data['客户购课编号']])
            #         res_txt=f'; {cus_name} 未开课的限时课程记录增加成功'
            #     except Exception as start_lmt_error:                    
            #         res_txt=f' \n{cus_name} 未开课限时课程记录未增加成功:\n可能是同一天购课导致的重复购课编码。'
            #         print(f'ERROR: {cus_name} 未开课限时课程记录未增加成功：{start_lmt_error}')
            #         raise FERROR('未开课限时课程记录未增加成功，可能是同一天购课导致的重复购课编码。')
            
            
            
            
            return jsonify({'res':'ok','msg':res_txt})
        except Exception as e:
            conn.rollback()
            print('wirte_buy_db() error, rollbacked!', e)
            return jsonify({'res':'failed','msg':res_txt})
        finally:
            cursor.close()
            conn.close()


    def cus_list(self):
        dic_li=[]
        for fn in os.listdir(os.path.join(self.config_lz['work_dir'],'01-会员管理','会员资料')):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                dic_li.append(fn.split('.')[0])
        return dic_li


    def ins_list(self):
        fn=os.path.join(self.config_lz['work_dir'],'03-教练管理','教练资料','教练信息.xlsx')
        df=pd.read_excel(fn,sheet_name='教练信息')
        ins_li=df['姓名'].tolist()

        return ins_li

    def  get_train_dic(self):
        fn=os.path.join(self.config_lz['work_dir'],'05-专业资料','训练项目.xlsx')
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


    def  get_train_list_db(self):
        # fn=os.path.join(self.config_lz['work_dir'],'05-专业资料','训练项目.xlsx')
        # df=pd.read_excel(fn,sheet_name='训练项目')
        # df.fillna('',inplace=True)
        print('from minghu database,get train item list.')
        conn=self.connect_mysql()
        cursor = conn.cursor()

        result={}

        # 获取收款人
        sql="select * from train_item_table"
        cursor.execute(sql)
        train_items = cursor.fetchall()
        cursor.close()
        conn.close()

        df=pd.DataFrame(train_items)
        df.columns=['形式','肌肉部位','动作大类','动作名称']
        # print(df)


        
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


    def  get_train_list(self):
        fn=os.path.join(self.config_lz['work_dir'],'05-专业资料','训练项目.xlsx')
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

    def get_template_info_db(self):
        # 创建一个游标对象
        print('from minghu database,from template_info.')
        conn=self.connect_mysql()
        cursor = conn.cursor()

        result={}

        # 获取收款人
        sql="select cashier_name from cashier_table"
        cursor.execute(sql)
        cashiers_res = cursor.fetchall()
        cashiers=[x[0] for x in cashiers_res]
        result['cashiers']=cashiers

        #获取课程类型
        sql="select cls_name from cls_type_table"
        cursor.execute(sql)
        cls_types_res=cursor.fetchall()

        cls_types=[x[0] for x in cls_types_res]
        result['cls_types']=cls_types


        #获取收入类型
        sql="select income_type from income_type_table"
        cursor.execute(sql)
        income_type_res=cursor.fetchall()

        income_types=[x[0] for x in income_type_res]
        result['income_types']=income_types

        # print(result)
        # 关闭游标和连接
        cursor.close()
        conn.close()

        return jsonify(result)

    #遍历会员资料生成名字
    def get_cus_list(self):
        dic_li=self.cus_list()
        return jsonify(dic_li)

    def get_cus_list_db(self):
        print('from minghu database, get cus list')
        conn=self.connect_mysql()
        cursor = conn.cursor()

        result={}

        # 获取收款人
        sql="select concat(cus_id,cus_name) from cus_info_table"
        cursor.execute(sql)
        cus_list_res = cursor.fetchall()
        cus_list=[x[0] for x in cus_list_res]
        # result['cus_list']=cus_list

        cursor.close()
        conn.close()

        return jsonify(cus_list)

        



    #获取教练信息
    def get_ins_list(self):
        ins_li=self.ins_list()
        return jsonify(ins_li)

    def get_ins_list_db(self):
         # 创建一个游标对象
        print('from minghu database,get ins list.')
        conn=self.connect_mysql()
        cursor = conn.cursor()

        result={}

        # 获取教练姓名，不包括管理员角色
        sql="select ins_id,ins_name from ins_table where role='ins';"
        cursor.execute(sql)
        ins_res = cursor.fetchall()
        ins_list=[x[0]+x[1] for x in ins_res]
        result['ins_list']=ins_list

        cursor.close()
        conn.close()

        return jsonify(result)

    # 定义前端页面路由
   
    def cus_infos(self):
        return render_template('cus_infos.html')



    def input_buy(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id')
        return render_template('input_buy.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def new_cus(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id') 
        return render_template('new_cus.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def cus_cls_input(self):
        session_name=session.get('user')
        session_role=session.get('role')
        session_id=session.get('ins_id') 
        return render_template('cus_cls_input.html',session_ins_name=session_name,session_role=session_role,session_ins_id=session_id)

    def success(self):
        return render_template('success.html')


    def get_cus_info(self):
        cus_name = request.json.get('selected_name')
        work_dir=self.wecom_dir()
        fn=os.path.join(work_dir,cus_name+'.xlsm')
        p=cus_data.CusData()
        res=p.cus_cls_rec_toweb(fn=fn,cls_types=self.config_lz['all_cls_types'],not_lmt_types=self.config_lz['not_lmt_cls_types'])
        res.fillna(0)
        data=res.iloc[0].to_dict()
        # print('get_cus_info() ',data)
        return jsonify(data)

    def get_cus_info_db(self):
        cus_id_name=request.json.get('selected_name')
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]
        print('from minghu database, get cus list')
        conn=self.connect_mysql()
        cursor = conn.cursor()

        result={}

        result['会员编码及姓名']=cus_id_name

        #限时课程不能转让，用cus_id查询上课的节数
        #常规课程可以转让或共享，用card_id查询上课的节数

        # 限时课程到期日
        # sql=f"select max(end_date) from lmt_cls_rec_table WHERE cus_name={cus_name} and cus_id={cus_id}"
        # sql=f"select end_date from lmt_cls_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' order by end_date desc limit 1 "
        # cursor.execute(sql)
        # maxdate_limit_cls_get = cursor.fetchall()
        # if maxdate_limit_cls_get:
        #     max_date=maxdate_limit_cls_get[0][0]
        #     max_date=max_date.strftime('%Y-%m-%d')
        # else:
        #     max_date='-'
        
        # sql=f"select end_date from lmt_cls_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' order by end_date desc limit 1 "
        # cursor.execute(sql)

        sql=f'''
        select card_id,max(end_time) as max_lmt_time from cards_table 
        where card_id in (select card_id from cardholder_card_table where cus_id=%s) and card_type='limit_prd' GROUP by card_id

        '''
        cursor.execute(sql,cus_id)
        try:
            max_date=cursor.fetchone()[1]
            print('max_date',max_date)  
        except:
            max_date='' 

        # 限时课程是否有效
        if max_date:
            if datetime.datetime.now().date()<=max_date:
                result['限时课程是否有效']='是'
            else:
                result['限时课程是否有效']='否'
        else:
            result['限时课程是否有效']='否'

        if not max_date:
            max_date='-'
        result['限时课程到期日']=max_date


        #总消费金额
        sql=f"select sum(real_pay) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        total_pay=cursor.fetchall()
        total_pay=total_pay[0][0]
        if not total_pay:
            total_pay=0

        result['总消费金额']=float(total_pay)

        #平均每单消费金额
        sql=f"select sum(real_pay)/count(real_pay) as avr_pay from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        avr_pay=cursor.fetchall()
        avr_pay=avr_pay[0][0]
        if not avr_pay:
            avr_pay=0

        result['平均每单消费金额']=float(avr_pay)

        #最后一次购课日期
        sql=f"select max(buy_date) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        latest_buy_date=cursor.fetchall()
        latest_buy_date=latest_buy_date[0][0]
        if latest_buy_date:
            latest_buy_date=latest_buy_date.strftime('%Y-%m-%d')
        else:
            latest_buy_date='-'

        result['最后一次购课日期']=latest_buy_date

        #开始上课日期
        sql=f"select min(cls_datetime) from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        first_cls_tkn_date=cursor.fetchall()
        first_cls_tkn_date=first_cls_tkn_date[0][0]
        if first_cls_tkn_date:
            first_cls_tkn_date=first_cls_tkn_date.strftime('%Y-%m-%d')
        else:
            first_cls_tkn_date='-'

        result['开始上课日期']=first_cls_tkn_date


        #最后一次上课日期
        sql=f"select max(cls_datetime) from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        latest_cls_tkn_date=cursor.fetchall()
        latest_cls_tkn_date=latest_cls_tkn_date[0][0]
        if latest_cls_tkn_date:
            latest_cls_tkn_date=latest_cls_tkn_date.strftime('%Y-%m-%d')
        else:
            latest_cls_tkn_date='-'

        result['最后一次上课日期']=latest_cls_tkn_date

        #上课总天数
        sql=f"select datediff(max(cls_datetime),min(cls_datetime)) as total_cls_tkn_days from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        total_cls_tkn_days=cursor.fetchall()
        total_cls_tkn_days=total_cls_tkn_days[0][0]
        if not total_cls_tkn_days:
            total_cls_tkn_days=0

        result['上课总天数']=total_cls_tkn_days        
        
        #上课次数-限时私教课
        sql='''
            select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table 
            WHERE card_id in 
                (select card_id from cards_table where
                card_id in 
                    (select card_id from cardholder_card_table where cus_id=%s)
                and card_type=%s 
                and card_name=%s 
                and card_start_time<=%s
                and end_time>=%s
                )
        
        '''
        current_time=datetime.datetime.now().strftime('%Y-%m-%d')
        try:
            cursor.execute(sql,(cus_id,'limit_prd','私教',current_time,current_time))
            cls_tkn_count_lmt_sj=cursor.fetchall()
            cls_tkn_count_lmt_sj=cls_tkn_count_lmt_sj[0][0]
            if not cls_tkn_count_lmt_sj:
                cls_tkn_count_lmt_sj=0
        except Exception as e:
            print('get cls_tkn_count_lmt_sj:',e)
            cls_tkn_count_lmt_sj=0

        result['上课次数-限时私教课']=cls_tkn_count_lmt_sj

        #上课次数-常规私教课，
        # 通过cus_id查询到card_id，再查询card_id在cls_tkn_rec_table中的节数（次数）
        # sql=f"select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and cls_type='常规私教课'"
        sql='''
            select count(card_id) from cls_tkn_rec_table 
            where card_id in (select card_id from cards_table 
                                where card_id in (select card_id from cardholder_card_table where cus_id=%s) 
                                and card_type=%s
                                and card_name=%s)
        '''
        try:
            cursor.execute(sql,(cus_id,'long_prd','私教'))
            cls_tkn_count_cg_sj=cursor.fetchall()
            cls_tkn_count_cg_sj=cls_tkn_count_cg_sj[0][0]
            if not cls_tkn_count_cg_sj:
                cls_tkn_count_cg_sj=0
        except Exception as e:
            print('get cls_tkn_count_cg_sj:',e)
            cls_tkn_count_cg_sj=0

        result['上课次数-常规私教课']=cls_tkn_count_cg_sj

        #上课次数-限时团课
        # sql=f"select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and cls_type='限时团课'"
        
        sql='''
            select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table 
            WHERE card_id in 
                (select card_id from cards_table where
                card_id in 
                    (select card_id from cardholder_card_table where cus_id=%s)
                and card_type=%s 
                and card_name=%s 
                and card_start_time<=%s
                and end_time>=%s
                )
        
        '''
        try:
            current_time=datetime.datetime.now().strftime('%Y-%m-%d')
            cursor.execute(sql,(cus_id,'limit_prd','团课',current_time,current_time))
            cls_tkn_count_lmt_grp=cursor.fetchall()
            cls_tkn_count_lmt_grp=cls_tkn_count_lmt_grp[0][0]
            if not cls_tkn_count_lmt_grp:
                cls_tkn_count_lmt_grp=0
        except Exception as e:
            print('get cls_tkn_count_lmt_grp error:',e)
            cls_tkn_count_lmt_grp=0

        result['上课次数-限时团课']=cls_tkn_count_lmt_grp

        #上课次数-常规团课
        # sql=f"select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and cls_type='常规团课'"
        sql='''
            select count(card_id) from cls_tkn_rec_table 
            where card_id in (select card_id from cards_table 
                                where card_id in (select card_id from cardholder_card_table where cus_id=%s) 
                                and card_type=%s
                                and card_name=%s)
        '''
        try:
            cursor.execute(sql,(cus_id,'long_prd','团课'))
            # cursor.execute(sql)
            cls_tkn_count_cg_grp=cursor.fetchall()
            cls_tkn_count_cg_grp=cls_tkn_count_cg_grp[0][0]
            if not cls_tkn_count_cg_grp:
                cls_tkn_count_cg_grp=0
        except Exception as e:
            print('get cls_tkn_count_cg_grp error :',e)
            cls_tkn_count_cg_grp=0

        result['上课次数-常规团课']=cls_tkn_count_cg_grp

        #上课总次数
        # sql=f"select count(cls_datetime) as cls_tkn_count from cls_tkn_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
        # cursor.execute(sql)
        # cls_tkn_count=cursor.fetchall()
        # cls_tkn_count=cls_tkn_count[0][0]
        # if not cls_tkn_count:
        #     cls_tkn_count=0
        try:
            result['上课总次数']=result['上课次数-限时私教课']+result['上课次数-常规私教课']+ \
                                result['上课次数-限时团课']+result['上课次数-常规团课']
        except Exception as e:
            print('calculate total cls tkn qty error :',e)
            result['上课总次数']=0

        

        #上课频率
        if result['上课总次数']!=0:
            cls_frqcy=total_cls_tkn_days/result['上课总次数']
        else:
            cls_frqcy=0
        result['上课频率']=cls_frqcy

        ######################################################################
        #购课次数-限时私教课
        sql=f"select count(buy_date) as buy_count from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='限时私教课'"
        cursor.execute(sql)
        buy_count_lmt_sj=cursor.fetchall()
        buy_count_lmt_sj=buy_count_lmt_sj[0][0]
        if not buy_count_lmt_sj:
            buy_count_lmt_sj=0

        result['购课次数-限时私教课']=buy_count_lmt_sj

        #购课次数-常规私教课
        sql=f"select count(buy_date) as buy_count from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规私教课'"
        cursor.execute(sql)
        buy_count_cg_sj=cursor.fetchall()
        buy_count_cg_sj=buy_count_cg_sj[0][0]
        if not buy_count_cg_sj:
            buy_count_cg_sj=0

        result['购课次数-常规私教课']=buy_count_cg_sj

        #购课次数-限时团课
        sql=f"select count(buy_date) as buy_count from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='限时团课'"
        cursor.execute(sql)
        buy_count_lmt_grp=cursor.fetchall()
        buy_count_lmt_grp=buy_count_lmt_grp[0][0]
        if not buy_count_lmt_grp:
            buy_count_lmt_grp=0

        result['购课次数-限时团课']=buy_count_lmt_grp

        #购课次数-常规团课
        sql=f"select count(buy_date) as buy_count from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规团课'"
        cursor.execute(sql)
        buy_count_cg_grp=cursor.fetchall()
        buy_count_cg_grp=buy_count_cg_grp[0][0]
        if not buy_count_cg_grp:
            buy_count_cg_grp=0

        result['购课次数-常规团课']=buy_count_cg_grp

        ###########################################
        #购课节数-常规私教课
        # sql=f"select sum(buy_num) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规私教课'"
        sql='''
        select sum(cls_qty) from cards_table 
        WHERE card_id in  (select card_id from cardholder_card_table where cus_id=%s) 
        and card_name='私教' 
        and card_type='long_prd';

        '''
        cursor.execute(sql,cus_id)
        buy_num_cg_sj=cursor.fetchone()[0]
        # buy_num_cg_sj=buy_num_cg_sj[0]
        if not buy_num_cg_sj:
            buy_num_cg_sj=0

        result['购课节数-常规私教课']=float(buy_num_cg_sj)



        #购课节数-常规团课
        # sql=f"select sum(buy_num) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规团课'"
        sql='''
        select sum(cls_qty) from cards_table 
        WHERE card_id in  (select card_id from cardholder_card_table where cus_id=%s) 
        and card_name='团课' 
        and card_type='long_prd';

        '''
        cursor.execute(sql,cus_id)
        buy_num_cg_grp=cursor.fetchone()[0]
        if not buy_num_cg_grp:
            buy_num_cg_grp=0

        result['购课节数-常规团课']=float(buy_num_cg_grp)


        ###########################################
        #剩余节数-常规私教课
        sql=f"SELECT cls_tkn_adj_num_cg_sj from adjust_table where cus_name='{cus_name}' and cus_id='{cus_id}'"
        cursor.execute(sql)
        adj_num_cg_sj=cursor.fetchall()
        if adj_num_cg_sj:
            adj_num_cg_sj=adj_num_cg_sj[0][0]
        else:
            adj_num_cg_sj=0

        result['剩余节数-常规私教课']=result['购课节数-常规私教课']-result['上课次数-常规私教课']-adj_num_cg_sj


        #########################################
        #消费金额-限时私教课
        sql=f"select sum(real_pay) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='限时私教课'"
        cursor.execute(sql)
        total_pay_lmt_sj=cursor.fetchall()
        total_pay_lmt_sj=total_pay_lmt_sj[0][0]
        if not total_pay_lmt_sj:
            total_pay_lmt_sj=0

        result['消费金额-限时私教课']=float(total_pay_lmt_sj)

        #消费金额-常规私教课
        sql=f"select sum(real_pay) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规私教课'"
        cursor.execute(sql)
        total_pay_cg_sj=cursor.fetchall()
        total_pay_cg_sj=total_pay_cg_sj[0][0]
        if not total_pay_cg_sj:
            total_pay_cg_sj=0

        result['消费金额-常规私教课']=float(total_pay_cg_sj)

        #消费金额-限时团课
        sql=f"select sum(real_pay) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='限时团课'"
        cursor.execute(sql)
        total_pay_lmt_grp=cursor.fetchall()
        total_pay_lmt_grp=total_pay_lmt_grp[0][0]
        if not total_pay_lmt_grp:
            total_pay_lmt_grp=0

        result['消费金额-限时团课']=float(total_pay_lmt_grp)


        #消费金额-常规团课
        sql=f"select sum(real_pay) from buy_rec_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' and buy_type='常规团课'"
        cursor.execute(sql)
        total_pay_cg_grp=cursor.fetchall()
        total_pay_cg_grp=total_pay_cg_grp[0][0]
        if not total_pay_cg_grp:
            total_pay_cg_grp=0


        #最近一次体测结果
        body_history=self.get_lst_body_history_db(cus_id_name,jsonify='no')

        result.update(body_history)

        cursor.close()
        conn.close()

        return jsonify(result)
    
    def get_cus_cards_by_type(self):
        print('get card ids by type....')
        data=request.json
        cus_id=data['cus_id_name'][:7]
        cus_name=data['cus_id_name'][7:]
        buy_type=data['buy_type']
        buy_date=data['buy_date']

        if buy_type=='':
            buy_type='限时私教课'
        
        
        cls_type=self.config_lz['cls_type_config'][buy_type]['type']        
        cls_name=self.config_lz['cls_type_config'][buy_type]['name']

        print(cls_type,cls_name)

        conn=self.connect_mysql()
        cursor=conn.cursor()
        
        try:
            #长期常规课程
            if cls_type=='long_prd':                
                sql='''
                    select distinct card_id from cards_table 
                    where card_id in (select card_id from cardholder_card_table where cus_id=%s)
                    and card_type=%s
                    and card_name=%s
                    and card_start_time<=%s 
                    and end_time>=%s
                '''
                cursor.execute(sql,(cus_id,cls_type,cls_name,buy_date,buy_date))
                res=cursor.fetchall()
                if res:
                    pass
                #如果不存在卡号，新生成卡号
                else:
                    print('get no cus card id,generate new long_prd card id')
                    cr_time=datetime.datetime.now().strftime('%Y%m%d%H%M%S')
                    new_card_id=self.config_lz['cus_id_prefix']+'C'+cr_time
                    res=[new_card_id]
            #限时课程
            elif cls_type=='limit_prd':
                # sql='''
                #     select card_id from cards_table 
                #     where card_id in (select card_id from cardholder_card_table where cus_id=%s)
                #     and card_type=%s
                #     and card_name=%s
                #     and card_start_time<=%s 
                #     and end_time>=%s
                # '''
                # cursor.execute(sql,(cus_id,cls_type,cls_name,buy_date,buy_date))
                # res=cursor.fetchall()
                # if res:
                #     pass
                # #如果不存在卡号，新生成卡号
                # else:
                #     print('get no cus card id,generate new long_prd card id')
                #     cr_time=datetime.datetime.now().strftime('%Y%m%d%H%M%S')
                #     new_card_id=self.config_lz['cus_id_prefix']+'L'+cr_time
                #     res=[new_card_id]

                cr_time=datetime.datetime.now().strftime('%Y%m%d%H%M%S')    
                new_card_id=self.config_lz['cus_id_prefix']+'L'+cr_time
                res=[new_card_id]
                

            return jsonify({'res':'ok','card_ids':res})

        except Exception as e:
            print('get cus card ids failed: ',e)
            return jsonify({'res':'failed'})
        finally:
            cursor.close()
            conn.close()


        

    def get_lst_body_history_db(self,cus_id_name,jsonify='yes'):
        if not cus_id_name:
            cus_id_name=request.text.decode('utf-8')
        cus_id,cus_name=cus_id_name[:7],cus_id_name[7:]
        conn=self.connect_mysql()
        cursor = conn.cursor()
        #获取最近一次的体测数据
        sql=f"select * from body_msr_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}' order by msr_date DESC limit 1"
        cursor.execute(sql)
        lst_body_history=cursor.fetchall()

        result={}
        if lst_body_history:
            for key,item in enumerate(['id','cus_id','cus_name','msr_date','ht','wt','bfr','chest','l_arm','r_arm','waist','hip','l_leg','r_leg','l_calf','r_calf','heart','balance','power','flex','core']):
                try:
                    result[item]=float(lst_body_history[0][key])
                except:
                    result[item]=lst_body_history[0][key]
            #修改键值名返回前端
            result['lst_msr_date']=result.pop('msr_date')
            result['lst_msr_date']=result['lst_msr_date'].strftime('%Y-%m-%d')
            result.pop('id')
            result.pop('cus_id')
            result.pop('cus_name')

            #获取体检日期列表
            sql=f"select msr_date from body_msr_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
            cursor.execute(sql)
            msr_dates=cursor.fetchall()
            msr_dates=[x[0].strftime('%Y-%m-%d') for x in msr_dates]
            msr_dates_txt='\n'.join(msr_dates)
            result['msr_dates']=msr_dates_txt

            #获取体测次数
            sql=f"select count(msr_date) from body_msr_table WHERE cus_name='{cus_name}' and cus_id='{cus_id}'"
            cursor.execute(sql)
            msr_num=cursor.fetchall()
            msr_num=msr_num[0][0]
            result['msr_num']=msr_num

        else:
            result={'lst_msr_date': '-', 'msr_num': '-', 
            'msr_dates': '-', 'bfr': '-', 'ht': '-', 'wt': '-', 'waist': '-', 
            'chest': '-', 'l_arm': '-', 'r_arm': '-', 'hip': '-', 'l_leg': '-', 
            'r_leg': '-', 'l_calf': '-', 'r_calf': '-', 
            'heart': '-', 'balance': '-', 'power': '-', 
            'flex': '-', 'core': '-'}

        if jsonify=='yes':
            cursor.close()
            conn.close()
            return jsonify({'lst_msr_date':result})
        else:
            return result
       



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

    def _get_max_code(self):        
        cus_li = self.cus_list()
        cus_num = [int(x[2:5]) for x in cus_li]
        max_num = max(cus_num)
        new_num = max_num + 1
        return str(new_num).zfill(3)


    def check_new(self):
        # dat=request.data.decode('utf-8')        
        # with self.app_lock:
        txt_num=self._get_max_code()            
        print('新客户自动编号：',txt_num)           
            
        return txt_num

    
    def generate_new(self):
        with self.app_lock:
            txt_num=self.check_new()
            print('generating new',txt_num)
            try:
                fn_in=request.data
                fn='MH'+txt_num+fn_in.decode('utf-8')
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

                return fn
            except Exception as e:
                self.release_event.set()
                return e
        
    def generate_new_db(self):
        with self.app_lock:            
            conn=self.connect_mysql()
            cursor=conn.cursor()
            sql=f"SELECT max(cast(substring(cus_id,3) as UNSIGNED)) FROM cus_info_table;"
            cursor.execute(sql)
            max_id=cursor.fetchall()[0][0]
            
            txt_num=str(max_id+1).zfill(5)
            try:
                dataRequest=request.json
     
                fn=self.cus_id_prefix+txt_num+dataRequest['cusName']
                # fn='MH00220王测试|王测试|女|199008|小红书|pc'
                data={}  
                data['trial_cus_name']=dataRequest['trialCusName']
                data['sex']=dataRequest['gender']
                data['mobile_phone']=dataRequest['mobilePhone']
                data['birthday']=dataRequest['birthDay']
                data['source']=dataRequest['cusSource']
                data['dvc']=dataRequest['dvc']
                data['cus_id']=self.cus_id_prefix+txt_num
                data['cus_name']=dataRequest['cusName']
                data['nick_name']=data['cus_name'] if len(data['cus_name'])<2 else data['cus_name'][1:]
                data['birthday_type']='ymd'
                ins_ids=dataRequest['insIds']
                # data['birthday_type']='ym'
                # trial_cus_name=data['trial_cus_name']
                cus_id_name=data['cus_id']+data['cus_name']

                data_col=['cus_id','cus_name','nick_name','sex','mobile_phone','birthday','birthday_type','source']
                sorted_data={key: data[key].strip() for key in data_col}
                values=tuple(sorted_data.values())
                today=datetime.datetime.now().strftime('%Y-%m-%d')

                #生成主管教练的姓名及教练ID
                ins_name_arr=[]
                ins_id_arr=[]
                with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
                    try:
                        for ins_id in ins_ids.split(','):
                            sql='''
                                select ins_id,ins_name from ins_table 
                                where ins_id=%s
                            '''
                            cursor.execute(sql,(ins_id,))
                            fetch_ins=cursor.fetchone()
                            ins_name_arr.append(fetch_ins['ins_name'])
                            ins_id_arr.append(fetch_ins['ins_id'])
                        ins_names_str=','.join(ins_name_arr)
                        ins_ids_str=','.join(ins_id_arr)                        
                    except Exception as e:
                        print(e) 
                        return jsonify({'res':'failed','error':f'error:{e}'})
                # except Exception as e:
                #     print(e)
                #     return jsonify({'res':'failed','error':f'error:{e}'})
                # finally:
                #     cursor.close()
                #     conn.close()
                
                
                try:
                    cursor=conn.cursor()
                    #新增会员
                    sql=f'''
                            insert into  cus_info_table (cus_id,cus_name,nick_name,sex,mobile_phone,birthday,birthday_type,source) 
                            values
                            (%s,%s,%s,%s,%s,%s,%s,%s)
                        '''
                    cursor.execute(sql,values)


                    #写入教练主管会员表
                    sql=f'''
                            insert into ins_control_cus_table (cus_id,cus_name,ins_ids,ins_names) 
                            values
                            (%s,%s,%s,%s)
                        '''
                    cursor.execute(sql,(data['cus_id'],data['cus_name'],ins_ids_str,ins_names_str))

                    #在体验课表写入体验课出单会员信息
                    if data['trial_cus_name']:
                        sql=f'''
                            update trial_cls_table 
                            set deal_yn='是',deal_date=%s,formal_cus_id_name=%s                        
                            where trial_cus_name=%s order by 'id' Desc LIMIT 1
                            '''
                        cursor.execute(sql,[today,cus_id_name,data['trial_cus_name']])

                        sql='''
                            insert into ins_control_cus_table 
                            
                        '''

                    conn.commit()  
                except Exception as e:
                    print(e)
                    return jsonify({'res':'failed','error':f'error:{e}'})
                

            except Exception as e:
                print('write into new cus error:',e)
                return jsonify({'res':'failed','error':f'error:{e}'})
            finally:
                cursor.close()
                conn.close()                
            
        
            return jsonify({'res':'ok','cus_name':f'{cus_id_name}'})

    def write_deal_cus_name_to_trial_table(self,formal_cus_name,trial_cus_name):
        try:
            fn=os.path.join(self.config_lz['work_dir'],'03-教练管理','体验课上课记录表.xlsx')
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
        print('welcome')
        return jsonify({'res':'关于我们页面','res2':'这个是welcome页面KK'})

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

class FERROR(Exception):
    pass

if __name__ == '__main__':
    app = MinghuService(__name__)
    app.secret_key='minghu8888 '
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
