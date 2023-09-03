import os
import sys
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
import re
import pymysql
from pymysql.constants import CLIENT
from sshtunnel import SSHTunnelForwarder

class Excel:
    def __init__(self,ssh_addr,ssh_port,ssh_user_name,ssh_pkey_pathname,bind_db_addr,bind_db_port,db_username,db_pwd,db_name) :
        self.ssh_addr=ssh_addr
        self.ssh_port=ssh_port
        self.ssh_user_name=ssh_user_name
        self.ssh_pkey_pathname=ssh_pkey_pathname
        self.bind_db_addr=bind_db_addr
        self.bind_db_port=bind_db_port
        self.db_username=db_username
        self.db_pwd=db_pwd
        self.db_name=db_name

    def connect_mysql_ssh(self):
        # 配置 SSH 连接
        server = SSHTunnelForwarder(
            (self.ssh_addr, self.ssh_port),  # B机器的配置
            # ssh_password='ssh_password',
            ssh_username= self.ssh_user_name,
            ssh_pkey=self.ssh_pkey_pathname,  # 私钥路径
            # ssh_private_key_password="m5!nf",  # 跳转机的用户密码
            remote_bind_address=(self.bind_db_addr, self.bind_db_port)
        )
        server.start()

        conn = pymysql.connect(host=self.bind_db_addr,  # 此处必须是是127.0.0.1
                                    port=server.local_bind_port,
                                    user=self.db_username,
                                    passwd=self.db_pwd,
                                    db=self.db_name)



        # server.close()#最后一定要server.close()，不然程序不会结束
        
            

        return conn
    
    def test_conn_db(self):
        conn=self.connect_mysql_ssh()
        try:            
            # with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
            #     print('connect successfully')
            #     sql='select * from basic_info_table order by "cus_id" desc limit 5;'
            #     cursor.execute(sql)
            #     res=cursor.fetchall()
            #     print(res)
            cursor=conn.cursor()
            sql='select * from basic_info_table order by "cus_id" desc limit 5;'
            cursor.execute(sql)
            res=cursor.fetchall()
            print(res)
        finally:
            conn.close()



    #单个表基本信息合并
    def one_basic_info(self,id='MH00220',fn='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH228枝枝.xlsm'):
        df=pd.read_excel(fn,sheet_name='基本情况')
        if len(str(df['出生年月'].tolist()[0]))==4:
            df['出生年月']=df['出生年月'].apply(lambda x: str(x)+'-01-01')
            birthday_type='y'
        elif len(str(df['出生年月'].tolist()[0]))==6:
            df['出生年月']=df['出生年月'].apply(lambda x: str(x)+'-01')
            birthday_type='ym'
        elif len(str(df['出生年月'].tolist()[0]))==8:
            df['出生年月']=df['出生年月'].apply(lambda x: str(x))
            birthday_type='ymd'
        else:
            birthday_type=None
        
        df['cus_id']=id
        df['birthday_type']=birthday_type
        df['id']='NULL'

        df_res=df[['id','cus_id','姓名','昵称','性别','出生年月','birthday_type','客户来源']]


        return df_res

    #批量基本信息合并
    def batch_basic_info(self,dir_start_id="MH228",db_start_id="MH00220",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        db_id_num=int(db_start_id[2:])
        dfs=[]
        for fn in os.listdir(dir):
            
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                if int(fn[2:5])>=int(dir_start_id[2:]):
                    print(fn)
                    id='MH'+str(db_id_num).zfill(5)
                    df=self.one_basic_info(id=id,fn=os.path.join(dir,fn))
                    dfs.append(df)
                    db_id_num+=1
        dfs_res=pd.concat(dfs)    

        return dfs_res
        
     #单个上课记录
    def one_cls_tkn(self,fn='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH228枝枝.xlsm'):
        df=pd.read_excel(fn,sheet_name='上课记录')
        cus_name=fn.split('\\')[-1].split('.')[0][5:]
        try:
            conn=self.connect_mysql_ssh()
            with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
                # print('connect successfully')
                sql=f'select cus_id from basic_info_table where cus_name="{cus_name}" limit 1;'
                cursor.execute(sql)
                res=cursor.fetchall()
                cus_id=res[0]['cus_id']
        except:
            cus_id=''

        finally:
            df.dropna(how='all',inplace=True)
            conn.close()
        df['cus_id']=cus_id
        df['cus_name']=cus_name
        df['id']='NULL'

        
        try:
            df['cls_datetime']=df['日期'] + pd.to_timedelta(df['时间'].astype(str))
        except Exception as e:
            print(e)
            df['cls_datetime']=''
        df_res=df[['id','cus_id','cus_name','cls_datetime','时长（小时）','课程类型','教练','备注']]
        
        return df_res
    
    #批量处理上课记录
    def batch_cls_tkn(self,dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        dfs=[]
        for fn in os.listdir(dir):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                if int(fn[2:5])>=int(dir_start_id[2:]):
                    print(fn)
                    df=self.one_cls_tkn(fn=os.path.join(dir,fn))
                    dfs.append(df)
        df_res=pd.concat(dfs)
    
        return df_res
    
    #单个购课记录
    def one_buy_rec(self,fn='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH228枝枝.xlsm'):
        df=pd.read_excel(fn,sheet_name='购课表')
        cus_name=fn.split('\\')[-1].split('.')[0][5:]
        try:
            conn=self.connect_mysql_ssh()
            with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
                # print('connect successfully')
                sql=f'select cus_id from basic_info_table where cus_name="{cus_name}" limit 1;'
                cursor.execute(sql)
                res=cursor.fetchall()
                cus_id=res[0]['cus_id']
        except:
            cus_id=''

        finally:
            df.dropna(how='all',inplace=True)
            conn.close()
        df['cus_id']=cus_id
        df['cus_name']=cus_name
        df['buy_code']=df['cus_id']+df['cus_name']+df['购课编码'].str[-8:]
        df['id']='NULL'

        # 创建一个条件，如果购课类型为'常规私教课'，则设置购课时长为99999，否则保持不变
        condition = df['购课类型'] == '常规私教课'
        df.loc[condition, '购课时长（天）'] = 99999

        df_res=df[['id','cus_id','cus_name','收款日期','buy_code','购课类型','购课节数','购课时长（天）','应收金额','实收金额','收款人','收入类别','备注']]

        
        return df_res

    #批量处理购课记录
    def batch_buy_rec(self,dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        dfs=[]
        for fn in os.listdir(dir):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                if int(fn[2:5])>=int(dir_start_id[2:]):
                    print(fn)
                    df=self.one_buy_rec(fn=os.path.join(dir,fn))
                    dfs.append(df)
        df_res=pd.concat(dfs)
    
        return df_res


    #单个限时课程记录
    def one_lmt_rec(self,fn='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料\\MH237甘玉莲.xlsm'):
        df=pd.read_excel(fn,sheet_name='限时课程记录')
        cus_name=fn.split('\\')[-1].split('.')[0][5:]
        try:
            conn=self.connect_mysql_ssh()
            with conn.cursor(cursor=pymysql.cursors.DictCursor) as cursor:
                # print('connect successfully')
                sql=f'select cus_id from basic_info_table where cus_name="{cus_name}" limit 1;'
                cursor.execute(sql)
                res=cursor.fetchall()
                cus_id=res[0]['cus_id']
        except:
            cus_id=''

        finally:
            df.dropna(how='all',inplace=True)
            conn.close()
        df['cus_id']=cus_id
        df['cus_name']=cus_name
        df['buy_code']=df['cus_id']+df['cus_name']+df['购课编码'].str[-8:]
        df['id']='NULL'
 

        df_res=df[['id','cus_id','cus_name','buy_code','限时课程起始日','限时课程结束日']]

        
        return df_res


    #批量处理限时课程开课记录
    def batch_lmt_rec(self,dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        dfs=[]
        for fn in os.listdir(dir):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                if int(fn[2:5])>=int(dir_start_id[2:]):
                    print(fn)
                    df=self.one_lmt_rec(fn=os.path.join(dir,fn))
                    dfs.append(df)
        df_res=pd.concat(dfs)
    
        return df_res

if __name__=='__main__':
    p=Excel(ssh_addr='3.tcp.cpolar.top',
            ssh_port=12910,
            ssh_user_name='jacklocal',
            ssh_pkey_pathname='c:\\users\\admin\\.ssh\\jack_mi_id_rsa',
            bind_db_addr='127.0.0.1',
            bind_db_port=3306,
            db_username='minghu',
            db_pwd='minghu88',
            db_name='minghu_trial')

    # res=p.batch_basic_info(dir_start_id="MH228",db_start_id="MH00220",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料')
    # res.to_excel('e:\\temp\\minghu\\from_MH228枝枝_basic_info_merge.xlsx',index=False)
    # p.test_conn_db()
    # res=p.batch_cls_tkn(dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料')
    # res.to_excel('e:\\temp\\minghu\\from_MH228枝枝_cls_tkn_merge.xlsx',index=False)
    # res=p.batch_buy_rec(dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料')
    # res.to_excel('e:\\temp\\minghu\\from_MH228枝枝_buy_rec_merge.xlsx',index=False)
    # res=p.batch_lmt_rec(dir_start_id="MH228",dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料')
    # res.to_excel('e:\\temp\\minghu\\from_MH228枝枝_lmt_rec_merge.xlsx',index=False)
    # print(res)