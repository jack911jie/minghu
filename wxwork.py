import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import get_data
import days_cal
import requests
import json
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class WeComRobot:
    def __init__(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室',dl_taken_fn='e:\\temp\\minghu\\教练工作日志.xlsx'):
        self.work_dir=work_dir
        self.dl_taken_fn=dl_taken_fn
        self.url='https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=cd85d3e6-c252-4c02-a631-984c813c6f67'    

    
    
    def send_data(self,cus_name='MH016徐颖丽',crs_type='常规私教课',crs_date='20220527',crs_time='1000-1100',ins='MHINS001陆伟杰'):
        txt_ready=get_data.ReadCourses(work_dir=self.work_dir,dl_taken_fn=self.dl_taken_fn)        
        txt_to_send=txt_ready.exp_txt(cus_name=cus_name,crs_type=crs_type,crs_date=crs_date,crs_time=crs_time,ins=ins)
        df_ins=txt_ready.ins_info(ins=ins)
        ins_name=ins[8:]
        ins_tel=str(df_ins['电话'].tolist()[0])
        ins_inform_txt= ins_name+'教练，'+'您的会员约课信息如下：'

        self.send(txt_to_send=txt_to_send,ins_inform_txt=ins_inform_txt,ins_name=ins_name,ins_tel=ins_tel)

    def send(self,txt_to_send,ins_inform_txt,ins_name,ins_tel):
        print(txt_to_send,'\n',ins_name,'\n',ins_tel,'\n')            
        
        data={
            "msgtype": "text",
            "text": {
                "content": txt_to_send,
                # "mentioned_list":["wangqing","@all"],
                # "mentioned_mobile_list":["15678892330"]
                }
            }

        ins_data={
            "msgtype": "text",
            "text": {
                "content": ins_inform_txt,
                # "mentioned_list":["wangqing","@all"],
                "mentioned_mobile_list":[ins_tel]
                }
            }
        
        # requests.post(self.url,json=ins_data).json()
        # requests.post(self.url,json=data).json()

        print('发送完成')

    def group_send(self,y_m='202206',crs_type='常规私教课'):
        txt_ready=get_data.ReadCourses(work_dir=self.work_dir)        
        txt_to_send=txt_ready.group_exp_txt(y_m=y_m,crs_type=crs_type)     

        for ins in txt_to_send:
            print('\n正在发送给 '+ins+' ')
            df_ins=txt_ready.ins_info(ins=ins)
            ins_tel=str(df_ins[df_ins['姓名']==ins]['电话'].tolist()[0])
            ins_data={
                        "msgtype": "text",
                        "text": {
                            "content": ins+'教练，'+'您的会员约课信息如下：',
                            # "mentioned_list":["wangqing","@all"],
                            "mentioned_mobile_list":[ins_tel]
                            }
                        }
            requests.post(self.url,json=ins_data).json()
            # print(ins_data)
            for num in txt_to_send[ins]:
                print('第'+str(num+1)+'条（共'+str(len(txt_to_send[ins]))+'条）……',end='')
                data={
                "msgtype": "text",
                "text": {
                    "content": txt_to_send[ins][num],
                    # "mentioned_list":["wangqing","@all"],
                    # "mentioned_mobile_list":["15678892330"]
                    }
                }

                requests.post(self.url,json=data).json()
                # print(data)
                print('完成')
            split_line={
                        "msgtype": "text",
                        "text": {
                            "content": '------分隔线------',
                            # "mentioned_list":["wangqing","@all"],
                            # "mentioned_mobile_list":[ins_tel]
                            }
                        }
            requests.post(self.url,json=split_line).json()
            # print(split_line)


        print('\n全部发送完成')


class Notification:
    def __init__(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室',dl_taken_fn='e:\\temp\\minghu\\教练工作日志.xlsx'):
        self.work_dir=work_dir
        self.dl_taken_fn=dl_taken_fn
        #开始启用系统计算日期，参看“20220430限时课程会员剩余课程节数.xlsx”，故从2022-5-1开始计算。
        self.begin_time='20220501'
        self.url='https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=cd85d3e6-c252-4c02-a631-984c813c6f67'

    def get_bodydata(self,cus_name):
        df_body=get_data.CusInfo().get_cus_body_data(os.path.join(self.work_dir,'01-会员管理','会员资料',cus_name+'.xlsx'))
        df_basic=get_data.CusInfo().get_cus_basic_data(os.path.join(self.work_dir,'01-会员管理','会员资料',cus_name+'.xlsx'))
        
        sex=df_basic['性别'].tolist()[0]
        age=df_basic['出生年月'].tolist()[0]

        if len(str(age))==4:
            age=str(age)+'0101'
        elif len(str(age))==6:
            age=str(age)+'01'
        elif len(str(age))==8:
            pass
        else:
            print('出生年月录入错误')
            exit(0)

        age=days_cal.calculate_age(birth_s=age)
        # print(df_body)
        df_body['体脂率']=df_body.apply(lambda x: get_data.cals().bfr(age=age,sex=sex,ht=x['身高'],wt=x['体重'],waist=x['腰围'],adj_bfr='no',adj_src='prg',gui='',formula=1),axis=1)        
        df_body=df_body.fillna(0)

        return df_body


    def data_prep(self,cus_name='MH010苏云',crs_types=['常规私教课','初级团课'],nowtime='20220804'):
        
        # cus_list=get_data.CusInfo().get_cus_list(work_dir=os.path.join(self.work_dir,'01-会员管理','会员资料'))
        # target_dir=os.path.join(self.work_dir,'01-会员管理')
        df_body=self.get_bodydata(cus_name=cus_name)

        # #计算给定日期为止上了几节
        # this_cus_crs=get_data.ReadCourses(work_dir=self.work_dir,dl_taken_fn=self.dl_taken_fn).cus_taken(cus_name=cus_name,crs_types=crs_types,start_time=self.begin_time,nowtime=nowtime)
        # this_total_crs=this_cus_crs['上课次数'].sum()
        # # this_total_crs=11
        # print('this_total_crs:',this_total_crs)

        #计算提醒的节数
        rng=[]
        for x in range(1,11):
            rng.extend([x*10-1,x*10,x*10+1])
        # rng=list(rng.sort)
        rng=sorted(rng)

        #计算从上次测量的时间到本次给定时间，期间上了几节
        latest_msr_time=df_body['时间'].max()
        if latest_msr_time<datetime.strptime(self.begin_time,'%Y%m%d'):
            latest_msr_time=self.begin_time
        else:
            latest_msr_time=latest_msr_time.strftime('%Y%m%d')
        cus_crs_last=get_data.ReadCourses(work_dir=self.work_dir,
                    dl_taken_fn=self.dl_taken_fn).cus_taken(cus_name=cus_name,crs_types=crs_types,start_time=latest_msr_time,nowtime=nowtime)
        
        if cus_crs_last.empty:
            last_total_crs=0
        else:
            last_total_crs=cus_crs_last['上课次数'].sum()
        # print('last_total_crs:',last_total_crs)

        if last_total_crs in rng:
            return [last_total_crs,latest_msr_time]

        else:

            # print('不需提醒')
            return [0,0]
        

    def send_body_note(self,cus_name='MH010苏云',crs_types=['常规私教课','初级团课'],nowtime='20220804',ins_name='MHINS001陆伟杰'):
        crs_after_last_msr,last_msr_time=self.data_prep(cus_name=cus_name,crs_types=crs_types,nowtime=nowtime)
        # print(crs_after_last_msr)
        if crs_after_last_msr==0:
            print('未到体测时间')
            return 0
        else:
            txt_to_send='● '+cus_name+' 上次体测日期为：'+last_msr_time[:4]+'-'+last_msr_time[4:6]+'-'+last_msr_time[6:]+'\n● 上次体测至今已上 '+str(crs_after_last_msr)+' 节课。建议体测。'
            rbt=WeComRobot(work_dir=self.work_dir,dl_taken_fn=self.dl_taken_fn)
            ins_xlsx=os.path.join(self.work_dir,'03-教练管理','教练资料','教练信息.xlsx')
            df_ins=get_data.InsInfo().get_info(ins_fn=ins_xlsx)
            ins_tel=df_ins[(df_ins['员工编号']==ins_name[0:8]) & (df_ins['姓名']==ins_name[8:])]['电话'].tolist()[0]
            ins_real_name=ins_name[8:]
            ins_txt_send='你的以下会员已到建议体测时间'
            rbt.send(txt_to_send=txt_to_send,ins_inform_txt=ins_txt_send,ins_name=ins_real_name,ins_tel=ins_tel)
            return txt_to_send

    
    def grp_send(self,crs_types=['常规私教课','初级团课'],nowtime='20220804',ins_name='MHINS001陆伟杰'):
        cus_list=get_data.CusInfo().get_cus_list(os.path.join(self.work_dir,'01-会员管理','会员资料'))
        for cus in cus_list:
            print(cus)
            try:
                res=self.send_body_note(cus_name=cus,crs_types=crs_types,nowtime=nowtime,ins_name=ins_name)
                print(res)
            except Exception as e:
                print(e)
            # res=self.send_body_note(cus_name=cus,crs_types=crs_types,nowtime=nowtime,ins_name=ins_name)

if __name__=='__main__':
    # m=WeComRobot(work_dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室')
    # m.send_data(cus_name='MH010苏云',crs_type='初级团课',crs_date='20220804',crs_time='1000-1100',ins='MHINS002韦越棋')
    # m.group_send(y_m='202206',crs_type='常规私教课')

    p=Notification(work_dir='D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室',dl_taken_fn='e:\\temp\\minghu\\教练工作日志.xlsx')
    # res=p.data_prep(cus_name='MH003吕雅颖',crs_types=['常规私教课','限时私教课','初级团课'],nowtime='20220804')
    # res=p.send_body_note(cus_name='MH069骆莹莹',crs_types=['常规私教课','限时私教课','初级团课'],nowtime='20220804')
    # print(res)
    # p.get_bodydata(cus_name='MH003吕雅颖')

    #运行之前先将教练工作日志下载到本地
    p.grp_send(crs_types=['常规私教课','限时私教课'],nowtime='20220809',ins_name='MHINS001陆伟杰')
