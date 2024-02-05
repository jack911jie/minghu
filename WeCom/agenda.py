import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules'))
import readconfig
import json
# from urllib import parse, request
import requests
from datetime import datetime


class WeCom:
    def __init__(self):
        config_fn=os.path.join(os.path.dirname(__file__),'config','minghu_corp_ip.config')
        config=readconfig.exp_json2(config_fn)
        self.corp_id=config['corp_id']
        self.secret=config['secret']


    def get_access_token(self,access_token_fn='d:\\temp\\minghu\\access_token\\access_token.txt'):        
        try:
            with open(os.path.join(access_token_fn),'r',encoding='utf-8') as f:
                lines=f.readlines()
            save_datetime=datetime.strptime(lines[0].strip(),'%Y-%m-%d %H:%M:%S')     
            nowdatetime=datetime.now()
            timeinterval=nowdatetime-save_datetime

            #超过2小时的则重新获取access_token
            if timeinterval.total_seconds()>7200:
                url=f'https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={self.corp_id}&corpsecret={self.secret}'
                response = requests.get(url)
            
                if response.status_code == 200:
                    result = response.json()
                    access_token = result.get("access_token")
                    if access_token:
                        txt=datetime.now().strftime('%Y-%m-%d %H:%M:%S')+'\n'+access_token
                        with open(os.path.join(access_token_fn),'w',encoding='utf-8') as f:
                            f.write(txt)
                        return access_token
                    else:
                        print("获取 Access Token 失败：" + result.get("errmsg", ""))
                else:
                    print("请求失败")
                
                return None
            #2小时以内的从文本读取access_token
            else:
                return lines[1].strip()
        except Exception as err:
            print('读取本地缓存access_token错误:',err)       
       

    def create_schedule(self,userids, desp, start_time, end_time,
                        access_token_fn='d:\\temp\\minghu\\access_token\\access_token.txt'):
        access_token=self.get_access_token(access_token_fn=access_token_fn)
        # url = f"https://qyapi.weixin.qq.com/cgi-bin/oa/schedule/add?access_token={access_token}&debug=1"
        url = f"https://qyapi.weixin.qq.com/cgi-bin/oa/schedule/add?access_token={access_token}"
        start_date=self.trans_date(start_time)
        end_date=self.trans_date(end_time)
        
        data = {
            "schedule": {		
                "start_time": start_date,
                "end_time": end_date,
                "attendees": [{
                    "userid": userid
                } for userid in userids],
                "summary": "【限时课程会员到期】点击查看名单",
                "description":desp,
                "reminders": {
                    "is_remind": 1,
                    "remind_before_event_secs": 86400,
                    "timezone": 8
                }
            }
        }
        
        response = requests.post(url,json=data)
        
        if response.status_code == 200:
            result = response.json()
            if result["errcode"] == 0:
                print("日程创建成功")
            else:
                print("日程创建失败：" + result["errmsg"])
        else:
            print("请求失败")

    def trans_date(self,date_input='2023-7-16 8:00:00'):
        # date_input=date_input+' 8:00:00'
        # date_input=datetime.now()
        try:
            dt=datetime.timestamp(datetime.strptime(date_input,'%Y-%m-%d %H:%M:%S'))
            # print(dt)
        except Exception as err:
            dt=datetime.timestamp(date_input)
            # print(dt)
        return int(dt)
            

 
if __name__=='__main__':
    p=WeCom()
    start_time='2023-7-22 8:00:00'
    end_time=datetime.strptime(start_time.split(' ')[0]+' 23:00:00','%Y-%m-%d %H:%M:%S')
    p.create_schedule(userids=['AXiao'], 
                      desp='MH101李测试，\nMH102王测试,\nMH333张测试', 
                      start_time=start_time,
                      end_time=end_time,                      
                    access_token_fn='e:\\temp\\minghu\\access_token\\access_token.txt')

    # 配置的可信域名为：jack911jie.github.io
    # 需配置好ip，在https://work.weixin.qq.com/wework_admin/frame#/apps/modApiApp/5629501898608866 页面中配置好企业可信IP，即：将公网IP写入。
    # 配置好以上两项后，才能使用代码写入日程。
    # start_time：日程开始时间，格式为：2023-7-16 8:00:00
    # end_time：日程结束时间，格式同start_time
    # 默认设置为提前一天提醒，不重复。
    # access_token_fn为临时缓存文件，如2小时内发起，则从文件读取，否则根据corp_id及secret重新获取access_token并写入该文件。