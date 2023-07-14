import os
import json
# from urllib import parse, request
import requests

class WeCom:
    def __init__(self):
        pass

    def get_access_token(self,corp_id='ww1b8f832af7bf55a1',secret='Ssnu5W15iuGhn-G_lTola2mAaMRg_GNutuvtb79viZE'):
        url='https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={}&corpsecret={}'.format(corp_id,secret)
        try:
            appid = "you appid"
            secret = "you secret"
            textmod = {"grant_type": "client_credential",
                "appid": appid,
                "secret": secret
            }
            textmod = parse.urlencode(textmod)
            header_dict = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko'}
            url = 'https://api.weixin.qq.com/cgi-bin/token'
            req = request.Request(url='%s%s%s' % (url, '?', textmod), headers=header_dict)
            res = request.urlopen(req)
            res = res.read().decode(encoding='utf-8')
            res = json.loads(res)
            access_token = res["access_token"]
            print('access_token:', access_token)
            return access_token
        except Exception as e:
            print(e)
            return False

    def creat_agenda(self):
        access_token='UZzRFwA3hoP2YClhsv9J7t43fZqhLqtt__0El8UnVgZ5k2j8z0qiFiU-tGhd3vSvHfGgRYz_gPZtVAtzXwoJjIZJwQsRg5fIUUWFR6jDyUu3nvsjgpVgwcq9gBbvWiyE6SvxZ_Jp1G6N1PPfOvtD5AehNWUZSi6-koPGDsPij1u5ZczEiQAhbR5JjFVXJnWoTo0bsgYHdKBGgqRWzJUTUA'
        url='https://qyapi.weixin.qq.com/cgi-bin/oa/schedule/add?access_token='+access_token

        data={
                "schedule": {                    
                    "start_time": 1689379200,
                    "end_time": 1689379200,                  
                    "summary": "测试铭湖日程",
                    "description": "2.0版本需求初步评审",
                    "attendees": [{
                            "userid": "AXiao"
                        }],
                    "reminders": {
                        "is_remind": 1,
                        "remind_before_event_secs": 3600,
                        "is_repeat": 0,
                        "repeat_type": 7,
                        "repeat_until": 1689379200,
                        "is_custom_repeat": 0,
                        "repeat_interval": 1,
                        "repeat_day_of_week": [3, 7],
                        "repeat_day_of_month": [10, 21],
                        "timezone": 8
                    },
                    "location": "南宁市民族大道"
                }
            }

        res=requests.post(url,data=data)
        print(res.status_code,res.text)


if __name__=='__main__':
    p=WeCom()
    p.creat_agenda()