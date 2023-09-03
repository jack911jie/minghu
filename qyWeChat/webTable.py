import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules'))
import get_data
import days_cal
import requests
import json
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class WeComDisk:
    def __init__(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室',dl_taken_fn='e:\\temp\\minghu\\教练工作日志.xlsx'):
        self.work_dir=work_dir
        self.dl_taken_fn=dl_taken_fn
        self.url='https://qyapi.weixin.qq.com/cgi-bin/wedrive/space_create?access_token='

    def connect(self,access_token):
        url=self.url+access_token
        data={
                "space_name": "99_test",
                "auth_info": [{
                    "type": 1
                    # "userid": "USERID",
                    # "auth": 7
                }]
            }

        res=requests.post(url,json=data).json()
        print(res)


if __name__=='__main__':
    p=WeComDisk()
    p.connect(access_token='q1vfO5bDi9X7T34gIoGpR1e83UwmGREBCwQbIJXQC1H_rsp6IBgvAORCllboM_GA6VJZAJNEC6G_Ik2vVcaggT_lDvShlRWpXfZDQcVkDdy6kOBTvwGhKCCD-NAnDivxExQN15SB94uLpaM-ujlo-duKUNg3uBPWnzJcwrkHMq_slyYqYJIvXbmx9Qz_joajk2n8yymyG2jPzxUQ17IbYg')
