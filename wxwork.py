import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import get_data
import requests
import json

class WeComRobot:
    def __init__(self,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室'):
        self.work_dir=work_dir
        self.url='https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=cd85d3e6-c252-4c02-a631-984c813c6f67'    

    def send(self,cus_name='MH016徐颖丽',crs_type='常规私教课',crs_date='20220527',crs_time='1000-1100',ins='MHINS001陆伟杰'):
        txt_ready=get_data.ReadCourses(work_dir=self.work_dir)        
        txt_to_send=txt_ready.exp_txt(cus_name=cus_name,crs_type=crs_type,crs_date=crs_date,crs_time=crs_time,ins=ins)


        print(txt_to_send)        
        df_ins=txt_ready.ins_info(ins=ins)
        ins_tel=str(df_ins['电话'].tolist()[0])
        
        
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
                "content": ins[8:]+'教练，'+'您的会员约课信息如下：',
                # "mentioned_list":["wangqing","@all"],
                "mentioned_mobile_list":[ins_tel]
                }
            }
        
        requests.post(self.url,json=ins_data).json()
        requests.post(self.url,json=data).json()

        print('发送完成')


if __name__=='__main__':
    m=WeComRobot(work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室')
    m.send(cus_name='MH016徐颖丽',crs_type='常规私教课',crs_date='20220527',crs_time='1000-1100',ins='MHINS002韦越棋')
