import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),'modules'))
import days_cal
import composing
import json
import pandas as pd
import numpy as np
from datetime import datetime
import re
from PIL import Image,ImageDraw,ImageFont

class MingHu:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        with open (os.path.join(self.dir,'config.dazhi'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line) 

        self.cus_file_dir=config['会员档案文件夹']

    def fonts(self,font_name,font_size):
        fontList={
            '丁永康硬笔楷书':'E:\\健身项目\\minghu\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',
            '方正韵动粗黑':'E:\\健身项目\\minghu\\fonts\\FZYunDongCuHei.ttf',
            '微软雅黑':'E:\\健身项目\\minghu\\fonts\\msyh.ttc',
            '上首金牛':'E:\\健身项目\\minghu\\fonts\\ShangShouJinNiuTi-2.ttf',
            '杨任东石竹体':'E:\\健身项目\\minghu\\fonts\\yangrendongzhushi-Regular.ttf',
            '优设标题黑':'E:\\健身项目\\minghu\\fonts\\yousheTitleHei.ttf'       
        }

        # ImageFont.truetype('j:\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',font_size)


        return ImageFont.truetype(fontList[font_name],font_size)

    def put_txt_img(self,img,t,total_dis,xy,dis_line,fill,font_name,font_size,addSPC='None'):
        
        fontInput=self.fonts(font_name,font_size)            
        if addSPC=='add_2spaces': 
            ind='yes'
        else:
            ind='no'
            
        # txt=self.split_txt(total_dis,font_size,t,Indent='no')
        txt,p_num=composing.split_txt_Chn_eng(total_dis,font_size,t,Indent=ind)

        # font_sig = self.fonts('丁永康硬笔楷书',40)
        draw=ImageDraw.Draw(img)   
        # logging.info(txt)
        n=0
        for t in txt:              
            m=0
            for tt in t:                  
                x,y=xy[0],xy[1]+(font_size+dis_line)*n
                if addSPC=='add_2spaces':   #首行缩进
                    if m==0:    
                        # tt='  '+tt #首先前面加上两个空格
                        # logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                        # logging.info(tt)
                        draw.text((x-font_size*0.2,y), tt, fill = fill,font=fontInput) 
                    else:                       
                        # logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                        # logging.info(tt)
                        draw.text((x,y), tt, fill = fill,font=fontInput)  
                else:
                    # logging.info('字数：'+str(len(tt))+'，坐标：'+str(x)+','+str(y))
                    # logging.info(tt)
                    draw.text((x,y), tt, fill = fill,font=fontInput)  
 
                m+=1
                n+=1

    def read_excel(self,cus='MH001韦美霜'):
        xls_name=os.path.join(self.cus_file_dir,cus+'.xlsx')
        df_basic=pd.read_excel(xls_name,sheet_name='基本情况')    
        df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        df_infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        return [df_basic,df_body,df_infos]


    def exp_cus_prd(self,cus='MH001韦美霜',start_time='20150101',end_time=''):
        df=self.read_excel(cus=cus)
        df_basic=df[0] #基本情况
        df_body=df[1] #身体围度
        infos=df[2] #训练情况

        #------------基本情况--------
        out=Vividict()       

        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        
        
        # df_basic=pd.read_excel(xls_name,sheet_name='基本情况')        
        out['nickname']=df_basic['昵称'].tolist()[0] #昵称
        out['sex']=df_basic['性别'].tolist()[0] #性别

        #年龄
        if not df_basic['出生年月'].empty:
            birth=df_basic['出生年月'].tolist()[0]
            if len(str(birth))==4:
                birth=str(birth)+'0101'
                age=days_cal.calculate_age(birth)
                out['age']=age
            elif len(str(birth))==6:
                birth=str(birth)+'01'
                age=days_cal.calculate_age(birth)
                out['age']=age
            elif len(str(birth))==8:
                age=days_cal.calculate_age(str(birth))
                out['age']=age
            else:
                out['age']=''            
        else:
            out['age']='' 
        

        #------------身体数据--------
        # df_body=pd.read_excel(xls_name,sheet_name='身体数据')
        df_body=df_body[(df_body['时间']>=start_time) & (df_body['时间']<=end_time)] #根据时间段筛选记录
        df_body=df_body.fillna(0)
        if df_body.empty:
            out['body']=''
        else:
            body_recent=df_body[df_body['时间']==df_body['时间'].max()].values.tolist()[0]       
            out['body']['time']=body_recent[0]
            out['body']['ht']=body_recent[1]
            out['body']['wt']=body_recent[2]
            out['body']['bfr']=body_recent[3]
            out['body']['chest']=body_recent[4]
            out['body']['l_arm']=body_recent[5]
            out['body']['r_arm']=body_recent[6]
            out['body']['waist']=body_recent[7]
            out['body']['hip']=body_recent[8]
            out['body']['l_leg']=body_recent[9]
            out['body']['r_leg']=body_recent[10]
            out['body']['l_calf']=body_recent[11]
            out['body']['r_calf']=body_recent[12]
            

        #------------训练数据--------
        # infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos=infos.iloc[:,0:10] #取前10列
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','次数','教练姓名','教练评语']

        infos=infos[(infos['时间']>=start_time) & (infos['时间']<=end_time)] #根据时间段筛选记录

        #起止日期
        out['interval']=[infos['时间'].min(),infos['时间'].max()]  
        out['interval_input']=[start_time,end_time]

        #次数
        train_times=infos.groupby(['时间'],as_index=False).nunique()['时间'].nunique()
        out['train_times']=train_times

        #抗阻训练
        train_dates=infos.groupby(['时间','目标肌群'])
        train_big_type=[]
        for dt,itm in train_dates:
            train_big_type.append(list(dt))
        df_train_big_type=pd.DataFrame(train_big_type)    
        if not df_train_big_type.empty:   
            df_train_big_type.columns=['时间','目标肌群']
            _sum_train_items=df_train_big_type.groupby(['目标肌群'],as_index=False)
            sum_train_items=pd.DataFrame(_sum_train_items.count())  
            sum_train_items.dropna(axis=0, how='any', inplace=True)
            sum_train_items=sum_train_items[sum_train_items['目标肌群']!=' '].values
            if len(sum_train_items)>0:
                for itm in sum_train_items:
                    out['train']['muscle'][itm[0]]=itm[1]
            else:
                out['train']['muscle']=''
        else:
            out['train']['muscle']=''

        #有氧训练总时长
        out['train']['oxy_time']=infos['有氧时长'].sum()

        # print(out)
        return out

    def draw(self,cus='MH001韦美霜',start_time='20150101',end_time=''):

        def txts():
            infos=self.exp_cus_prd(cus=cus,start_time=start_time,end_time=end_time)
        
            txts=Vividict()
            #文字
            nickname=infos['nickname']
            sex=infos['sex']
            if sex=='女':
                sex='美女'
            else:
                sex='帅哥'

            txts['sex']=sex
            txts['age']=infos['age']
        
            #测量
            latest_msr_time=infos['body']['time']
            txts['latest_msr_time']=datetime.strftime(latest_msr_time,'%Y年%m月%d日')
            txts['ht']='身高 '+str(infos['body']['ht'])+'厘米'
            txts['wt']='体重 '+str(infos['body']['wt']) +'千克'
            txts['bfr']='体脂率 '+str(infos['body']['bfr']) 
            txts['chest']='胸围 '+str(infos['body']['chest']) 
            txts['l_arm']='左臂围 '+str(infos['body']['l_arm']) 
            txts['r_arm']='右臂围 '+str(infos['body']['r_arm']) 
            txts['waist']='腰围 '+str(infos['body']['waist']) 
            txts['hip']='臀围 '+str(infos['body']['hip']) 
            txts['l_leg']='左大腿围 '+str(infos['body']['l_leg']) 
            txts['r_leg']='右大腿围 '+str(infos['body']['r_leg']) 
            txts['l_calf']='左小腿围 '+str(infos['body']['l_calf']) 
            txts['r_calf']='右大腿围 '+str(infos['body']['r_calf']) 

            #训练情况
            intervals_input=infos['interval_input'][1]-infos['interval_input'][0]
            txts['intervals_train_0']='您在{0}-{1}的'.format(datetime.strftime(infos['interval_input'][0],'%Y年%m月%d日'),datetime.strftime(infos['interval_input'][1],'%Y年%m月%d日'))
            txts['intervals_train_1']='{0}天里锻炼了{1}次'.format(str(intervals_input.days+1),str(infos['train_times']))

            if infos['train']:
                t=''
                for items in infos['train']:
                    # txts['train_content'][items]='    '+str(infos[items])+' 次'
                    
                    if items=='muscle':
                        for k in infos['train'][items]:
                            t=t+str(k)+'    '+str(infos['train'][items][k])+'次\n'
                        
                    elif items=='oxy_time':
                        _oxy_time=infos['train'][items]
                        if _oxy_time>60:
                            if _oxy_time%60==0:
                                _oxy_time='有氧训练    '+str(int(_oxy_time//60))+'分钟\n'
                            else:
                                _oxy_time='有氧训练    '+str(int(_oxy_time//60))+'分钟'+str(int(_oxy_time%60))+'秒\n'
                        t=t+_oxy_time
                        t.rstrip()
                        
                        txts['train_content']=t.rstrip()
            else:
                txts['train_content']=''

        
            return txts

        def exp_pic(t):
            # print(t)
            dis_line=20
            ft_size=36
            num_prgr=len(t['train_content'].split('\n'))
            y_item=dis_line*(num_prgr-1)+ft_size*num_prgr+50
            img = Image.new("RGB",(684,y_item),(255,255,255))
            draw=ImageDraw.Draw(img)
            # draw.text((10,10), t['train_content'], fill = '#ff9966',font=self.fonts('杨任东石竹体',36))  #大题目)
            self.put_txt_img(img=img,t=t['train_content'],total_dis=600,xy=[40,20],dis_line=20,fill='#ff9966',font_name='杨任东石竹体',font_size=36)
            img.show()

        t=txts()
        exp_pic(t)


            
        

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    p=MingHu()
    p.draw(cus='MH001韦美霜',start_time='20200901',end_time='20201107')


