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
import random

class MingHu:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        with open (os.path.join(self.dir,'config.linux'),'r',encoding='utf-8') as f:
            lines=f.readlines()
        _line=''
        for line in lines:
            newLine=line.strip('\n')
            _line=_line+newLine
        config=json.loads(_line) 

        self.cus_file_dir=config['会员档案文件夹']
        self.material_dir=config['素材文件夹']
        self.inst_dir=config['教练文件夹']

    def fonts(self,font_name,font_size):
        # fontList={
        #     '丁永康硬笔楷书':'E:\\健身项目\\minghu\\fonts\\2012DingYongKangYingBiKaiShuXinBan-2.ttf',
        #     '方正韵动粗黑':'E:\\健身项目\\minghu\\fonts\\FZYunDongCuHei.ttf',
        #     '微软雅黑':'E:\\健身项目\\minghu\\fonts\\msyh.ttc',
        #     '上首金牛':'E:\\健身项目\\minghu\\fonts\\ShangShouJinNiuTi-2.ttf',
        #     '杨任东石竹体':'E:\\健身项目\\minghu\\fonts\\yangrendongzhushi-Regular.ttf',
        #     '优设标题黑':'E:\\健身项目\\minghu\\fonts\\yousheTitleHei.ttf'       
        # }

        fontList={
            '丁永康硬笔楷书':'/home/jack/data/健身项目/minghu/fonts/2012DingYongKangYingBiKaiShuXinBan-2.ttf',
            '方正韵动粗黑':'/home/jack/data/健身项目/minghu/fonts/FZYunDongCuHei.ttf',
            '微软雅黑':'/home/jack/data/健身项目/minghu/fonts/msyh.ttc',
            '上首金牛':'/home/jack/data/健身项目/minghu/fonts/ShangShouJinNiuTi-2.ttf',
            '杨任东石竹体':'/home/jack/data/健身项目/minghu/fonts/yangrendongzhushi-Regular.ttf',
            '优设标题黑':'/home/jack/data/健身项目/minghu/fonts/yousheTitleHei.ttf'       
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
        # df_body=df_body[(df_body['时间']>=start_time) & (df_body['时间']<=end_time)] #根据时间段筛选记录
        df_body=df_body[df_body['时间']<=end_time] #根据时间段筛选记录
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

        # print('168 line:',infos)

        #起止日期
        out['interval']=[infos['时间'].min(),infos['时间'].max()]  
        out['interval_input']=[start_time,end_time]

        #次数
        train_times=infos.groupby(['时间'],as_index=False).nunique()['时间'].nunique()
        out['train_times']=train_times

        #抗阻训练
        train_dates=infos.groupby(['时间','目标肌群'])
        # print(train_dates)
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

        # print('201 line:',out)
        return out

    def draw(self,cus='MH001韦美霜',ins='韦越棋',start_time='20150101',end_time=''):

        def txts():
            infos=self.exp_cus_prd(cus=cus,start_time=start_time,end_time=end_time)        
            # print(infos) 
            txts=Vividict()
            #文字
            txts['nickname']=infos['nickname']
            sex=infos['sex']
            if sex=='女':
                sex='美女'
            else:
                sex='帅哥'

            txts['sex']=sex
            txts['age']=infos['age']
        
            #测量
            if infos['body']:
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
            else:
                txts['latest_msr_time']=0

            #训练情况
            # print('242line:',infos['train'])
            if infos['train']:
                t_muscle=''
                t_oxy=''
                for items in infos['train']:
                    # txts['train_content'][items]='    '+str(infos[items])+' 次'
                    # print(items)
                    if items=='muscle':
                        if infos['train'][items]:
                            for k in infos['train'][items]:
                                t_muscle=t_muscle+str(k)+'    '+str(infos['train'][items][k])+'次\n'
                        else:
                            t_muscle=''
                    elif items=='oxy_time':
                        _oxy_time=infos['train'][items]
                        if _oxy_time!=0:
                            if _oxy_time>60:
                                if _oxy_time%60==0:
                                    _oxy_time='有氧训练    '+str(int(_oxy_time//60))+'分钟\n'
                                else:
                                    _oxy_time='有氧训练    '+str(int(_oxy_time//60))+'分钟'+str(int(_oxy_time%60))+'秒\n'
                            t_oxy=t_oxy+_oxy_time
                            t_oxy.rstrip()
                        else:
                            t_oxy=''

                t=t_muscle+t_oxy
                txts['train_content']=t.rstrip()
            else:
                txts['train_content']=''
            
    

            if txts['train_content']:
                intervals_input=infos['interval_input'][1]-infos['interval_input'][0]
                if intervals_input.days==0:
                    txts['intervals_train_0']='今天的你非常棒，因为，'
                    txts['intervals_train_1']=''
                else:
                    txts['intervals_train_0']='您在{0}-{1}的'.format(datetime.strftime(infos['interval_input'][0],'%Y年%m月%d日'),datetime.strftime(infos['interval_input'][1],'%Y年%m月%d日'))
                    txts['intervals_train_1']='{0}天里锻炼了{1}次'.format(str(intervals_input.days+1),str(infos['train_times']))
            else:
                txts['intervals_train_0']=''
                txts['intervals_train_1']=''

            # print(txts)
        
            return txts

        def exp_pic(t):
            # print('279 line:',t)
            dis_line=20
            ft_size=36
            num_prgr=len(t['train_content'].split('\n'))
            block_wid=684

            s_top=40
            gap=20
            s_name=180
            s_title=40
            # r2=15
            if t['latest_msr_time']==0:
                s_title_body=0
                s_body=0
                gap_body=0
            else:
                s_title_body=s_title
                s_body=500
                gap_body=gap

            if t['train_content']:
                s_train_content=dis_line*(num_prgr-1)+ft_size*num_prgr+60
                if s_train_content<300:
                    s_train_content=300
                s_train=s_train_content+300
                gap_train=gap
                s_title_train=s_title
            else:
                s_title_train=0
                s_train_content=0
                s_train=0
                gap_train=0

            s_slogan=120
            s_logo=680
            s_bottom=40

            y0=0
            y_name=y0+s_top+gap

            y_title_body=y_name+s_name+gap
            y_body=y_title_body+s_title_body

            y_title_train=y_body+s_body+gap_body
            y_train=y_title_train+s_title_train

            y_slogan=y_train+s_train+gap_train
            y_logo=y_slogan+s_slogan+gap
            y_bottom=y_logo+s_logo+gap

            x_l=18
            x_r=x_l+block_wid

            def bg():                
                # y_item=dis_line*(num_prgr-1)+ft_size*num_prgr+50
                img = Image.new("RGB",(720,y_bottom+s_bottom),(255,255,255))
                
                draw=ImageDraw.Draw(img)

                #--------框-----------
                draw.rectangle((0,0,720,y0+s_top),fill='#fff4ee') #top
                draw.rectangle((x_l,y_name,x_r,y_name+s_name),fill='#fff4ee') #name
                y_pic_box=y_name+int(s_name*0.2/2)
                draw.rectangle((x_l+20,y_pic_box,x_l+20+int(s_name*0.8),y_name+int(s_name*0.9)),fill='#ffffff') #head pic box

                if t['latest_msr_time']!=0:
                    draw.rectangle((x_l,y_title_body,x_l+254,y_title_body+s_title_body),fill='#fff4ee') #body title
                    draw.rectangle((x_l,y_body,x_r,y_body+s_body),fill='#fff4ee') #body

                if t['train_content']:
                    draw.rectangle((x_l,y_title_train,x_l+254,y_title_train+s_title_train),fill='#fff4ee') #train title
                    draw.rectangle((x_l,y_train,x_r,y_train+s_train),fill='#fff4ee') #train
                    draw.rectangle((x_l+40,y_train+200,x_r-40,y_train+200+s_train_content),fill='#ffffff') #train content
                
                draw.rectangle((x_l,y_slogan,x_r,y_slogan+s_slogan),fill='#fff4ee') #slogan
                draw.rectangle((x_l,y_logo,x_r,y_logo+s_logo),fill='#fff4ee') #logo
                draw.rectangle((0,y_bottom,720,y_bottom+s_bottom),fill='#fff4ee') #bottom

                 #--------图片-----------

                #头像
                if t['sex']=='美女':
                    pic_head_src=os.path.join(self.material_dir,'女性头像01.png')
                else:
                    pass #男性
                pic_head=Image.open(pic_head_src)
                w_head,h_head=pic_head.size
                pic_head=pic_head.resize((int(w_head*120/h_head),120))
                r1,g1,b1,a1=pic_head.split()
                img.paste(pic_head,(x_l+20+int((s_name*0.8-pic_head.size[0])/2),y_name+30),mask=a1)

                #模特
                if t['latest_msr_time']!=0:
                    model_src=os.path.join(self.material_dir,'size_model_female.png')
                    pic_model=Image.open(model_src)
                    w_model,h_model=pic_model.size
                    pic_model=pic_model.resize((280,int(h_model*280/w_model)))
                    r2,g2,b2,a2=pic_model.split()
                    img.paste(pic_model,(x_l+int((block_wid-pic_model.size[0])/2),y_body+s_body-pic_model.size[1]-20),mask=a2)

                if t['train_content']:
                    teach_pic_src=os.path.join(self.material_dir,'指导.png')
                    pic_teach=Image.open(teach_pic_src)
                    w_teach,h_teach=pic_teach.size
                    pic_teach=pic_teach.resize((150,int(h_teach*150/w_teach)))
                    r3,g3,b3,a3=pic_teach.split()
                    img.paste(pic_teach,(x_r-150-40,y_train+200+s_train_content-150),mask=a3)
                    # img.paste(pic_teach,())   x_l+40,y_train+200,x_r-40,y_train+200+s_train_content

                #logo
                logo=Image.open(os.path.join(self.inst_dir,'minghulogo.png'))
                w_logo,h_logo=logo.size
                logo=logo.resize((300,int(h_logo*300/w_logo)))
                r4,g4,b4,a4=logo.split()
                img.paste(logo,(int(x_l+(s_logo-300)/2),y_logo+30),mask=a4)

                #qrcode

                qrcode=Image.open(os.path.join(self.inst_dir,ins+'二维码.jpg'))
                w_qrcode,h_qrcode=qrcode.size
                qrcode=qrcode.resize((150,int(h_qrcode*150/w_qrcode)))
                # r5,g5,b5,a5=qrcode.split()
                img.paste(qrcode,(int(x_l+(s_logo-150)/2),y_logo+logo.size[1]+220))

                #------文字-----------
                x_nickname=250
                draw.text((x_nickname,110), t['nickname'], fill = '#ff6667',font=self.fonts('杨任东石竹体',80))  #姓名
                draw.text((x_nickname+len(t['nickname'])*80+30,150), t['sex'], fill = '#ff6667',font=self.fonts('杨任东石竹体',40))  #性别
                if t['latest_msr_time']!=0:
                    draw.text((x_l+30,y_title_body+5), '看看棒棒的自己', fill = '#ff9c6c',font=self.fonts('上首金牛',30))  #看看棒棒的自己
                    draw.text((x_l+115,y_title_body+65), '您最近一次测量身体围度，是在', fill = '#898886',font=self.fonts('杨任东石竹体',36))  #您最近一次测量身体围度
                    draw.text((x_l+205,y_title_body+115), t['latest_msr_time'], fill = '#ff9c6c',font=self.fonts('杨任东石竹体',40))  #测围度日期

                    draw.text((x_l+50,y_title_body+190), t['r_arm'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #右臂
                    draw.text((x_l+90,y_title_body+270), t['hip'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  # 臀
                    draw.text((x_l+50,y_title_body+380), t['r_leg'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #右大腿
                    draw.text((x_l+50,y_title_body+460), t['r_calf'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #右小腿

                    draw.text((x_l+500,y_title_body+190), t['chest'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #胸
                    draw.text((x_l+500,y_title_body+240), t['l_arm'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #左臂
                    draw.text((x_l+500,y_title_body+280), t['waist'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #腰
                    draw.text((x_l+500,y_title_body+370), t['l_leg'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #左大腿
                    draw.text((x_l+500,y_title_body+470), t['l_calf'], fill = '#000000',font=self.fonts('杨任东石竹体',25))  #左小腿
                
                if t['train_content']:
                    draw.text((x_l+30,y_title_train+5), '看看努力的自己', fill = '#ff9c6c',font=self.fonts('上首金牛',30))  #看看努力的自己                    
                    if t['intervals_train_1']:
                        draw.text((x_l+35,y_train+35), t['intervals_train_0'], fill = '#898886',font=self.fonts('杨任东石竹体',36))  #您在。。。
                        draw.text((x_l+160,y_train+85), t['intervals_train_1'], fill = '#ff9c6c',font=self.fonts('杨任东石竹体',40))  #XX天里
                        draw.text((x_l+160,y_train+140), '完成了下面的训练内容', fill = '#898886',font=self.fonts('杨任东石竹体',38))  #完成了下面的训练内容
                        self.put_txt_img(img,t=t['train_content'],total_dis=420,xy=[x_l+95,y_train+230],dis_line=16,fill='#ff9c6c',font_name='杨任东石竹体',font_size=36)
                        
                        percent=random.randint(70,93)
                        draw.text((x_l+105,y_train+550), '击败了铭湖健身 {} 的会员!'.format(str(percent)+'%'), fill = '#ff9c6c',font=self.fonts('杨任东石竹体',40))  #击败了
                    else:
                        draw.text((x_l+145,y_train+45), t['intervals_train_0'], fill = '#898886',font=self.fonts('杨任东石竹体',45))  #您在。。。
                        # draw.text((x_l+160,y_train+85), t['intervals_train_1'], fill = '#ff9c6c',font=self.fonts('杨任东石竹体',40))  #XX天里
                        draw.text((x_l+50,y_train+115), '你成了下面的训练内容:', fill = '#898886',font=self.fonts('杨任东石竹体',45))  #完成了下面的训练内容
                        self.put_txt_img(img,t=t['train_content'],total_dis=420,xy=[x_l+95,y_train+230],dis_line=16,fill='#ff9c6c',font_name='杨任东石竹体',font_size=36)
                        draw.text((x_l+55,y_train+550), '保持这样的状态，好身材还远吗？', fill = '#ff9c6c',font=self.fonts('杨任东石竹体',40))  #击败了

                # 鸡汤
                draw.text((x_l+20,y_slogan+15),'不经历风雨，怎么见彩虹，\n期待你在铭湖健身遇见更好的自己！',fill='#cd8c52',font=self.fonts('优设标题黑',40))

                # addr
                draw.text((x_l+10,y_logo+240),'南宁市青秀区民族大道88-1号铭湖金曲A座802室',fill='#693607',font=self.fonts('微软雅黑',30))
                draw.text((x_l+125,y_logo+310),'让健身变得有趣',fill='#693607',font=self.fonts('丁永康硬笔楷书',60))
                draw.text((x_l+255,y_logo+570),ins[0]+'教练',fill='#693607',font=self.fonts('丁永康硬笔楷书',50))
                draw.text((x_l+115,y_logo+630),'电话：XXXXXXXXXXX',fill='#693607',font=self.fonts('丁永康硬笔楷书',40))

                img.save('/home/jack/data/temp/minghu_test.jpg',quality=95,subsampling=0)

                img.show()

            bg()

        t=txts()
        exp_pic(t)
        # bg()


            
        

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    p=MingHu()
    p.draw(cus='MH000丽看看',ins='陆伟杰',start_time='20200104',end_time='')


