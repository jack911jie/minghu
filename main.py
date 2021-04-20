import os
import sys
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),'modules'))
import pic_transfer
import days_cal
import readconfig
import composing
import json
import pandas as pd
import numpy as np
from datetime import datetime
import re
import random
from PIL import Image,ImageDraw,ImageFont
import random
import matplotlib.pyplot as plt 
import matplotlib.font_manager as fm
# from matplotlib.backends.backend_agg import FigureCanvasAgg
plt.rcParams['font.sans-serif']=['SimHei']  # 黑体

class MingHu:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        config=readconfig.exp_json(os.path.join(self.dir,'configs','config.minghu'))
        self.cus_file_dir=config['会员档案文件夹']
        self.material_dir=config['素材文件夹']
        self.ins_dir=config['教练文件夹']
        self.slogan_dir=config['文案文件夹']
        self.save_dir=config['输出文件夹']

    def fonts(self,font_name,font_size):
        fontList=readconfig.exp_json(os.path.join(self.dir,'configs','FontList.minghu'))
        # print(fontList)
        return ImageFont.truetype(fontList[font_name],font_size)

    def color_list(self,sex='美女',color_name=''):

        if sex=='美女':
            if color_name=='':
                color_name='light_pink'
            if color_name=='light_orange':
                colors={
                    'comment_bg':'#fff4ee',
                    'title_bg':'#fff4ee',
                    'logo_bg':'#fff4ee',
                    'train_content_bg':'#ffffff',
                    'txt_person':'#ff6667',
                    'txt_title':'#ff9c6c',
                    'txt_date':'#ff9c6c',
                    'txt_fix':'#898886',
                    'txt_dimension':'#000000',
                    'txt_train':'#ff9c6c',
                    'txt_slogan':'#cd8c52',
                    'gym_info':'#693607'
                }
            elif color_name=='light_pink':
                colors={
                    'comment_bg':'#fdf7f9',
                    'title_bg':'#dfcbe4',
                    'logo_bg':'#fbfbfb',
                    'train_content_bg':'#ffffff',
                    'txt_person':'#d584d0',
                    'txt_title':'#ffffff',
                    'txt_date':'#cf86cd',
                    'txt_fix':'#717171',
                    'txt_dimension':'#000000',
                    'txt_train':'#cf86cd',
                    'txt_slogan':'#b0b0b0',
                    'gym_info':'#b0b0b0'
                }
        elif sex=='帅哥':
            if color_name=='':
                color_name='strong_blue'
            if color_name=='strong_blue':
                colors={
                    'comment_bg':'#e5f5fd',
                    'title_bg':'#e5f5fd',
                    'logo_bg':'#e5f5fd',
                    'txt_person':'#3c5ebb',
                    'txt_title':'#3c5ebb',
                    'txt_date':'#3c5ebb',
                    'txt_fix':'#9c9fa0',
                    'txt_dimension':'#000000',
                    'txt_train':'#3c5ebb',
                    'txt_slogan':'#8da8db',
                    'gym_info':'#2c2e35'
                }
        else:
            pass

        return colors

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
            out['body']['ht_lung']=body_recent[13]
            out['body']['balance']=body_recent[14]
            out['body']['power']=body_recent[15]
            out['body']['flexibility']=body_recent[16]
            out['body']['core']=body_recent[17]

            bfr_data=cals()
            bfr=bfr_data.bfr(age=age,sex=out['sex'],ht=out['body']['ht'],wt=out['body']['wt'],waist=out['body']['waist'],formula=1)
            out['body']['bfr']=bfr

        #------------训练数据--------
        # infos=pd.read_excel(xls_name,sheet_name='训练情况',skiprows=2,header=None)
        infos=infos.iloc[:,0:10] #取前10列
        infos.columns=['时间','形式','目标肌群','有氧项目','有氧时长','力量内容','重量','次数','教练姓名','教练评语']
        # print(infos.dropna(how='all'))
        if infos.dropna(how='all').shape[0]!=0:
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
        else:
            out['train']=''

        # print('201 line:',out)
        return out

    def draw(self,cus='MH001韦美霜',ins='MHINS002韦越棋',start_time='20150101',end_time=''):
        
        def slogan():
            df_slogans=pd.read_excel(os.path.join(self.slogan_dir,'文案.xlsx'))
            # print(df_slogans)
            df_slogans.dropna(axis=0,how='any',subset=['文案'],inplace=True)
            slogans=df_slogans['文案'].values.tolist()
            slogan=random.choice(slogans)
            # print('189 line:',df_slogans)
            return slogan

        def ins_info():
            df_ins=pd.read_excel(os.path.join(self.ins_dir,'教练信息.xlsx'))
            ins_inf={}
            ins_inf['nickname']=df_ins[df_ins['员工编号']==ins[0:8]]['昵称'].values.tolist()[0].strip()
            tel=str(df_ins[df_ins['员工编号']==ins[0:8]]['电话'].values.tolist()[0]).strip()
            tel=tel[0:3]+'-'+tel[3:7]+'-'+tel[7:]
            ins_inf['tel']='电话：'+tel
            return ins_inf

        def txts():
            infos=self.exp_cus_prd(cus=cus,start_time=start_time,end_time=end_time)        
            # print(infos) 
            txts=Vividict()
            #文字
            txts['nickname']=infos['nickname']
            sex=infos['sex']
            if sex=='女':
                sex='美女'
            elif sex=='男':
                sex='帅哥'
            else:
                pass

            txts['sex']=sex
            txts['age']=infos['age']
        
            #测量
            if infos['body']:
                latest_msr_time=infos['body']['time']
                txts['latest_msr_time']=datetime.strftime(latest_msr_time,'%Y年%m月%d日')
                txts['ht']='身高 '+str('{:g}'.format(infos['body']['ht']))+' cm'
                txts['wt']='体重 '+str('{:g}'.format(infos['body']['wt']))+' Kg'
                txts['bfr']='体脂率 '+str(infos['body']['bfr'])
                txts['chest']='胸围 '+str('{:g}'.format(infos['body']['chest']))+' cm'
                txts['l_arm']='左臂围 '+str('{:g}'.format(infos['body']['l_arm']))  +' cm'
                txts['r_arm']='右臂围 '+str('{:g}'.format(infos['body']['r_arm'])) +' cm'
                txts['waist']='腰围 '+str('{:g}'.format(infos['body']['waist'])) +' cm'
                txts['hip']='臀围 '+str('{:g}'.format(infos['body']['hip']))  +' cm'
                txts['l_leg']='左大腿围 '+str('{:g}'.format(infos['body']['l_leg']))  +' cm'
                txts['r_leg']='右大腿围 '+str('{:g}'.format(infos['body']['r_leg']))  +' cm'
                txts['l_calf']='左小腿围 '+str('{:g}'.format(infos['body']['l_calf']))  +' cm'
                txts['r_calf']='右大腿围 '+str('{:g}'.format(infos['body']['r_calf']))  +' cm'
                txts['bfr']='体脂率：'+str('{:.2%}'.format(infos['body']['bfr']))
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

            radar_data={'ht_lung':infos['body']['ht_lung'],'balance':infos['body']['balance'],'power':infos['body']['power'], \
                        'flexibility':infos['body']['flexibility'],'core':infos['body']['core']}
        
            return {'txts':txts,'radar_data':radar_data}

        def radar(data,sex):
            color=self.color_list(sex=sex)
            # print(data)
            # 构造数据
            values = list(data.values())
            feature =list(data.keys()) 

            N = len(values)
            # 设置雷达图的角度，用于平分切开一个圆面
            angles = np.linspace(0, 2 * np.pi, N, endpoint=False)


            # 为了使雷达图一圈封闭起来，需要下面的步骤
            values = np.concatenate((values, [values[0]]))
            angles = np.concatenate((angles, [angles[0]]))

            # print(values,angles)

            # 绘图
            fig = plt.figure(figsize=(6,5))
            # 这里一定要设置为极坐标格式
            ax = fig.add_subplot(111, polar=True)
            # ccl=ax.patch

            # 绘制折线图
            ax.plot(angles, values, 'o-', linewidth=2,color=color['txt_train'])
            # 填充颜色
            ax.fill(angles, values, color=color['txt_train'],alpha=0.25)
            # 添加每个特征的标签
            ax.set_thetagrids(angles * 180 / np.pi, '',color='r',fontsize=13)
            # 设置雷达图的范围
            r_distance=10
            ax.set_rlim(0, r_distance)

            ax.grid(color='#F1E0D6', alpha=0.25, lw=3)
            ax.spines['polar'].set_color('#F1E0D6')
            ax.spines['polar'].set_alpha(0.2)
            ax.spines['polar'].set_linewidth(2)
            # ax.spines['polar'].set_linestyle('-.')

            #项目名称：
            a=[0,0,np.pi/30,-np.pi/50,0,0,0]
            b=[r_distance*1.1,r_distance*1.1,r_distance*1.3,r_distance*1.4,r_distance*1.12]

            e_to_c={'ht_lung': '心肺', 'balance': '平衡', 'power': '力量', 'flexibility': '柔韧性', 'core': '核心'}
            for k,i in enumerate(angles):
                try:
                    # print(k,i,e_to_c[feature[k]])
                    ax.text(i+a[k],b[k],e_to_c[feature[k]],fontsize=18,color=color['txt_train'])
                except:
                    pass

            #分值：
            # c = [1, 0.6, 1.6, 2.3, 1.5, 1,1]
            # print(len(angles))
            # for j,i in enumerate(angles):
            #     try:
            #         r=values[j]-2*i/np.pi
            #         ax.text(i,values[j]+c[j],values[j],color='#218FBD',fontsize=18)
            #     except:
            #         pass

            # 添加标题
            #plt.title('活动前后员工状态表现')
            # 添加网格线
            ax.grid(True,color='grey',alpha=0.1)

            # a=np.arange(0,2*np.pi,0.01)
            # ax.plot(a,10*np.ones_like(a),linewidth=2,color='b')


            ax.set_yticklabels([])
            # plt.savefig(savefilename,transparent=True,bbox_inches='tight')
            # 显示图形
            # plt.show()

            #将matplotlib的图形转换为PIL的对象
            image=pic_transfer.mat_to_pil_img(fig)


            return image

        def save_pic_name(cus):
            save_dir=os.path.join(self.save_dir,cus)
            if not os.path.exists(save_dir):
                os.mkdir(save_dir)
            _date=datetime.strftime(datetime.now(),"%Y%m%d_%H%M%S")
            save_name=os.path.join(save_dir,_date+'_'+cus+'.jpg')
            print('文件名：'+save_name)
            return save_name

        def exp_pic(dat):           

            t=dat['txts']
            color=self.color_list(sex=t['sex'])
            radar_data=dat['radar_data']
            
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
                s_body=960
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

            slogan_txt=slogan()
            slogan_txt=slogan_txt+'\n期待您在铭湖健身遇见更好的自己。'
            dis_line_slogan=15
            ft_size_slogan=36
            # print('327 line',slogan_txt)
            num_prgr_slogan=len(slogan_txt.split('\n'))
            # s_slogan=120
            s_slogan=dis_line_slogan*(num_prgr_slogan-1)+ft_size_slogan*num_prgr_slogan+40
            # print(s_slogan,num_prgr_slogan)

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

            def bg(ins=ins):           
                # y_item=dis_line*(num_prgr-1)+ft_size*num_prgr+50
                img = Image.new("RGB",(720,y_bottom+s_bottom),(255,255,255))
                
                draw=ImageDraw.Draw(img)

                #--------框-----------
                draw.rectangle((0,0,720,y0+s_top),fill=color['comment_bg']) #top
                # draw.rectangle((x_l,y_name,x_r,y_name+s_name),fill='#fff4ee') #name
                y_pic_box=y_name+int(s_name*0.2/2)
                draw.rectangle((x_l+20,y_pic_box,x_l+20+int(s_name*0.8),y_name+int(s_name*0.9)),fill=color['train_content_bg']) #head pic box

                if t['latest_msr_time']!=0:
                    draw.rectangle((x_l,y_title_body,x_l+254,y_title_body+s_title_body),fill=color['title_bg']) #body title
                    draw.rectangle((x_l,y_body,x_r,y_body+s_body),fill=color['comment_bg']) #body

                if t['train_content']:
                    draw.rectangle((x_l,y_title_train,x_l+254,y_title_train+s_title_train),fill=color['title_bg']) #train title
                    draw.rectangle((x_l,y_train,x_r,y_train+s_train),fill=color['comment_bg']) #train
                    y_train_content_bottom=y_train+200+s_train_content
                    draw.rectangle((x_l+40,y_train+200,x_r-40,y_train_content_bottom),fill=color['train_content_bg']) #train content                    
                draw.rectangle((x_l,y_slogan,x_r,y_slogan+s_slogan),fill=color['logo_bg']) #slogan
                draw.rectangle((x_l,y_logo,x_r,y_logo+s_logo),fill=color['logo_bg']) #logo
                draw.rectangle((0,y_bottom,720,y_bottom+s_bottom),fill=color['logo_bg']) #bottom

                 #--------图片-----------

                #头像
                if t['sex']=='美女':
                    pics_f=[]
                    for fn in os.listdir(self.material_dir):
                        if re.match(r'女性头像\d{2}.jpg',fn) or re.match(r'女性头像\d{2}.png',fn):
                            pics_f.append(fn)
                    filename=random.choice(pics_f)                   
                    pic_head_src=os.path.join(self.material_dir,filename)
                else:
                    pics_f=[]
                    for fn in os.listdir(self.material_dir):
                        if re.match(r'男性头像\d{2}.jpg',fn) or re.match(r'男性头像\d{2}.png',fn):
                            pics_f.append(fn)
                    filename=random.choice(pics_f)                   
                    pic_head_src=os.path.join(self.material_dir,filename)
                    # pic_head_src=os.path.join(self.material_dir,'男性头像03.png')
                    # pass #男性
                    # pass
                pic_head=Image.open(pic_head_src)
                pic_head=pic_transfer.round_corner(pic_head)
                w_head,h_head=pic_head.size
                pic_head=pic_head.resize((int(w_head*120/h_head),120))
                r1,g1,b1,a1=pic_head.split()
                img.paste(pic_head,(x_l+20+int((s_name*0.8-pic_head.size[0])/2),y_name+30),mask=a1)

                #模特
                if t['latest_msr_time']!=0:
                    if t['sex']=='美女':
                        model_pic='size_model_female.png'
                    elif t['sex']=='帅哥':
                        model_pic='size_model_male.png'
                    else:
                        pass
                    model_src=os.path.join(self.material_dir,model_pic)
                    pic_model=Image.open(model_src)
                    w_model,h_model=pic_model.size
                    pic_model=pic_model.resize((280,int(h_model*280/w_model)))
                    r2,g2,b2,a2=pic_model.split()
                    img.paste(pic_model,(x_l+int((block_wid-pic_model.size[0])/2),y_title_body+175),mask=a2)

                    #雷达图
                    img_radar=radar(radar_data,sex=t['sex'])
                    # img_radar.show()
                    # print(img_radar.size)
                    img_radar=img_radar.resize((400,int(400*img_radar.size[1]/img_radar.size[0])))
                    img.paste(img_radar,(x_l+int((block_wid-img_radar.size[0])/2),y_title_body+175+int(h_model*280/w_model)+100))

                if t['train_content']:
                    teach_pic_src=os.path.join(self.material_dir,'指导.png')
                    pic_teach=Image.open(teach_pic_src)
                    w_teach,h_teach=pic_teach.size
                    pic_teach=pic_teach.resize((150,int(h_teach*150/w_teach)))
                    r3,g3,b3,a3=pic_teach.split()
                    img.paste(pic_teach,(x_r-150-40,y_train+200+s_train_content-150),mask=a3)
                    # img.paste(pic_teach,())   x_l+40,y_train+200,x_r-40,y_train+200+s_train_content

                #logo
                logo=Image.open(os.path.join(self.ins_dir,'minghulogo.png'))
                w_logo,h_logo=logo.size
                logo=logo.resize((300,int(h_logo*300/w_logo)))
                r4,g4,b4,a4=logo.split()
                img.paste(logo,(int(x_l+(s_logo-300)/2),y_logo+30),mask=a4)

                #qrcode

                qrcode=Image.open(os.path.join(self.ins_dir,ins+'二维码.jpg'))
                w_qrcode,h_qrcode=qrcode.size
                qrcode=qrcode.resize((150,int(h_qrcode*150/w_qrcode)))
                # r5,g5,b5,a5=qrcode.split()
                img.paste(qrcode,(int(x_l+(s_logo-150)/2),y_logo+logo.size[1]+220))

                #------文字-----------
                if t['sex']=='美女':
                    title_01='看看棒棒的自己'
                elif t['sex']=='帅哥':
                    title_01='看看很酷的自己'
                else:
                    pass

                x_nickname=250
                draw.text((x_nickname,110), t['nickname'], fill = color['txt_person'],font=self.fonts('汉仪心海楷体w',80))  #姓名
                if t['sex']=='帅哥':
                    sex='先生'
                elif t['sex']=='美女':
                    sex='女士'
                else:
                    print('warning:性别有误')
                draw.text((x_nickname+len(t['nickname'])*80+30,150), sex, fill = color['txt_person'],font=self.fonts('汉仪心海楷体w',40))  #性别
                if t['latest_msr_time']!=0:
                    draw.text((x_l+30,y_title_body+5), title_01, fill = color['txt_title'],font=self.fonts('上首金牛',30))  #看看棒棒的自己
                    draw.text((x_l+115,y_title_body+65), '您最近一次测量身体围度，是在', fill = color['txt_fix'],font=self.fonts('aa楷体',36))  #您最近一次测量身体围度
                    draw.text((x_l+205,y_title_body+115), t['latest_msr_time'], fill = color['txt_date'],font=self.fonts('aa楷体',40))  #测围度日期

                    draw.text((x_l+20,y_title_body+190), t['r_arm'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #右臂
                    draw.text((x_l+75,y_title_body+270), t['hip'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  # 臀
                    draw.text((x_l+20,y_title_body+380), t['r_leg'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #右大腿
                    draw.text((x_l+20,y_title_body+460), t['r_calf'],fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #右小腿
                    draw.text((x_l+500,y_title_body+190), t['chest'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #胸
                    draw.text((x_l+500,y_title_body+240), t['l_arm'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #左臂
                    draw.text((x_l+500,y_title_body+280), t['waist'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #腰
                    draw.text((x_l+500,y_title_body+370), t['l_leg'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #左大腿
                    draw.text((x_l+500,y_title_body+470), t['l_calf'], fill = color['txt_dimension'],font=self.fonts('杨任东石竹体',25))  #左小腿
                    draw.text((x_l+180,y_title_body+550), t['wt'], fill = color['txt_dimension'],font=self.fonts('aa楷体',25))  #体重
                    draw.text((x_l+360,y_title_body+550), t['bfr'], fill = color['txt_dimension'],font=self.fonts('aa楷体',25))  #体脂率         

                if t['train_content']:
                    draw.text((x_l+30,y_title_train+5), '看看努力的自己', fill = color['txt_title'],font=self.fonts('上首金牛',30))  #看看努力的自己                    
                    if t['intervals_train_1']:
                        draw.text((x_l+45,y_train+35), t['intervals_train_0'], fill =color['txt_fix'],font=self.fonts('aa楷体',34))  #您在。。。
                        ft_size_days=40
                        x_days=composing.center_align_x(start_x=x_l,wide=block_wid,ft_size=ft_size_days,t=t['intervals_train_1'])
                        draw.text((x_days,y_train+85), t['intervals_train_1'], fill =color['txt_train'],font=self.fonts('aa楷体',ft_size_days))  #XX天里（居中）
                        draw.text((x_l+180,y_train+140), '完成了下面的训练内容', fill = color['txt_fix'],font=self.fonts('aa楷体',32))  #完成了下面的训练内容
                        self.put_txt_img(img,t=t['train_content'],total_dis=420,xy=[x_l+95,y_train+230],dis_line=16,fill=color['txt_train'],font_name='杨任东石竹体',font_size=38)
                        percent=random.randint(70,93)
                        draw.text((x_l+145,y_train_content_bottom+20), '击败了铭湖健身 {} 的会员!'.format(str(percent)+'%'), fill = color['txt_train'],font=self.fonts('aa楷体',32))  #击败了
                    else:
                        draw.text((x_l+145,y_train+45), t['intervals_train_0'], fill = color['txt_fix'],font=self.fonts('aa楷体',40))  #您在。。。
                        draw.text((x_l+160,y_train+85), t['intervals_train_1'], fill = color['txt_train'],font=self.fonts('aa楷体',40))  #XX天里
                        draw.text((x_l+50,y_train+115), '你完成了下面的训练内容:', fill = color['txt_fix'],font=self.fonts('aa楷体',40))  #完成了下面的训练内容
                        self.put_txt_img(img,t=t['train_content'],total_dis=420,xy=[x_l+95,y_train+230],dis_line=16,fill=color['txt_train'],font_name='杨任东石竹体',font_size=36)
                        draw.text((x_l+55,y_train_content_bottom+20), '保持这样的状态，好身材还远吗？', fill =color['txt_train'],font=self.fonts('aa楷体',40))  #击败了

                # 鸡汤
                draw.text((x_l+20,y_slogan+15),slogan_txt,fill=color['txt_slogan'],font=self.fonts('优设标题黑',ft_size_slogan))

                # addr
                draw.text((x_l+10,y_logo+240),'南宁市青秀区民族大道88-1号铭湖经典A座802室',fill=color['gym_info'],font=self.fonts('微软雅黑',30))
                draw.text((x_l+125,y_logo+310),'让健身变得有趣',fill=color['gym_info'],font=self.fonts('丁永康硬笔楷书',60))

                ins=ins_info()
                draw.text((x_l+255,y_logo+570),ins['nickname'],fill=color['gym_info'],font=self.fonts('丁永康硬笔楷书',50))
                draw.text((x_l+115,y_logo+630),ins['tel'],fill=color['gym_info'],font=self.fonts('丁永康硬笔楷书',40))

                save_name=save_pic_name(cus)
                print(cus)
                img.save(save_name,quality=95,subsampling=0)
                print('完成')

                # img.show()

            bg()

        t=txts()
        # radar(t['radar_data'])
        exp_pic(t)
        # slogan()
        # ins_info()

class FitData2Pic:
    def __init__(self):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        self.default_title='会员健身数据比较'
        config=readconfig.exp_json(os.path.join(self.dir,'configs','config.minghu'))
        self.fn=os.path.join(config['会员档案文件夹'],'MH000唐青剑.xlsx')
        # self.fn='D:\\Documents\\WXWork\\1688851376227744\WeDrive\\铭湖健身工作室\\铭湖健身工作室\\会员MH000唐青剑.xlsx'
        # self.font='/home/jack/data/健身项目/minghu/fonts/msyh.ttc'
        self.font='E:\\铭湖健身\\fonts\\msyh.ttc'

    def to_pic(self,title='',fn='',d_font=''):
        if title=='':
            title=self.default_title
        if fn=='':
            fn=self.fn
        if d_font=='':
            d_font=self.font
        
        myfont = fm.FontProperties(fname=d_font) # 设置字体

        df=pd.read_excel(fn,sheet_name='身体数据')

        x=[datetime.strftime(d,'%Y-%m-%d') for d in df['时间'].tolist()]
        y_wt=df['体重'].tolist()
        y_chest=df['胸围'].tolist()
        y_waist=df['腰围'].tolist()
        y_l_arm=df['左臂围'].tolist()
        y_r_arm=df['右臂围'].tolist()
        y_hip=df['臀围'].tolist()
        y_l_leg=df['左腿围'].tolist()
        y_r_leg=df['右腿围'].tolist()
        y_l_calf=df['左小腿围'].tolist()
        y_r_calf=df['右小腿围'].tolist()

        fig=plt.figure(figsize=(9,20))

        ax1=fig.add_axes([0.1, 0.08, 0.8, 0.12],facecolor='#FFF5FB')
        ax1.plot(x,y_wt,'o-',color='#FF4747',label='体重')
        ax1.set_ylabel('体重(Kg)',fontproperties=myfont,color='#FF4747')
        ax1.tick_params(axis='y',colors='#FF4747')
        ax1.tick_params(axis='x',colors='#A65817')
        ax1.set_xticklabels(x,rotation=25)
        # ax1.legend(prop=myfont)
        ax1.set_ylim(min(y_wt)*0.98,max(y_wt)*1.02)
        for xy in list(zip(x,y_wt)):
            ax1.text(xy[0],xy[1]+0.5,xy[1],color='#FF4747')

        ax2=fig.add_axes([0.1, 0.20, 0.8, 0.12],facecolor='#F5F6FF')
        ax2.plot(x,y_r_calf,marker='s',color='#4D85A6',label='右小腿围')
        ax2.plot(x,y_l_calf,marker='s',color='violet',label='左小腿围')
        ax2.set_ylabel('小腿围(cm)',fontproperties=myfont,color='#4D85A6')
        ax2.tick_params(axis='y',colors='#4D85A6')
        ax2.set_xticks([])
        ax2.legend(prop=myfont)
        ax2.set_ylim(min(y_r_calf)*0.95,max(y_r_calf)*1.05)
        for xy in list(zip(x,y_r_calf)):
            ax2.text(xy[0],xy[1]+0.4,xy[1],color='#4D85A6')
        for xy in list(zip(x,y_l_calf)):
            ax2.text(xy[0],xy[1]-0.9,xy[1],color='violet')

        ax3=fig.add_axes([0.1, 0.32, 0.8, 0.12],facecolor='#F5F6FF')
        ax3.plot(x,y_r_leg,marker='s',color='#4D85A6',label='右大腿围')
        ax3.plot(x,y_l_leg,marker='s',color='violet',label='左大腿围')
        ax3.set_ylabel('大腿围(cm)',fontproperties=myfont,color='#4D85A6')
        ax3.tick_params(axis='y',colors='#4D85A6')
        ax3.set_xticks([])
        ax3.legend(prop=myfont)
        ax3.set_ylim(min(y_r_leg)*0.95,max(y_r_leg)*1.05)
        for xy in list(zip(x,y_r_leg)):
            ax3.text(xy[0],xy[1]+0.4,xy[1],color='#4D85A6')
        for xy in list(zip(x,y_l_leg)):
            ax3.text(xy[0],xy[1]-1.2,xy[1],color='violet')

        ax4=fig.add_axes([0.1, 0.44, 0.8, 0.12],facecolor='#F5F6FF')
        ax4.plot(x,y_r_arm,marker='s',color='#4D85A6',label='右臂围')
        ax4.plot(x,y_l_arm,marker='s',color='violet',label='左臂围')
        ax4.set_ylabel('臂围(cm)',fontproperties=myfont,color='#4D85A6')
        ax4.tick_params(axis='y',colors='#4D85A6')
        ax4.set_xticks([])
        ax4.legend(prop=myfont)
        ax4.set_ylim(min(y_r_arm)*0.95,max(y_r_arm)*1.05)
        for xy in list(zip(x,y_r_arm)):
            ax4.text(xy[0],xy[1]+0.3,xy[1],color='#4D85A6')
        for xy in list(zip(x,y_l_arm)):
            ax4.text(xy[0],xy[1]-0.8,xy[1],color='violet')


        ax5=fig.add_axes([0.1, 0.56, 0.8, 0.12],facecolor='#FFFAF4')
        ax5.plot(x,y_waist,marker='s',color='orange',label='腰围')
        ax5.set_ylabel('腰围(cm)',fontproperties=myfont,color='orange')
        ax5.tick_params(axis='y',colors='orange')
        ax5.set_xticks([])
        # ax5.legend(prop=myfont)
        ax5.set_ylim(min(y_waist)*0.95,max(y_waist)*1.05)
        for xy in list(zip(x,y_waist)):
            ax5.text(xy[0],xy[1]+0.5,xy[1],color='orange')

        ax6=fig.add_axes([0.1, 0.68, 0.8, 0.12],facecolor='#FFFAF4')
        ax6.plot(x,y_hip,marker='s',color='orange',label='臀围')
        ax6.set_ylabel('臀围(cm)',fontproperties=myfont,color='orange')
        ax6.tick_params(axis='y',colors='orange')
        ax6.set_xticks([])
        # ax6.legend(prop=myfont)
        ax6.set_ylim(min(y_hip)*0.95,max(y_hip)*1.05)
        for xy in list(zip(x,y_hip)):
            ax6.text(xy[0],xy[1]+0.5,xy[1],color='orange')



        ax7=fig.add_axes([0.1, 0.80, 0.8, 0.12],facecolor='#FFFAF4')
        ax7.plot(x,y_chest,marker='s',color='orange',label='胸围')
        ax7.set_ylabel('胸围(cm)',fontproperties=myfont,color='orange')
        ax7.tick_params(axis='y',colors='orange')
        ax7.set_xticks([])
        # ax4.legend(prop=myfont)
        ax7.set_ylim(min(y_chest)*0.95,max(y_chest)*1.05)
        for xy in list(zip(x,y_chest)):
            ax7.text(xy[0],xy[1]+0.5,xy[1],color='orange')


        ax7.set_title(title,fontproperties=myfont,y=1.1,fontsize=20,color='#BF8D30')

        for ax in fig.axes:
            clr='#BF8D30'
            for bdr in ['left','right','bottom','top']:
                ax.spines[bdr].set_color(clr)


        # plt.savefig('/home/jack/data/temp/mhdata.jpg')
        plt.show()
        return plt
            
class cals:
    def bfr(self,age,sex,ht,wt,waist,formula=1):
            # 女：
            # 参数a=腰围（cm）×0.74
            # 参数b=体重（kg）×0.082+34.89
            # 体脂肪重量（kg）=a－b
            # 体脂率=（身体脂肪总重量÷体重）×100%
            # 男：
            # 参数a=腰围（cm）×0.74
            # 参数b=体重（kg）×0.082+44.74
            # 体脂肪重量（kg）=a－b
            # 体脂率=（身体脂肪总重量÷体重）×100%
        if formula==1:
            if sex=='女':
                k=34.89
            if sex=='男':
                k=44.74
            a=waist*0.74
            b=wt*0.082+k
            fat=a-b

            bfr=fat/wt

        elif formula==2:
            # 1.2×BMI+0.23×年龄-5.4-10.8×性别（男为1，女为0）
            if sex=='女':
                k=0
            if sex=='男':
                k=1

            bmi=wt/((ht/100)*(ht/100))
            bfr=1.2*bmi+0.23*age-5.4-10.8*k


        return bfr

class Vividict(dict):
    def __missing__(self, key):
        value = self[key] = type(self)()
        return value

if __name__=='__main__':
    #根据训练数据生成阶段报告
    p=MingHu()
    p.draw(cus='MH001韦美霜',ins='MHINS002韦越棋',start_time='20200315',end_time='20210320')

    # 根据多次体测数据生成折线图
    # fitdata=FitData2Pic()
    # fitdata.to_pic()

    #计算体脂率
    # my=cals()
    # print(my.bfr(age=40,sex='男',ht=170,wt=63.8,waist=82,formula=1))
