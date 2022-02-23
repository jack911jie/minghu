import os
import sys
import pic_transfer
import pics_fill
import pandas as pd
from PIL import Image,ImageDraw,ImageFont
import numpy as np
import matplotlib.pyplot as plt 
import matplotlib.font_manager as fm
plt.rcParams['font.sans-serif']=['SimHei']  # 黑体
import readconfig
from datetime import datetime



class DrawRadar:
    def __init__(self):
        pass
        # self.data=
        # {'ht_lung':infos['body']['ht_lung'],'balance':infos['body']['balance'],'power':infos['body']['power'], \
        #                 'flexibility':infos['body']['flexibility'],'core':infos['body']['core']}

    def color_list(self):
        light_orange={
                "comment_bg": "#fff4ee", 
                "title_bg": "#fff4ee", 
                "logo_bg": "#fff4ee", 
                "train_content_bg": "#ffffff", 
                "txt_person": "#ff6667", 
                "txt_title": "#ff9c6c", 
                "txt_date": "#ff9c6c", 
                "txt_fix": "#898886", 
                "txt_dimension": "#000000", 
                "txt_train": "#ff9c6c", 
                "txt_slogan": "#cd8c52", 
                "gym_info": "#693607"
            }

        return light_orange

    def draw(self,data,bgcolor='#ffee00'):
        color=self.color_list()
        # print(data)
        # 构造数据
        # print(data)
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
        fig = plt.figure(facecolor=bgcolor,figsize=(6,5))
        # 这里一定要设置为极坐标格式
        ax = fig.add_subplot(111, polar=True)
        ax.patch.set_facecolor(bgcolor)
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
                ax.text(i+a[k],b[k],feature[k],fontsize=18,color=color['txt_train'])
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


class PeriodChart:
    def __init__(self,font_fn='G:\\健身项目\\minghu\\fonts\\msyh.ttc'):
        self.dir=os.path.dirname(os.path.abspath(__file__))
        self.default_title='会员健身数据比较' 
        # self.fn='D:\\Documents\\WXWork\\1688851376227744\WeDrive\\铭湖健身工作室\\铭湖健身工作室\\会员MH000唐青剑.xlsx'
        # self.font='/home/jack/data/健身项目/minghu/fonts/msyh.ttc'
        self.font=font_fn

    def to_pic(self,cus_dir='e:\\temp\\铭湖健身测试\\会员资料',cus_fn='MH003吕雅颖.xlsx',start_time='20000101',end_time='',d_font='',title='',bgcolor='#ffee00',items=['wt','cal','arm','leg','waist','hip','chest']):
        if title=='':
            # title=self.default_title
            pass
        if d_font=='':
            d_font=self.font

        start_time=datetime.strptime('-'.join([start_time[0:4],start_time[4:6],start_time[6:]]),'%Y-%m-%d')
        if end_time=='':
            end_time=datetime.now()
        else:
            end_time=datetime.strptime('-'.join([end_time[0:4],end_time[4:6],end_time[6:]]),'%Y-%m-%d')
        
        myfont = fm.FontProperties(fname=d_font) # 设置字体

        fn=os.path.join(cus_dir,cus_fn)
        # print(fn)
        df=pd.read_excel(fn,sheet_name='身体数据')
        # print(df)
        df=df[(df['时间']>=start_time) & (df['时间']<=end_time)]
        if df.empty:
            print('在输入的时间段内无训练数据或身体数据，请检查表格中的时间。')

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


        if items=='all':
            items=['wt','cal','arm','leg','waist','hip','chest']  

        y_value=0.08
        ht_fig=10
        ht_axes=0.8/len(items)
        fig=plt.figure(facecolor=bgcolor,figsize=(8,ht_fig))
             
        
        if 'wt' in items:    
            ax1=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#D9CBC6')
            ax1.plot(x,y_wt,'o-',color='#FF4747',label='体重')
            ax1.set_ylabel('体重(Kg)',fontproperties=myfont,color='#B9ACA8')
            ax1.tick_params(axis='y',colors='#FF4747')
            ax1.tick_params(axis='x',colors='#A65817')
            ax1.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax1.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax1.set_xticks([])
            # ax1.set_xticklabels(x,rotation=25)
            # ax1.legend(prop=myfont)
            ax1.set_ylim(min(y_wt)*0.98,max(y_wt)*1.02)
            for xy in list(zip(x,y_wt)):
                ax1.text(xy[0],xy[1]+0.5,xy[1],color='#FF4747')
            y_value+=1.1*ht_axes

        if 'cal' in items:
            ax2=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#F5F6FF')
            ax2.plot(x,y_r_calf,marker='s',color='#4D85A6',label='右小腿围')
            ax2.plot(x,y_l_calf,marker='s',color='violet',label='左小腿围')
            ax2.set_ylabel('小腿围(cm)',fontproperties=myfont,color='#4D85A6')
            ax2.tick_params(axis='y',colors='#4D85A6')
            ax2.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax2.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax2.set_xticks([])
            ax2.legend(prop=myfont)
            ax2.set_ylim(min(y_r_calf)*0.95,max(y_r_calf)*1.05)
            for xy in list(zip(x,y_r_calf)):
                ax2.text(xy[0],xy[1]+0.4,xy[1],color='#4D85A6')
            for xy in list(zip(x,y_l_calf)):
                ax2.text(xy[0],xy[1]-0.9,xy[1],color='violet')
            y_value+=1.1*ht_axes

        if 'leg' in items:    
            ax3=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#F5F6FF')
            ax3.plot(x,y_r_leg,marker='s',color='#4D85A6',label='右大腿围')
            ax3.plot(x,y_l_leg,marker='s',color='violet',label='左大腿围')
            ax3.set_ylabel('大腿围(cm)',fontproperties=myfont,color='#4D85A6')
            ax3.tick_params(axis='y',colors='#4D85A6')
            ax3.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax3.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax3.set_xticks([])
            ax3.legend(prop=myfont)
            ax3.set_ylim(min(y_r_leg)*0.95,max(y_r_leg)*1.05)
            for xy in list(zip(x,y_r_leg)):
                ax3.text(xy[0],xy[1]+0.4,xy[1],color='#4D85A6')
            for xy in list(zip(x,y_l_leg)):
                ax3.text(xy[0],xy[1]-1.2,xy[1],color='violet')
            y_value+=1.1*ht_axes

        if 'arm' in items:    
            ax4=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#F5F6FF')
            ax4.plot(x,y_r_arm,marker='s',color='#4D85A6',label='右臂围')
            ax4.plot(x,y_l_arm,marker='s',color='violet',label='左臂围')
            ax4.set_ylabel('臂围(cm)',fontproperties=myfont,color='#4D85A6')
            ax4.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            ax4.tick_params(axis='y',colors='#4D85A6')
            if y_value==0.08:
                ax4.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax4.set_xticks([])
            ax4.legend(prop=myfont)
            ax4.set_ylim(min(y_r_arm)*0.95,max(y_r_arm)*1.05)
            for xy in list(zip(x,y_r_arm)):
                ax4.text(xy[0],xy[1]+0.3,xy[1],color='#4D85A6')
            for xy in list(zip(x,y_l_arm)):
                ax4.text(xy[0],xy[1]-0.8,xy[1],color='violet')
            y_value+=1.1*ht_axes

        if 'waist' in items:
            ax5=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#FCFBF3')
            ax5.plot(x,y_waist,marker='s',color='#F3CC7F',label='腰围')
            ax5.set_ylabel('腰 围 (cm)',fontproperties=myfont,color='#8D8D8D')
            ax5.tick_params(axis='y',colors='#8D8D8D')
            
            # ax5.xaxis.set_major_locator(mticker.FixedLocator(x))
            # ax5.set_xticklabels(x,rotation=25)
            ax5.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax5.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax5.set_xticks([])
            # ax5.legend(prop=myfont)
            ax5.set_ylim(min(y_waist)*0.95,max(y_waist)*1.05)
            for xy in list(zip(x,y_waist)):
                ax5.text(xy[0],xy[1]+0.5,xy[1],color='#F3CC7F')
            y_value+=1.1*ht_axes

        if 'hip' in items:    
            ax6=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#F6FCF6')
            ax6.plot(x,y_hip,marker='s',color='#85B29C',label='臀围')
            ax6.set_ylabel('臀 围 (cm)',fontproperties=myfont,color='#8D8D8D')
            ax6.tick_params(axis='y',colors='#8D8D8D')
            # ax6.tick_params(axis='x',colors='#EAE8E8')
            ax6.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax6.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax6.set_xticks([])
            # ax6.legend(prop=myfont)
            ax6.set_ylim(min(y_hip)*0.95,max(y_hip)*1.05)
            for xy in list(zip(x,y_hip)):
                ax6.text(xy[0],xy[1]+0.5,xy[1],color='#85B29C')
            y_value+=1.1*ht_axes


        if 'chest' in items:
            ax7=fig.add_axes([0.1, y_value, 0.8, ht_axes],facecolor='#FFFAFE')
            ax7.plot(x,y_chest,marker='s',color='#EAC5E3',label='胸围')
            ax7.set_ylabel('胸 围 (cm)',fontproperties=myfont,color='#8D8D8D')
            ax7.tick_params(axis='y',colors='#8D8D8D')
            ax7.tick_params(axis='x',direction='in',labelrotation=40,labelsize=10,pad=5) #选择x轴
            if y_value==0.08:
                ax7.tick_params(axis='x',colors='#8D8D8D')
            if y_value!=0.08:
                ax7.set_xticks([])
            # ax4.legend(prop=myfont)
            ax7.set_ylim(min(y_chest)*0.95,max(y_chest)*1.05)
            for xy in list(zip(x,y_chest)):
                ax7.text(xy[0],xy[1]+0.5,xy[1],color='#EAC5E3')

            ax7.set_title(title,fontproperties=myfont,y=1.1,fontsize=20,color='#85B29C')
            y_value+=1.1*ht_axes

        for ax in fig.axes:
            clr='#DDDDDD'
            for bdr in ['left','right','bottom','top']:
                ax.spines[bdr].set_color(clr)

        # print(y_value)
        # plt.savefig('/home/jack/data/temp/mhdata.jpg')
        # plt.show()
        image=pic_transfer.mat_to_pil_img(fig)
        return image

class Scale:
    def __init__(self,scale_name='BMI',stage=[0,18.5,24,28,50],stage_name=['','','','超重','肥胖',''],colors=('#F4DDA4','#DFF4A4','#F4DDA4','#FBB2AB','#FE8E8E')):
        self.scale_name=scale_name
        self.colors=colors
        self.stage=stage
        self.stage_name=stage_name

    def draw(self,val=23,color_val='#84B6B9',box=(1200,600),scale_adj=10,color_bg='#FFFFFF',back_transparent_color='',arrow_fn='D:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\素材\\公共素材\\UI图标\\倒三角.png'):

        bg=Image.new('RGBA',box,color=color_bg)
        arrow=Image.open(arrow_fn)
        arrow=arrow.resize((28,arrow.size[1]*28//arrow.size[0]))
        arrow_mask=arrow.split()[3]

        #画总刻度条
        rec=Image.new('RGBA',(800,40),'#FFFFFF')
        rec=pics_fill.FillGradient().fill_multi_gradient_rct_rgb(rec,self.colors,horizontal=(True, True, True),direction='horizon')

        #画刻度
        rec_draw=ImageDraw.Draw(rec)
        bg_draw=ImageDraw.Draw(bg)
        line_max=self.stage[-1]-self.stage[0]

        line_scales=self.stage[1:-1]
        if self.scale_name.lower()=='bfr':
            txt_line_scales=[str(x)+'%' for x in line_scales]
            txt_val=str(val)+'%'
        else:
            txt_line_scales=line_scales
            txt_val=str(val)
        stg_name=self.stage_name[1:-1]

        for _n,l_scale in enumerate(line_scales):
            x=l_scale*rec.size[0]//line_max
            rec_draw.line((int(x-scale_adj),0,int(x-scale_adj),40), fill='#9F9C9C')
            bg_draw.text((int(x-scale_adj)+185,416),str(txt_line_scales[_n]),fill='#6B6565',font=ImageFont.truetype('simhei',28))
            bg_draw.text((int(x-scale_adj)+175,450),stg_name[_n],fill='#6B6565',font=ImageFont.truetype('simhei',24))
        # bg_draw.text((50,91),'44444',fill='#000000',font=ImageFont.truetype('simhei',18))
        bg.paste(rec,(200,360))

        #箭头标记输入的数值
        x_icon=val*rec.size[0]//line_max
        bg.paste(arrow,(int(x_icon-scale_adj)+200-arrow.size[0]//2,330),mask=arrow_mask)
        bg_draw.text((int(x_icon-scale_adj)+200-16,290),txt_val,fill=color_val,font=ImageFont.truetype('simhei',32))

        if back_transparent_color=='':
            pass
        else:
            bg=pic_transfer.pure_bg_transparent(img=bg,bg_color=back_transparent_color)

        return bg


if __name__=='__main__':
    data={'ht_lung':8,'balance':8,'power':5, 'flexibility':3,'core':7}
    p=DrawRadar().draw(data)
    p.show()

    # data=PeriodChart(font_fn='G:\\健身项目\\minghu\\fonts\\msyh.ttc')
    # img=data.to_pic(cus_dir='e:\\temp\\铭湖健身测试\\会员资料',cus_fn='MH003吕雅颖.xlsx',start_time='20210329',end_time='',d_font='',title='',bgcolor='#ffee00',items=['waist','hip','chest'])
    # img.show()

    # p=Scale(scale_name='BMI',stage=[10,18.5,24,28,40],stage_name=['','','超重','肥胖',''],colors=('#F4DDA4','#F4DDA4','#DFF4A4','#F4DDA4','#FBB2AB','#FE8E8E'))
    # res=p.draw(val=22.5,color_val='#84B6B9',scale_adj=200,color_bg='#FB2121',back_transparent_color='',arrow_fn='D:\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\素材\\公共素材\\UI图标\\倒三角_blue.png')
    # res.show()