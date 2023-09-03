import os
import sys
import numpy as np
from tkinter.constants import CENTER
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import readconfig
import menu_style
from days_cal import calculate_days_2
import run
import AfterClass
import pandas as pd
import tkinter as tk
from datetime import date
from PIL import Image,ImageTk
import re
import warnings

warnings.filterwarnings('ignore')

class GUI:
    def __init__(self,place='seven',wecomid_replace='yes',wecomid_pair=['$wecomid$','1688850049985213']):
        config=readconfig.exp_json(os.path.join(os.path.dirname(__file__),'configs','main_'+place+'.config'),
                                    wecomid_replace=wecomid_replace,wecomid_pair=wecomid_pair)
        self.cus_dir=config['会员档案文件夹']
        self.material_dir=config['素材文件夹']
        self.public_dir=config['公共素材文件夹']
        self.ins_dir=config['教练文件夹']
        self.output_dir=config['课后反馈文件夹']
        self.df_ins=pd.read_excel(os.path.join(self.ins_dir,'教练信息.xlsx'),sheet_name='教练信息')
        self.place=place
        with open(os.path.join(self.material_dir,'txt_public.txt'),'r',encoding='utf-8') as txt_pub:
            self.txt_public=txt_pub.readlines()
        self.cus_instance_name=self.txt_public[5].strip()
        self.prefix=self.cus_instance_name[0:2]
        self.gym_name=self.txt_public[4]
        self.gym_addr=self.txt_public[3]
        if '%' in self.gym_addr:
            self.gym_addr=''
        self.txt_ins_word=self.txt_public[2]
        self.txt_mini_title=self.txt_public[1]
        self.txt_slogan=self.txt_public[0]

        
    def creat_gui(self):
        global fr_grp
        window =tk.Tk()
        window.title(self.gym_name+' | 教练工作小程序 v1.1')
        window.geometry('500x600')
        window.attributes("-toolwindow", 2)
        window.resizable(0,0)

        fr_grp=tk.Frame(window)

        cover_im=Image.open(os.path.join(self.public_dir,'logo及二维码','logo.jpg'))
        cover_im=cover_im.resize((200,200))
        cover_img=ImageTk.PhotoImage(cover_im)
        cover_label=tk.Label(fr_grp,image=cover_img)
        # cover_label.place(x=150,y=200,width=300,height=300,anchor=CENTER)
        cover_label.pack(pady=30)
        cover_txt=tk.Label(fr_grp,text=self.txt_slogan,font=('幼圆',18),fg='#665B2C')
        cover_txt.pack()


        menubar=tk.Menu(window)
        pre_class_menu=tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label='客户预约文本',menu=pre_class_menu)
        pre_class_menu.add_cascade(label='录入信息',command=self.after_batch)        

        fr_grp.pack()
        

        window.config(menu=menubar)
        # window.update_idletasks()
        window.mainloop()
    
    def after_batch(self):
        self.fr_destroy(fr_grp)
        self.cus_reservation(fr_grp)



    def fr_destroy(self,fr):
        for widget in fr.winfo_children():
            widget.destroy()

    #生成预约文本
    def cus_reservation(self,window):
        lb_title=tk.Label(window,text='生成并发送预约信息',bg='#F8FAF5',font=('黑体',13),fg='#455233',width=500,height=3)
        lb_title.pack()
        lb_ins=tk.Label(window,text='选择教练',bg='#DCF7FC',font=('楷体',12),width=500,height=2)
        lb_ins.pack()

        feed_back=tk.Text(window)
        # today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')
        
        ins_list=self.get_ins_list()

        ins = tk.StringVar()    # 定义一个var用来将radiobutton的值和Label的值联系在一起.
        for ins_name in ins_list:        
            if ins_name!=np.nan:   
                ins.set(ins_list[0])
                ins1= tk.Radiobutton(window, text=ins_name[8:], variable=ins, value=ins_name)
                ins1.pack()

        lb_cus=tk.Label(window,text='输入会员编号及姓名（'+self.cus_instance_name+'）',bg='#DCF7FC',font=('楷体',12),width=500,height=2)
        lb_cus.pack()
        var_cus_name=tk.StringVar()
        cus_name=tk.Entry(window,textvariable=var_cus_name,font=('宋体',12),width=18)
        # cus_name.pack(pady=10)

        
        cus_list=self.get_cus_list()
        # print(cus_list)
        LB1 = tk.Listbox(window, height=5)
        y_pdbox=169+(len(ins_list)-1)*27
        cus_listbox=menu_style.PullDownBox(LB1,cus_name,x=328,y=y_pdbox)
        cus_name.bind('<Key>', cus_listbox.handlerAdaptor(cus_listbox.text_entry_box_change,cus_list))    
        LB1.bind('<Double-Button-1>', cus_listbox.send)
        # cus_name.place(x=50, y=30)
        cus_name.pack(pady=10)
        cus_gap=tk.Label(window,text=' ' ,bg='#f0f0f0',font=('楷体',12),width=500,height=2)
        cus_gap.pack()
        # cus_gap.pack()


        var_crs_type=tk.StringVar()
        var_crs_type.set('常规私教课')
        crs_type_title=tk.Label(window,text='课程类型',bg='#DCF7FC',font=('楷体',12),width=500,height=2)
        crs_type_title.pack()
        crs_type= tk.Radiobutton(window, text='常规私教课', variable=var_crs_type, value='常规私教课')
        crs_type.pack()
        
            
        # print(cus_name,cus_list)
        def open_cus_file():
            cus_name=var_cus_name.get().upper()
            if cus_name in cus_list:
                os.startfile(os.path.join(self.cus_dir,cus_name+'.xlsx'))
            else:
                feed_back.insert('insert','会员ID不在列表内，请检查。')

        # btn_open_cus_file=tk.Button(window,text='打开会员文件',command=open_cus_file)
        # btn_open_cus_file.pack(pady=10)

        var_date_start=tk.StringVar()
        var_time_prd=tk.StringVar()

        date_start=tk.Entry(window,textvariable=var_date_start,font=('宋体',12),width=8)

        date_start_title=tk.Label(window,text='输入上课日期',bg='#DCF7FC',font=('楷体',13),width=500,height=2)
        date_start_title.pack()
        date_start.pack()
  
        time_prd=tk.Entry(window,textvariable=var_time_prd,font=('宋体',12),width=9)        
        time_prd_title=tk.Label(window,text='输入上课时间（如：1000-1100）',bg='#DCF7FC',font=('楷体',13),width=500,height=2)
        time_prd_title.pack()
        time_prd.pack()



        

        def exp_cus_rsv_txt():
            date_s=date_start.get()
            # date_e=date_end.get()
            time_period=time_prd.get()
            ins_name=ins.get()
            crs_type_name=var_crs_type.get()
            cus_name=var_cus_name.get()

            # print(cus_name,crs_type,date_s,time_period,ins_name)

            if len(date_s)==8 and self.isValidDate(int(date_s[:4]), int(date_s[4:6]), int(date_s[6:])) and re.match(r'\d{4}-\d{4}',time_period) and self.isValidTimePrd(time_period):               

                mystd = myStdout(feed_back)	# 实例化重定向类                
                # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')
                cus_name=var_cus_name.get().upper()
                if cus_name in cus_list:
                    feed_back.delete('1.0','end')
                              # run.cus_feedback(place=self.place,cus=cus_name,ins=ins.get(),start_time=date_s,end_time=date_e,adj_bfr='yes',adj_src='gui',gui=window)
                    # 根据不同的场所设置小结logo的高度
                    if self.place=='minghu':
                        logo_ht=52
                    elif self.place=='seven':
                        logo_ht=72
                    else:
                        logo_ht=52
                        
                    print('正在生成及发送信息\n')
                    run.wecom_send(place=self.place,work_dir='D:\\Documents\\WXWork\\1688851376196754\\WeDrive\\铭湖健身工作室',
                                    cus_name=cus_name,crs_type=crs_type_name,crs_date=date_s,crs_time=time_period,ins=ins_name)

                else:
                    feed_back.delete('1.0','end')
                    print('会员ID不在列表内，请检查。')
                mystd.restoreStd()
            else:
                feed_back.delete('1.0','end')
                feed_back.insert('insert','日期或时间错误：'+date_s+','+time_period)


        btn=tk.Button(window,text='发送预约信息到群',font=('幼圆',12),width=18,command=exp_cus_rsv_txt)
        btn.pack(pady=10)
        feed_back.pack()

    def get_ins_list(self):        
        df_ins_list=self.df_ins['员工编号']+self.df_ins['姓名']
        df_ins_list.dropna(how=any,axis=0,inplace=True)
        ins_list=df_ins_list.tolist()
        return ins_list

    def isValidDate(self,year, month, day):
        # print(year,month,day)
        try:
            date(year, month, day)
        except:
            return False
        else:
            return True

    def isValidTimePrd(self,timeprd):
        # print(int(timeprd[:2]),int(timeprd[2:4]),int(timeprd[5:7]),int(timeprd[7:9]))
        if int(timeprd[:2])<=24 and int(timeprd[2:4])<=59 and int(timeprd[5:7])<=24 and int(timeprd[7:9])<=59:
            return True
        else:
            return False

    def get_cus_list(self):
        cus_list=[]
        for fn in os.listdir(os.path.join(self.cus_dir)):
            if re.match(self.prefix+'.*.xlsx$',fn):
               cus_list.append(fn[0:-5]) 
        return cus_list

class myStdout():	# 重定向类
    def __init__(self,t):
        self.t=t
    	# 将其备份
        self.stdoutbak = sys.stdout		
        self.stderrbak = sys.stderr
        # 重定向
        sys.stdout = self
        sys.stderr = self

    def write(self, info):
        # info信息即标准输出sys.stdout和sys.stderr接收到的输出信息
        self.t.insert('end', info)	# 在多行文本控件最后一行插入print信息
        self.t.update()	# 更新显示的文本，不加这句插入的信息无法显示
        self.t.see(tk.END)	# 始终显示最后一行，不加这句，当文本溢出控件最后一行时，不会自动显示最后一行

    def restoreStd(self):
        # 恢复标准输出
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak

if __name__=='__main__':
    minghu_gui=GUI(place='minghu')
    minghu_gui.creat_gui()
    # minghu_gui.get_cus_list()
    # print(minghu_gui.isValidTimePrd('2000-2100'))
