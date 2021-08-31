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
    def __init__(self,place='seven'):
        config=readconfig.exp_json(os.path.join(os.path.dirname(__file__),'configs','main_'+place+'.config'))
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
        window.title(self.gym_name+' | 会员管理及反馈小程序 v1.0')
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
        after_class_menu=tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label='课后反馈生成',menu=after_class_menu)
        after_class_menu.add_cascade(label='批量',command=self.after_batch)        
        after_class_menu.add_cascade(label='个人',command=self.after_individual)        

        menubar.add_cascade(label='生成会员总结',command=self.cus_summary_menu)
        menubar.add_cascade(label='批量录入会员训练信息',command=self.gp_input_train_menu)
        menubar.add_cascade(label='生成新的会员资料表',command=self.new_cus_excel)
        fr_grp.pack()
        

        window.config(menu=menubar)
        # window.update_idletasks()
        window.mainloop()
    
    def after_batch(self):
        self.fr_destroy(fr_grp)
        self.feedback_after_class(fr_grp,group='yes')

    def after_individual(self):
        self.fr_destroy(fr_grp)
        self.feedback_after_class(fr_grp,group='no')

    def cus_summary_menu(self):
        self.fr_destroy(fr_grp)
        self.cus_summary(fr_grp)

    def gp_input_train_menu(self):
        self.fr_destroy(fr_grp)
        self.group_train_input(fr_grp)

    def new_cus_excel(self):
        self.fr_destroy(fr_grp)
        self.add_new_cus(fr_grp)

    def fr_destroy(self,fr):
        for widget in fr.winfo_children():
            widget.destroy()

    #课后生成反馈图片
    def feedback_after_class(self,window,group='yes'):
        if group=='yes':
            title='团课批量课后反馈'
        else:
            title='个人课后反馈'
        lb_title=tk.Label(window,text=title,bg='#F3D7AC',font=('黑体',13),fg='#246e4c',width=500,height=3)
        lb_title.pack()


        def open_gp_list():
            gp_list_src=os.path.join(self.cus_dir,'00-团课分班录入表.xlsx')
            os.startfile(gp_list_src)

        if group=='yes':
            txt_grp=tk.Label(window,text='请先在“00-团课分班录入表”中录入会员名单\n并保存',font=('宋体',13),fg='#246e4c',bg='#FFFFEE',width=500,padx=10,pady=20)
            txt_grp.pack()
            
            btn_open_gp_list=tk.Button(window,text='打开“00-团课分班录入表”',font=('幼圆',10),width=28,command=open_gp_list)
            btn_open_gp_list.pack(pady=10)
        
        # today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')

        lb_ins=tk.Label(window,text='选择教练',bg='#FFFFEE',font=('楷体',12),width=500,height=2)
        lb_ins.pack()

       
        ins_list=self.get_ins_list()

        ins = tk.StringVar()    # 定义一个var用来将radiobutton的值和Label的值联系在一起.
        for ins_name in ins_list:        
            ins.set(ins_list[0])
            ins1= tk.Radiobutton(window, text=ins_name[8:], variable=ins, value=ins_name)
            ins1.pack()

        if group!='yes':
            # print('individual')
            lb_cus=tk.Label(window,text='录入会员姓名（'+self.cus_instance_name+'）',bg='#FFFFEE',font=('楷体',12),width=500,height=2)
            lb_cus.pack()
            var_cus_name=tk.StringVar()
            cus_name_input=tk.Entry(window,textvariable=var_cus_name,show=None,font=('宋体', 14),width=15)
            cus_name_input.pack()

            cus_list=self.get_cus_list()
            # print(cus_list)
            LB1 = tk.Listbox(window, height=4)
            y_pdbox=169+(len(ins_list)-1)*27
            cus_listbox=menu_style.PullDownBox(LB1,cus_name_input,x=328,y=y_pdbox)
            cus_name_input.bind('<Key>', cus_listbox.handlerAdaptor(cus_listbox.text_entry_box_change,cus_list))    
            LB1.bind('<Double-Button-1>', cus_listbox.send)
            # cus_name.place(x=50, y=30)
            cus_name_input.pack(pady=10)
            def open_cus_file():
                cus_name=var_cus_name.get().upper()
                if cus_name in cus_list:
                    os.startfile(os.path.join(self.cus_dir,cus_name+'.xlsx'))
                else:
                    feed_back.delete('1.0','end')
                    feed_back.insert('insert','会员ID不在列表内，请检查。')

            btn_open_cus_file=tk.Button(window,text='打开会员文件',command=open_cus_file)
            btn_open_cus_file.pack(pady=10)

        lb_date=tk.Label(window,text='输入日期（YYYYMMDD）',bg='#FFFFEE',font=('楷体',12),width=500,height=2,pady=6)
        lb_date.pack()

        var_date=tk.StringVar()
        date_input=tk.Entry(window, textvariable=var_date,show=None, font=('宋体', 14),width=8)
        date_input.pack()

        feed_back=tk.Text(window)

        #按钮触发的函数
        def exp_feedback_after_class():
            date_txt=date_input.get()
            feed_back.delete('1.0','end')
            if len(date_txt)==8 and self.isValidDate(int(date_txt[0:4]),int(date_txt[4:6]),int(date_txt[6:])):               

                mystd = myStdout(feed_back)	# 实例化重定向类
                ac=AfterClass.FeedBack()
                if group=='yes':
                    ac.today_feedback_group(place=self.place,ins=ins.get(),date_input=date_txt,open_dir='no')
                    os.startfile(self.output_dir)
                else:
                    cus_name=var_cus_name.get().upper()
                    cus_list=self.get_cus_list()
                    if cus_name in cus_list:
                        ac.today_feedback(place=self.place,cus=cus_name,ins=ins.get(),date_input=date_txt)
                        os.startfile(os.path.join(self.output_dir,cus_name))
                    else:
                        print('会员ID不在列表内，请检查。')
                mystd.restoreStd()
            else:
                feed_back.insert('insert','日期错误：'+date_txt)
                # feed_back.pack()
                
                #在窗口界面设置放置Button按键
        
        b = tk.Button(window, text='生成课后反馈图', font=('幼圆', 8), width=18, height=2, command=exp_feedback_after_class)
        b.pack(pady=10)

        feed_back.pack() 

    # 私教会员生成训练总结图片
    def cus_summary(self,window):
        lb_title=tk.Label(window,text='生成会员总结',bg='#F3D7AC',font=('黑体',13),fg='#246e4c',width=500,height=3)
        lb_title.pack()
        lb_ins=tk.Label(window,text='选择教练',bg='#FFFFEE',font=('楷体',12),width=500,height=2)
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

        lb_cus=tk.Label(window,text='输入会员编号及姓名（'+self.cus_instance_name+'）',bg='#FFFFEE',font=('楷体',12),width=500,height=2)
        lb_cus.pack()
        var_cus_name=tk.StringVar()
        cus_name=tk.Entry(window,textvariable=var_cus_name,font=('宋体',12),width=18)
        # cus_name.pack(pady=10)

        
        cus_list=self.get_cus_list()
        # print(cus_list)
        LB1 = tk.Listbox(window, height=4)
        y_pdbox=169+(len(ins_list)-1)*27
        cus_listbox=menu_style.PullDownBox(LB1,cus_name,x=328,y=y_pdbox)
        cus_name.bind('<Key>', cus_listbox.handlerAdaptor(cus_listbox.text_entry_box_change,cus_list))    
        LB1.bind('<Double-Button-1>', cus_listbox.send)
        # cus_name.place(x=50, y=30)
        cus_name.pack(pady=10)
        
            
        # print(cus_name,cus_list)
        def open_cus_file():
            cus_name=var_cus_name.get().upper()
            if cus_name in cus_list:
                os.startfile(os.path.join(self.cus_dir,cus_name+'.xlsx'))
            else:
                feed_back.insert('insert','会员ID不在列表内，请检查。')

        btn_open_cus_file=tk.Button(window,text='打开会员文件',command=open_cus_file)
        btn_open_cus_file.pack(pady=10)

        var_date_start=tk.StringVar()
        var_date_end=tk.StringVar()
        date_start=tk.Entry(window,textvariable=var_date_start,font=('宋体',12),width=8)
        date_end=tk.Entry(window,textvariable=var_date_end,font=('宋体',12),width=8)
        date_start_title=tk.Label(window,text='输入起始日期',bg='#ffffee',font=('楷体',13),width=500,height=2)
        date_end_title=tk.Label(window,text='输入结束日期',bg='#ffffee',font=('楷体',13),width=500,height=2)
        date_start_title.pack()
        date_start.pack()
        date_end_title.pack()
        date_end.pack()


        

        def exp_cus_summary():
            date_s=date_start.get()
            date_e=date_end.get()
            if len(date_s)==8 and len(date_e)==8 and self.isValidDate(int(date_s[0:4]),int(date_s[4:6]),int(date_s[6:])) and \
                                self.isValidDate(int(date_e[0:4]),int(date_e[4:6]),int(date_e[6:])) and \
                                calculate_days_2(date_s,date_e)>=0:               

                mystd = myStdout(feed_back)	# 实例化重定向类                
                # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')
                cus_name=var_cus_name.get().upper()
                if cus_name in cus_list:
                    feed_back.delete('1.0','end')
                    print('正在生成会员训练总结')
                    run.cus_feedback(place=self.place,cus=cus_name,ins=ins.get(),start_time=date_s,end_time=date_e,adj_bfr='yes',adj_src='gui',gui=window)
                else:
                    feed_back.delete('1.0','end')
                    print('会员ID不在列表内，请检查。')
                mystd.restoreStd()
            else:
                feed_back.delete('1.0','end')
                feed_back.insert('insert','日期错误：'+date_s+','+date_e)
        btn=tk.Button(window,text='生成会员训练总结',font=('幼圆',12),width=18,command=exp_cus_summary)
        btn.pack(pady=10)
        feed_back.pack()

    #批量录入团课训练信息
    def group_train_input(self,window):
        txt_grp=tk.Label(window,text='请先在“00-团课分班录入表”中录入会员资料\n并保存',font=('黑体',13),fg='#246e4c',bg='#f3d7ac',padx=10,pady=20)
        txt_grp.pack()
        feed_back_gp_input=tk.Text(window)
        def gp_input_train():            
            feed_back_gp_input.delete('1.0','end')
            fd_screen=myStdout(feed_back_gp_input)
            run.group_input(place=self.place)
            fd_screen.restoreStd()     

        def open_gp_list():
                gp_list_src=os.path.join(self.cus_dir,'00-团课分班录入表.xlsx')
                os.startfile(gp_list_src)
            
        btn_open_gp_list=tk.Button(window,text='打开“00-团课分班录入表”',font=('幼圆',10),width=28,command=open_gp_list)
        btn_open_gp_list.pack(pady=10)   
        
        btn_gp_input=tk.Button(window,text='点击开始\n批量录入训练信息',font=('幼圆',12),width=18,command=gp_input_train)
        btn_gp_input.pack()
        feed_back_gp_input.pack(pady=10)

    def add_new_cus(self,window):
            lb_title=tk.Label(window,text='生成新的会员表',bg='#F3D7AC',font=('黑体',13),fg='#246e4c',width=500,height=3)
            lb_title.pack() 
            lb_cus=tk.Label(window,text='请输入新的会员姓名',bg='#FFFFEE',font=('楷体',12),width=500,height=2)
            lb_cus.pack()
            value_cus_name=tk.StringVar()
            cus_name_input=tk.Entry(window,textvariable=value_cus_name,font=('宋体',12),width=8)
            cus_name_input.pack()
            msg_box=tk.Text(window)

            def new_cus():
                msg_box.delete('1.0','end')
                if value_cus_name.get():
                    new_msg=myStdout(msg_box)
                    new_cus_fn=run.auto_xls(place=self.place,cus_name_input=value_cus_name.get().upper(),mode='gui',gui=msg_box)
                    print(new_cus_fn)
                    if new_cus_fn:
                        print('正在打开文件……')
                        os.startfile(os.path.join(self.cus_dir,new_cus_fn+'.xlsx'))
                    new_msg.restoreStd()                    
                else:
                    msg_box.insert('insert','请输入姓名')            

            btn_add_new=tk.Button(window,text='新增会员',font=('幼圆',12),width=18,command=new_cus)
            btn_add_new.pack(pady=10)
            msg_box.pack()
            
    def get_ins_list(self):        
        df_ins_list=self.df_ins['员工编号']+self.df_ins['姓名']
        df_ins_list.dropna(how=any,axis=0,inplace=True)
        ins_list=df_ins_list.tolist()
        return ins_list

    def isValidDate(self,year, month, day):
        try:
            date(year, month, day)
        except:
            return False
        else:
            return True

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