import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import readconfig
import run
import AfterClass
import tkinter as tk
from datetime import date
import re

class GUI:
    def __init__(self):
        config=readconfig.exp_json(os.path.join(os.path.dirname(__file__),'configs','main.config'))
        self.cus_dir=config['会员档案文件夹']
        
    def creat_gui(self):
        global window,fr_grp
        window =tk.Tk()
        window.title('铭湖健身会员课后反馈')
        window.geometry('500x600')

        # main_title=tk.Text(window,width=100,height=5)
        # main_title.insert('insert','铭湖健身')
        # main_title.pack()

        fr_grp=tk.Frame(window)

        menubar=tk.Menu(window)
        after_class_menu=tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label='课后录入',menu=after_class_menu)
        after_class_menu.add_cascade(label='批量录入',command=self.after_batch)        
        after_class_menu.add_cascade(label='个人录入',command=self.after_individual)        

        menubar.add_cascade(label='生成会员总结',command=self.cus_summary_menu)
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

    def fr_destroy(self,fr):
        for widget in fr.winfo_children():
            widget.destroy()

    #课后生成反馈图片
    def feedback_after_class(self,window,group='yes'):
        if group=='yes':
            title='团课批量录入'
        else:
            title='个人录入'
        lb_title=tk.Label(window,text=title,bg='#F3D7AC',font=('幼圆',13),width=500,height=3)
        lb_title.pack()
        lb_ins=tk.Label(window,text='选择教练',bg='#FFFFEE',font=('黑体',12),width=500,height=2)
        lb_ins.pack()

        # today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')
        
        ins = tk.StringVar()    # 定义一个var用来将radiobutton的值和Label的值联系在一起.
        ins.set('MHINS001陆伟杰')
        ins1= tk.Radiobutton(window, text='陆伟杰', variable=ins, value='MHINS001陆伟杰')
        ins1.pack()
        ins2 = tk.Radiobutton(window, text='韦越棋', variable=ins, value='MHINS002韦越棋')
        ins2.pack()

        if group!='yes':
            # print('individual')
            lb_cus=tk.Label(window,text='录入会员姓名（MH000李铭湖）',bg='#FFFFEE',font=('黑体',12),width=500,height=2)
            lb_cus.pack()
            var_cus_name=tk.StringVar()
            cus_name_input=tk.Entry(window,textvariable=var_cus_name,show=None,font=('楷体', 14),width=15)
            cus_name_input.pack()

        lb_date=tk.Label(window,text='输入日期（YYYYMMDD）',bg='#FFFFEE',font=('黑体',12),width=500,height=2,pady=6)
        lb_date.pack()

        var_date=tk.StringVar()
        date_input=tk.Entry(window, textvariable=var_date,show=None, font=('Arial', 14),width=8)
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
                    ac.today_feedback_group(ins=ins.get(),date_input=date_txt,open_dir='no')
                else:
                    cus_name=var_cus_name.get()
                    cus_list=self.get_cus_list()
                    if cus_name in cus_list:
                        ac.today_feedback(cus=cus_name,ins=ins.get(),date_input=date_txt)
                    else:
                        print('会员ID不在列表内，请检查。')
                mystd.restoreStd()
            else:
                feed_back.insert('insert','日期错误：'+date_txt)
                # feed_back.pack()
                
                #在窗口界面设置放置Button按键
        
        b = tk.Button(window, text='生成课后反馈图', font=('黑体', 8), width=18, height=2, command=exp_feedback_after_class)
        b.pack()

        feed_back.pack() 

    # 私教会员生成训练总结图片
    def cus_summary(self,window):
        lb_title=tk.Label(window,text='生成会员总结',bg='#F3D7AC',font=('幼圆',13),width=500,height=3)
        lb_title.pack()
        lb_ins=tk.Label(window,text='选择教练',bg='#FFFFEE',font=('黑体',12),width=500,height=2)
        lb_ins.pack()

        # today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')
        
        ins = tk.StringVar()    # 定义一个var用来将radiobutton的值和Label的值联系在一起.
        ins.set('MHINS001陆伟杰')
        ins1= tk.Radiobutton(window, text='陆伟杰', variable=ins, value='MHINS001陆伟杰')
        ins1.pack()
        ins2 = tk.Radiobutton(window, text='韦越棋', variable=ins, value='MHINS002韦越棋')
        ins2.pack()

        lb_cus=tk.Label(window,text='录入会员姓名（MH000李铭湖）',bg='#FFFFEE',font=('黑体',12),width=500,height=2)
        lb_cus.pack()
        var_cus_name=tk.StringVar()
        cus_name=tk.Entry(window,textvariable=var_cus_name,font=('楷体',12),width=18)
        cus_name.pack()

        var_date_start=tk.StringVar()
        var_date_end=tk.StringVar()
        date_start=tk.Entry(window,textvariable=var_date_start,font=('黑体',12),width=8)
        date_end=tk.Entry(window,textvariable=var_date_end,font=('黑体',12),width=8)
        date_start_title=tk.Label(window,text='输入起始日期',bg='#ffffee',font=('幼圆',13),width=500,height=2)
        date_end_title=tk.Label(window,text='输入结束日期',bg='#ffffee',font=('幼圆',13),width=500,height=2)
        date_start_title.pack()
        date_start.pack()
        date_end_title.pack()
        date_end.pack()


        feed_back=tk.Text(window)

        def exp_cus_summary():
            date_s=date_start.get()
            date_e=date_end.get()
            if len(date_s)==8 and len(date_e)==8 and self.isValidDate(int(date_s[0:4]),int(date_s[4:6]),int(date_s[6:])) and \
                                self.isValidDate(int(date_e[0:4]),int(date_e[4:6]),int(date_e[6:])):               

                mystd = myStdout(feed_back)	# 实例化重定向类                
                # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')
                cus_name=var_cus_name.get()
                cus_list=self.get_cus_list()
                if cus_name in cus_list:
                    print('正在生成会员训练总结')
                    run.cus_feedback(cus=cus_name,ins=ins.get(),start_time=date_s,end_time=date_e,adj_bfr='yes',adj_src='gui',gui=window)
                else:
                    print('会员ID不在列表内，请检查。')
                mystd.restoreStd()
            else:
                feed_back.insert('insert','日期错误：'+date_s+','+date_e)
        btn=tk.Button(window,text='生成会员训练总结',font=('黑体',12),width=18,command=exp_cus_summary)
        btn.pack()
        feed_back.pack()


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
            if re.match(r'MH.*.xlsx$',fn):
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
    minghu_gui=GUI()
    minghu_gui.creat_gui()
    # minghu_gui.get_cus_list()