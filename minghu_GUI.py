import os
import sys
import AfterClass
import tkinter as tk
from datetime import date

class GUI:
    def __init__(self):
        pass

    def group_after_class(self):
        
        window =tk.Tk()
        window.title('铭湖健身团课反馈')
        window.geometry('500x550')
        lb_ins=tk.Label(window,text='选择教练',bg='#FFFFEE',font=('黑体',12),width=500,height=2)
        lb_ins.pack()

        # today_feedback(cus='MH024刘婵桢',ins='MHINS001陆伟杰',date_input='20210623')
        
        ins = tk.StringVar()    # 定义一个var用来将radiobutton的值和Label的值联系在一起.
        ins.set('MHINS001陆伟杰')
        ins1= tk.Radiobutton(window, text='陆伟杰', variable=ins, value='MHINS001陆伟杰')
        ins1.pack()
        ins2 = tk.Radiobutton(window, text='韦越棋', variable=ins, value='MHINS002韦越棋')
        ins2.pack()

        lb_date=tk.Label(window,text='输入日期（YYYYMMDD）',bg='#FFFFEE',font=('黑体',12),width=500,height=2,pady=6)
        lb_date.pack()

        var_date=tk.StringVar()
        date_input=tk.Entry(window, textvariable=var_date,show=None, font=('Arial', 14),width=8)
        date_input.pack()

        feed_back=tk.Text(window)
        def hit_me():
            # afterclass.today_feedback(cus='MH024刘婵桢',ins=ins.get(),date_input=date_input.get())
            # afterclass.today_feedback(cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210324')
            date_txt=date_input.get()
            feed_back.delete('1.0','end')
            if len(date_txt)==8 and self.isValidDate(int(date_txt[0:4]),int(date_txt[4:6]),int(date_txt[6:])):               

                mystd = myStdout(feed_back)	# 实例化重定向类
                ac=AfterClass.FeedBack()
                ac.today_feedback_group(ins=ins.get(),date_input=date_txt,open_dir='no')

                mystd.restoreStd()
            else:
                feed_back.insert('insert','日期错误：'+date_txt)
                # feed_back.pack()
            
                #在窗口界面设置放置Button按键
        b = tk.Button(window, text='批量产生\n课后反馈图', font=('黑体', 8), width=10, height=2, command=hit_me)
        b.pack()

        feed_back.pack() 
        window.update()
        window.mainloop()
        
    def isValidDate(self,year, month, day):
        try:
            date(year, month, day)
        except:
            return False
        else:
            return True

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
    minghu_gui.group_after_class()