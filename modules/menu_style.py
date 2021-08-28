import tkinter as tk
import re

class PullDownBox:
    def __init__(self,pull_down_box,text_entry_box,x=50,y=50):
        self.text_entry_box=text_entry_box
        self.pull_down_box=pull_down_box
        self.x,self.y=x,y

# 输入框文字改变事件监听
    def text_entry_box_change(self,event,result):
        current = self.text_entry_box.get()
        current=current.upper()
        # print(current)

        #event.char 获取最近一次键盘事件输入的值
        current_value = event.char
        # \x08就是 删除按钮
        if current_value == '\x08':
            current = current[:-1]
        else:
            current += current_value
        # print('current 2=',current)
        if current.strip()!= "":
            self.handlerResult(current,result)
        else:
            # place_forget隐藏控件
            self.pull_down_box.place_forget()
            # self.pull_down_box.delete(0,tk.END)

# 生成与输入文字相关的列表
    def handlerResult(self,current,result):
        # print("-----------------------------")
        # print(result)
        self.pull_down_box.delete(0, tk.END)
        # print(result)
        for res in result:
            # 总数据中包含输入框中的内容
            if len(re.findall(current,res))>0:
                    self.pull_down_box.insert(tk.END,res)
        self.pull_down_box.place(x=self.x, y=self.y)

        #创建Scrollbar
        # yscrollbar = tk.Scrollbar(self.pull_down_box,command=self.pull_down_box.yview)
        # yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # self.pull_down_box.config(yscrollcommand=yscrollbar.set)

    # 鼠标双击列表时，将点击文字放入输入框
    # event参数是必须的
    def send(self,event):
        # 将双击列表后的文字显示到输入框上
        self.text_entry_box.delete(0, tk.END)
        self.text_entry_box.insert(0, str(self.pull_down_box.get(self.pull_down_box.curselection())))
        # print(self.pull_down_box.get(self.pull_down_box.curselection()))

        # 隐藏列表
        self.pull_down_box.delete(0, tk.END)
        self.pull_down_box.place_forget()
        


    def handlerAdaptor(self,fun, res):
        '''事件处理函数的适配器，相当于中介，那个event是从那里来的呢，我也纳闷，这也许就是python的伟大之处吧'''
        return lambda event, fun=fun, kwds=res: fun(event, kwds)




if __name__=='__main__':
    root = tk.Tk()
    # 设置窗体标题
    root.title('快递价格')
    # 设置窗口大小和位置
    root.geometry('1000x400+570+200')
    label = tk.Label(root, text='发出地:')
    text_entry_box = tk.Entry(root, bg='white', width=20)
    # 实时监听内容发生变化 **results**为我们的总数据
    results=('mh001刘大','mh002李大','mh003焦大','mh024王大')
    #下拉菜单
    pull_down_box = tk.Listbox(root, height=8)
    cus_lists=PullDownBox(pull_down_box,text_entry_box)
    '''
    一般Tkinter事件绑定函数是不带参数的（bind会默认带event事件参数）
    所以，使用bind的时候，event也是一个参数，所以将事件、参数绑定起来执行有两种做法：    
    一是通过中介函数handlerAdaptor，将要真正执行的函数、事件（即event）和参数“捆绑”在一起执行，
    二是通过lambda函数将事件（event）和参数传入真正执行的函数中。
    '''
    # text_entry_box.bind('<Key>', cus_lists.handlerAdaptor(cus_lists.text_entry_box_change,results))    
    text_entry_box.bind('<Key>', lambda event:cus_lists.text_entry_box_change(event,results))
    #鼠标按钮n被双击，1为左键，2中键，3右键
    pull_down_box.bind('<Double-Button-1>', cus_lists.send)
    label.place(x=5, y=30)
    text_entry_box.place(x=50, y=30)
    root.mainloop()
