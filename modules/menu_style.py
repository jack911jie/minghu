
import tkinter as tk
import re



class CusLists:
    def __init__(self,LB1,text1,x=50,y=50):
        self.LB1=LB1
        self.text1=text1
        self.x,self.y=x,y
# 发出地 文字改变事件监听
    def text_change(self,event,result):
        current = self.text1.get()
        current_value = event.char
        # \x08就是 删除按钮
        if current_value == '\x08':
            current = current[:-1]
        else:
            current += current_value
        if current.strip()!= "":
            self.handlerResult(current,result)
        else:
            # place_forget隐藏控件
            self.LB1.place_forget()

    def handlerResult(self,current,result):
    # LB1.place()
        # print("-----------------------------")
        # print(result)
        self.LB1.delete(0, tk.END)
        for res in result[0]:
            # 总数据中包含输入框中的内容   这里根据你的总数据自己设置
            # if current in i[0]:
            # for res_ in res:
            #     print(res_)
            if len(re.findall(current,res))>0:
                    # LB1.insert(END,i[0]+","+i[1]+","+i[2]+","+i[3])
                    self.LB1.insert(tk.END,res)
        self.LB1.place(x=self.x, y=self.y)
            # self.LB1.pack()

        #创建Scrollbar
        # yscrollbar = tk.Scrollbar(self.LB1,command=self.LB1.yview)
        # yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # self.LB1.config(yscrollcommand=yscrollbar.set)


    def send(self,event):
        # 将双击列表后的文字显示到输入框上
        self.text1.delete(0, tk.END)
        self.text1.insert(0, str(self.LB1.get(self.LB1.curselection())))
        # print(self.LB1.get(self.LB1.curselection()))
        # 隐藏列表
        self.LB1.delete(0, tk.END)
        self.LB1.place_forget()
        


    def handlerAdaptor(self,fun, *args):
        '''事件处理函数的适配器，相当于中介，那个event是从那里来的呢，我也纳闷，这也许就是python的伟大之处吧'''
        return lambda event, fun=fun, kwds=args: fun(event, kwds)




# root = tk.Tk()
# # 设置窗体标题
# root.title('快递价格')
# # 设置窗口大小和位置
# root.geometry('1000x400+570+200')
# label1 = tk.Label(root, text='发出地:')
# text1 = tk.Entry(root, bg='white', width=20)
# # 实时监听内容发生变化 **results**为我们的总数据
# results=('mh001刘大','mh002李大','mh003焦大','mh024王大')
# text1.bind('<Key>', handlerAdaptor(text_change,results))
# # 列表
# LB1 = tk.Listbox(root, height=7)
# LB1.bind('<Double-Button-1>', send)
# label1.place(x=5, y=30)
# text1.place(x=50, y=30)
# root.mainloop()

if __name__=='__main__':
    root = tk.Tk()
    # 设置窗体标题
    root.title('快递价格')
    # 设置窗口大小和位置
    root.geometry('1000x400+570+200')
    label1 = tk.Label(root, text='发出地:')
    text1 = tk.Entry(root, bg='white', width=20)
    # 实时监听内容发生变化 **results**为我们的总数据
    results=('mh001刘大','mh002李大','mh003焦大','mh024王大','mh010刘大小','mh011李大小','mh012黄大小')
    #下拉菜单
    LB1 = tk.Listbox(root, height=4)
    cus_lists=CusLists(LB1,text1)
    text1.bind('<Key>', cus_lists.handlerAdaptor(cus_lists.text_change,results))    
    LB1.bind('<Double-Button-1>', cus_lists.send)
    label1.place(x=5, y=30)
    text1.place(x=50, y=30)
    root.mainloop()

