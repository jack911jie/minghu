import os
import sys
sys.path.extend([os.path.join(os.path.dirname(os.path.dirname(__file__)),'data_analysis'),os.path.join(os.path.dirname(os.path.dirname(__file__)),'modules')])
import readconfig
import cus_data
import re
import xlwings as xw
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
from flask import Flask, request, jsonify,render_template

class MinghuService(Flask):
    def __init__(self,*args,**kwargs):
        super(MinghuService, self).__init__(*args, **kwargs)
        config_fn=os.path.join(os.path.join(os.path.dirname(__file__),'config','minghu_service.config'))
        self.config_mh=readconfig.exp_json2(config_fn)

        #路由
        self.add_url_rule('/',view_func=self.index)
        self.add_url_rule('/welcome',view_func=self.welcome)
        self.add_url_rule('/get_cus_list', view_func=self.get_cus_list,methods=['GET','POST'])
        self.add_url_rule('/get_cus_info', view_func=self.get_cus_info,methods=['GET','POST'])
        self.add_url_rule('/open_cus_fn', view_func=self.open_cus_fn,methods=['GET','POST'])
        self.add_url_rule('/check_new', view_func=self.check_new,methods=['GET','POST'])
        self.add_url_rule('/generate_new', view_func=self.generate_new,methods=['GET','POST'])
        self.add_url_rule('/new_cus', view_func=self.new_cus,methods=['GET','POST'])

    

    def wecom_dir(self):
        # fn=os.path.join(os.path.dirname(__file__),'config','wecom_dir.config')        
        res=os.path.join(self.config_mh['work_dir'].strip(),'01-会员管理','会员资料')
        return res

    def cus_list(self):
        dic_li=[]
        for fn in os.listdir('D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
            if re.match(r'^MH\d{3}.*.xlsm$',fn):
                dic_li.append(fn.split('.')[0])
        return dic_li

    #遍历会员资料生成名字
    def get_cus_list(self):
        dic_li=self.cus_list()
        return jsonify(dic_li)


    # 定义前端页面路由
   
    def index(self):
        return render_template('index.html')


    def new_cus(self):
        return render_template('new_cus.html')


    def get_cus_info(self):
        cus_name = request.json.get('selected_name')
        work_dir=self.wecom_dir()
        fn=os.path.join(work_dir,cus_name+'.xlsm')
        p=cus_data.CusData()
        res=p.cus_cls_rec_toweb(fn=fn,cls_types=self.config_mh['all_cls_types'],not_lmt_types=self.config_mh['not_lmt_cls_types'])
        res.fillna(0)
        data=res.iloc[0].to_dict()
        return jsonify(data)


    def open_cus_fn(self):
        cus_name=request.data.decode('utf-8')
        cus_li=self.cus_list()
        if cus_name and cus_name in cus_li:
            work_dir=self.wecom_dir()
            fn=os.path.join(work_dir,cus_name+'.xlsm')
            # os.startfile(fn)
            return f'正在打开 {cus_name} 的会员档案'
        else:
            return '会员编码及编码为空/无此会员档案'


    def check_new(self):
        dat=request.data
        cus_li=self.cus_list()
        cus_num=[int(x[2:5]) for x in cus_li]
        max_num=max(cus_num)
        new_num=max_num+1
        txt_num=str(new_num).zfill(3)
        # new_name='MH'+new_num.zfill(3)+cus_name+'.xlsm'
        # new_name=os.path.join(wecom_dir,new_name)
        return txt_num


    def generate_new(self):
        try:
            fn_in=request.data
            fn='MH'+fn_in.decode('utf-8')
            fn,dvc=fn.split('|')
            work_dir=self.wecom_dir()
            tplt_dir=os.path.dirname(work_dir)
            new_fn=os.path.join(work_dir,fn+'.xlsm')

            app=xw.App(visible=False)
            wb=app.books.open(os.path.join(tplt_dir,'模板.xlsm'))
            sht=wb.sheets['基本情况']
            sht['A2'].value=fn[0:5]
            sht['B2'].value=fn[5:]
            if len(fn[5:])>1:
                sht['C2'].value=fn[5:][1:]
            else:
                sht['C2'].value=fn[5:]

            wb.save(new_fn)
            wb.close()
            app.quit()

            # os.startfile(work_dir)
            if dvc=='pc':
                os.startfile(new_fn)

            return new_fn
        except Exception as e:
            return e


    def welcome(self):
        return '关于我们页面'

if __name__ == '__main__':
    app = MinghuService(__name__)
    # app.run(debug=True)
    app.run(debug=True,host='192.168.1.38',port=5000)
    # res=wecom_dir()
    # print(res)
