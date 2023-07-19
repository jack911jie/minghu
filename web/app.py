import os
import sys
sys.path.extend([os.path.join(os.path.dirname(os.path.dirname(__file__)),'data_analysis')])
import cus_data
import re
import xlwings as xw
import pandas as pd
pd.set_option('display.unicode.east_asian_width', True) #设置输出右对齐
# pd.set_option('display.max_columns', None) #显示所有列
from flask import Flask, request, jsonify,render_template


app = Flask(__name__)

def wecom_dir():
    fn=os.path.join(os.path.dirname(__file__),'config','wecom_dir.config')
    with open(fn,'r',encoding='utf-8') as f:
        lines=f.readlines()
    res=os.path.join(lines[0].strip(),'01-会员管理','会员资料')
    return res

def cus_list():
    dic_li=[]
    for fn in os.listdir('D:\\Documents\\WXWork\\1688851376239499\\WeDrive\\铭湖健身工作室\\01-会员管理\\会员资料'):
        if re.match(r'^MH\d{3}.*.xlsm$',fn):
            dic_li.append(fn.split('.')[0])
    return dic_li

#遍历会员资料生成名字
@app.route('/get_cus_list')
def get_cus_list():
    dic_li=cus_list()
    return jsonify(dic_li)


# 定义前端页面路由
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/new_cus')
def new_cus():
    return render_template('new_cus.html')

@app.route('/get_info', methods=['GET','POST'])
def get_info():
    # 从前端获取选择的姓名
    selected_name = request.json.get('selected_name')
    name_data=cus_list()

    # 在姓名数据中查找匹配的信息
    if selected_name in name_data:
        info = "姓名：{}，信息：这是{}的信息。".format(selected_name, selected_name)
    else:
        info = "姓名：{}，信息：没有找到匹配的信息。".format(selected_name)

    # 将匹配的信息返回给前端
    return jsonify({"info": info})

@app.route('/get_cus_info', methods=['GET','POST'])
def get_cus_info():
    cus_name = request.json.get('selected_name')
    work_dir=wecom_dir()
    fn=os.path.join(work_dir,cus_name+'.xlsm')
    p=cus_data.CusData()
    res=p.cus_cls_rec_toweb(fn=fn,cls_types=['常规私教课','限时私教课','常规团课','限时团课'],not_lmt_types=['常规私教课','常规团课'])
    res.fillna(0)
    data=res.iloc[0].to_dict()
    return jsonify(data)

@app.route('/open_cus_fn',methods=['POST'])
def open_cus_fn():
    cus_name=request.data.decode('utf-8')
    cus_li=cus_list()
    if cus_name and cus_name in cus_li:
        work_dir=wecom_dir()
        fn=os.path.join(work_dir,cus_name+'.xlsm')
        # os.startfile(fn)
        return f'正在打开 {cus_name} 的会员档案'
    else:
        return '会员编码及编码为空/无此会员档案'

@app.route('/check_new',methods=['POST'])
def check_new():
    dat=request.data
    cus_li=cus_list()
    cus_num=[int(x[2:5]) for x in cus_li]
    max_num=max(cus_num)
    new_num=max_num+1
    txt_num=str(new_num).zfill(3)
    # new_name='MH'+new_num.zfill(3)+cus_name+'.xlsm'
    # new_name=os.path.join(wecom_dir,new_name)
    return txt_num

@app.route('/generate_new',methods=['POST'])
def generate_new():
    try:
        fn=request.data
        fn='MH'+fn.decode('utf-8')
        work_dir=wecom_dir()
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
        os.startfile(new_fn)

        return new_fn
    except Exception as e:
        return e




@app.route('/welcome')
def welcome():
    return '关于我们页面'

if __name__ == '__main__':
    app.run(debug=True)
    # app.run(debug=True,host='192.168.1.38',port=5000)
    # res=wecom_dir()
    # print(res)
