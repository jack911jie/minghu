import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main
import seven_GUI

# print(os.path.join(os.path.dirname(__file__),'modules'))

def cus_feedback(place='minghu',cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701',adj_bfr='yes',adj_src='prg',gui=''):
    p=main.MingHu(place=place,adj_bfr=adj_bfr,adj_src=adj_src,gui=gui)
    p.draw(cus=cus,ins=ins,start_time=start_time,end_time=end_time)

def period_summary(place='minghu',cus_name_input='MH017李俊娴',ins='MHINS001陆伟杰',start_date='20210401',end_date='20220201',
                    theme='lightgrey',ico_size=(40,40),diary_font_size=26,diet_font_size=26,diet_boxwid=580,adj_bfr='yes',adj_src='prg',gui='',logo_ht=52):
    p=main.PeroidSummary(place=place,adj_bfr=adj_bfr,adj_src=adj_src,gui=gui)
    p.exp_chart(cus_name_input=cus_name_input,ins=ins,start_date=start_date,end_date=end_date,
                theme=theme,ico_size=ico_size,diary_font_size=diary_font_size,diet_font_size=diet_font_size,diet_boxwid=diet_boxwid,logo_ht=logo_ht)
    
def group_input(place='minghu'):
    p=main.GroupDataInput(place=place)
    p.data_input()

def today_feedback(place='minghu',cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210324'):
    p=main.FeedBackAfterClass(place=place)
    p.draw_new(cus=cus,ins=ins,date_input=date_input)

def auto_xls(place='seven',cus_name_input='',mode='prgrm',gui=''):
    p=main.MingHu(place=place)
    new_cus_fn=p.auto_cus_xls(cus_name_input=cus_name_input,mode=mode,gui=gui)
    return new_cus_fn

def run_seven_gui(place='seven'):
    minghu_gui=seven_GUI.GUI(place=place)
    minghu_gui.creat_gui()
    # minghu_gui.get_cus_list()


if __name__=='__main__':
    #反馈
    # cus_feedback(cus='MH017李俊娴',ins='MHINS001陆伟杰',start_time='20210526',end_time='20210701')

    #批量录入
    # group_input()

    #当天课后生成
    # today_feedback(cus='MH024刘婵桢',ins='MHINS002韦越棋',date_input='20210619')

    #新增会员
    # auto_xls(cus_name_input='测试',mode='prgrm',gui='')

    #运行最新版本的gui
    run_seven_gui(place='seven')

    # period_summary(place='minghu',cus_name_input='MH017李俊娴',ins='MHINS001陆伟杰',start_date='20210401',end_date='20220201',
    #                             theme='lightgrey',ico_size=(40,40),diary_font_size=26,diet_font_size=26,diet_boxwid=580,adj_bfr='yes',adj_src='prg',gui='')
             

    # period_summary(place='minghu',cus_name_input='MH017李俊娴',ins='MHINS001陆伟杰',start_date='20210401',end_date='20220201',
    #                             theme='lightgrey',ico_size=(40,40),diary_font_size=26,diet_font_size=26,diet_boxwid=580,adj_bfr='yes',adj_src='prg',gui='')
             



