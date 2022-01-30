import os
import sys
import pic_transfer
import numpy as np
import matplotlib.pyplot as plt 
plt.rcParams['font.sans-serif']=['SimHei']  # 黑体


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

    def draw(self,data):
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
        fig = plt.figure(figsize=(6,5))
        # 这里一定要设置为极坐标格式
        ax = fig.add_subplot(111, polar=True)
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

if __name__=='__main__':
    data={'ht_lung':8,'balance':8,'power':5, 'flexibility':3,'core':7}
    p=DrawRadar().draw(data)
    p.show()