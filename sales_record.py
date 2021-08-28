import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import readconfig
import pandas



class MingHuSales:
    def __init__(self):
        config=readconfig.exp_json(os.path.join(os.path.dirname(__file__),'configs','main.config')) 
        self.sale_dir=config['购课文件夹']
        print(self.sale_dir)




if __name__=='__main__':
    sales=MingHuSales()
