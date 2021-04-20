import os
import sys
sys.path.append(os.path.join(os.path.dirname(__file__),'modules'))
import main

# print(os.path.join(os.path.dirname(__file__),'modules'))

p=main.MingHu()
p.draw(cus='MH024刘姐',ins='MHINS001陆伟杰',start_time='20200315',end_time='20210420')

