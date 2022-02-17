import os
import sys
from PIL import Image,ImageFont,ImageDraw
from cv2 import cv2
import numpy as np


class ColorTransfer:
    def  draw_fill(self):
        img=Image.new('RGB',(300,500),'#FFFF88')
        img.show()

    
    def hex_to_rgb(self,hex):
        if hex[0]=='#':
            hex=hex[1:]
        
        r = int(hex[0:2],16)
        g = int(hex[2:4],16)
        b = int(hex[4:], 16)
        # print('hex to rgb result:', r,g,b)
        return r,g,b

    def rgb_to_hsv(self,rgb):
        r, g, b = rgb[0]/255.0, rgb[1]/255.0, rgb[2]/255.0
        mx = max(r, g, b)
        mn = min(r, g, b)
        m = mx-mn
        if mx == mn:
            h = 0
        elif mx == r:
            if g >= b:
                h = ((g-b)/m)*60
            else:
                h = ((g-b)/m)*60 + 360
        elif mx == g:
            h = ((b-r)/m)*60 + 120
        elif mx == b:
            h = ((r-g)/m)*60 + 240
        if mx == 0:
            s = 0
        else:
            s = m/mx
        v = mx
        return h, s, v

    def hsl_to_rgb(self,hsl):
        hls = [[[hsl[0], hsl[2], hsl[1]]]]  # hsl to hls
        rgb_normal = cv2.cvtColor(np.array(hls, dtype=np.float32), cv2.COLOR_HLS2RGB)
        return int(rgb_normal[0][0][0] * 255), int(rgb_normal[0][0][1] * 255), int(rgb_normal[0][0][2] * 255)

    # HSL渐变色
    def get_multi_colors_by_hsl(self,begin_color, end_color, color_count):

        if color_count < 2:
            return []

        if begin_color[0]=='#':
            begin_color=self.hex_to_rgb(begin_color)
        if end_color[0]=='#':
            end_color=self.hex_to_rgb(end_color)

        colors = []
        hsl1 = self.rgb_to_hsv(begin_color)
        hsl2 = self.rgb_to_hsv(end_color)
        steps = [(hsl2[i] - hsl1[i]) / (color_count - 1) for i in range(3)]
        for color_index in range(color_count):
            hsl = [hsl1[i] + steps[i] * color_index for i in range(3)]
            colors.append(self.hsl_to_rgb(hsl))
        return colors

class FillGradient(ColorTransfer):
    def draw_rct_hsv(self,w,h,begin_color,end_color,step='auto'):
        if step=='auto':
            step=h
        fill_colors=self.get_multi_colors_by_hsl(begin_color=begin_color,end_color=end_color,color_count=step)
        rct=Image.new('RGB',(w,h),'#ffffff')
        draw=ImageDraw.Draw(rct)

        h_step=round(h/step)
        for n in range(0,step):
            for x in range(0,w):
                for y in range(round(n*h_step-h_step),round(n*h_step)):
                    draw.point((x,y),fill=fill_colors[n])
        rct.show()

    def RGB(self,r,g,b): 
        return (r,g,b)

    def make_img_data(self,width, height, rgb):
        '''Make image data'''
        result = np.zeros((height, width, 3), dtype=np.uint8)
        for i, v in enumerate(rgb):
            result[:,:,i] = np.tile(np.linspace(v, v, width), (height, 1))
        
        return result

    def make_gradation_img_data(self,width, height, begin_color, end_color, horizontal=(True, True, True),direction='vertical'):

        result = np.zeros((height, width, 3), dtype=np.uint8)
        for i, (m,n,o) in enumerate(zip(begin_color, end_color, horizontal)):
            if o:
                if direction=='vertical':
                    result[:,:,i] = np.tile(np.linspace(m, n, height), (width, 1)).T
                elif direction=='horizon':
                    result[:,:,i] = np.tile(np.linspace(m, n, width), (height, 1))
            else:
                if direction=='vertical':
                    result[:,:,i] = np.tile(np.linspace(m, n, height), (width, 1))
                elif direction=='horizon':
                    result[:,:,i] = np.tile(np.linspace(m, n, width), (height, 1)).T
        
        return result


    def draw_gradient_rct_rgb(self,w,h,begin_color,end_color,horizontal=(True, True, True),direction='vertical'):
        # self,w,h,begin_color,end_color,step='auto'
        if begin_color[0]=='#':
            begin_color=self.hex_to_rgb(begin_color)
        if end_color[0]=='#':
            end_color=self.hex_to_rgb(end_color)
        # make_img = lambda w, h, rgb: Image.fromarray(self.make_img_data(w, h, rgb))
        make_gradient_img = lambda w, h, begin_color, end_color, horizontal: Image.fromarray(self.make_gradation_img_data(w, h, begin_color, end_color, horizontal,direction=direction))
        rct=make_gradient_img(w, h, begin_color, end_color, horizontal=horizontal)

        return rct

    def fill_gradient_rct_rgb(self,img,begin_color,end_color,horizontal=(True, True, True),direction='vertical'):
        w,h=img.size[0],img.size[1]
        if begin_color[0]=='#':
            begin_color=self.hex_to_rgb(begin_color)
        if end_color[0]=='#':
            end_color=self.hex_to_rgb(end_color)
        make_gradient_img = lambda w, h, begin_color, end_color, horizontal: Image.fromarray(self.make_gradation_img_data(w, h, begin_color, end_color, horizontal,direction=direction))
        rct=make_gradient_img(w, h, begin_color, end_color, horizontal=horizontal)

        return rct

    def fill_multi_gradient_rct_rgb(self,img,colors,horizontal=(True, True, True),direction='vertical'):
        w,h=img.size[0],img.size[1]
        parts=[]
        if direction=='vertical':
            for part_num in range(len(colors)-1):
                if part_num==len(colors)-2:
                    parts.append([h*part_num//(len(colors)-1),h+1])
                else:
                    parts.append([h*part_num//(len(colors)-1),h*(part_num+1)//(len(colors)-1)])

            for i,part in enumerate(parts):
                if i<len(colors)-1:
                    begin_color,end_color=colors[i],colors[i+1]
                    if i==len(parts)-1:
                        img_part=img.crop((0,parts[i][0],w,parts[i][1]-1))
                    else:
                        img_part=img.crop((0,parts[i][0],w,parts[i][1]))
                    cover=self.fill_gradient_rct_rgb(img=img_part,begin_color=begin_color,end_color=end_color,horizontal=horizontal,direction=direction)
                    img.paste(cover,(0,parts[i][0]))
        elif direction=='horizon':
            for part_num in range(len(colors)-1):
                if part_num==len(colors)-2:
                    parts.append([w*part_num//(len(colors)-1),w+1])
                else:
                    parts.append([w*part_num//(len(colors)-1),w*(part_num+1)//(len(colors)-1)])
            # print(parts)
            for i,part in enumerate(parts):
                if i<len(colors)-1:
                    begin_color,end_color=colors[i],colors[i+1]
                    if i==len(parts)-1:
                        img_part=img.crop((parts[i][0],0,parts[i][1]-1,h))
                    else:
                        img_part=img.crop((parts[i][0],0,parts[i][1],h))
                    cover=self.fill_gradient_rct_rgb(img=img_part,begin_color=begin_color,end_color=end_color,horizontal=horizontal,direction=direction)
                    img.paste(cover,(parts[i][0],0))

        return img

if __name__=='__main__':
    # p=ColorTransfer()
    # clrs=p.get_multi_colors_by_hsl((255,255,136),(230,180,90),10)
    rct=FillGradient()
    # img=rct.draw_gradient_rct_rgb(200,500,'#fffdee','#ffad6a',30,horizontal=(True, True, True),direction='vertical')
    img=Image.new('RGBA',(600,800),'#ffffff')
    # img=rct.fill_gradient_rct_rgb(img,'#fffdee','#ffad6a',horizontal=(True, True, True),direction='vertical')
    img=rct.fill_multi_gradient_rct_rgb(img,('#fffdee','#ffad6a','#fed32a'),horizontal=(True, True, True),direction='vertical')
    img.show()