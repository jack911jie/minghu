import numpy as np
import matplotlib.pyplot as plt 
import matplotlib.font_manager as fm
from matplotlib.backends.backend_agg import FigureCanvasAgg
from PIL import Image

def mat_to_pil_img(fig):
    # 将plt转化为numpy数据
    canvas = FigureCanvasAgg(plt.gcf())
    # 绘制图像
    canvas.draw()
    # 获取图像尺寸
    w, h = canvas.get_width_height()
    # 解码string 得到argb图像
    buf = np.frombuffer(canvas.tostring_argb(), dtype=np.uint8)
    fig.canvas.draw()
    # 获取图像尺寸
    w, h = fig.canvas.get_width_height()
    # 获取 argb 图像
    buf = np.frombuffer(fig.canvas.tostring_argb(), dtype=np.uint8)
    # 重构成w h 4(argb)图像
    buf.shape = (w, h, 4)
    # 转换为 RGBA
    buf = np.roll(buf, 3, axis=2)
    # 得到 Image RGBA图像对象 (需要Image对象的同学到此为止就可以了)
    image = Image.frombytes("RGBA", (w, h), buf.tobytes())
    # # 转换为numpy array rgba四通道数组
    # image = np.asarray(image)
    # # 转换为rgb图像
    # rgb_image = image[:, :, :3]
    return image