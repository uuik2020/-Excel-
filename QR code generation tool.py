# 导入所需的库
import tkinter as tk
import qrcode
from PIL import ImageTk, Image
import pandas as pd
import os
from datetime import datetime

# 初始化全局变量
label = None
num_label = None
num = None

# 定义生成二维码的函数
def generate_qr():
    global label, num_label, num
    # 如果之前生成过二维码，先删除旧的二维码
    if label is not None:
        label.destroy()
    if num_label is not None:
        num_label.destroy()

    # 创建二维码对象
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    # 添加数据到二维码对象
    qr.add_data(str(num))
    # 生成二维码
    qr.make(fit=True)

    # 保存二维码为图片
    img = qr.make_image(fill='black', back_color='white')
    img.save('qrcode.png')

    # 显示二维码图片
    image = Image.open('qrcode.png')
    photo = ImageTk.PhotoImage(image)

    label = tk.Label(image=photo)
    label.image = photo
    label.pack()

    # 显示数字
    num_label = tk.Label(text=str(num))
    num_label.pack()

# 定义保存数据到Excel的函数
def save_to_excel():
    global num
    # 获取输入的数据和数字
    data = [entry.get(), num, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    # 指定数据存储路径
    data_storage_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Data Storage')
    # 如果数据存储路径不存在，则创建
    if not os.path.exists(data_storage_path):
        os.mkdir(data_storage_path)
    # 指定数据文件路径
    data_file_path = os.path.join(data_storage_path, 'data.xlsx')
    # 如果数据文件存在，则读取数据文件
    if os.path.exists(data_file_path):
        df = pd.read_excel(data_file_path, header=None)
        # 添加新的数据到数据文件
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    else:
        # 如果数据文件不存在，则创建新的数据文件
        df = pd.DataFrame([data])
    # 保存数据到数据文件
    df.to_excel(data_file_path, index=False, header=False)
    # 数字加一
    num += 1

# 创建主窗口
root = tk.Tk()
# 设置窗口标题
root.title("二维码生成工具")

# 指定数据存储路径和数据文件路径
data_storage_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Data Storage')
data_file_path = os.path.join(data_storage_path, 'data.xlsx')
# 如果数据文件存在，则读取数据文件中的最后一个数字
if os.path.exists(data_file_path):
    df = pd.read_excel(data_file_path, header=None)
    num = int(df.iloc[-1, 1])
else:
    # 如果数据文件不存在，则设置初始数字
    num = 11380458

# 创建输入框
entry = tk.Entry(root)
entry.pack()

# 创建生成二维码的按钮
button_qr = tk.Button(root, text="生成二维码", command=generate_qr)
button_qr.pack()

# 创建保存数据到Excel的按钮
button_save = tk.Button(root, text="保存到Excel", command=save_to_excel)
button_save.pack()

# 运行主循环
root.mainloop()
