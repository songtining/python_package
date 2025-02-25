import os
import re
import time
import datetime
import threading
from PIL import Image as PilImage
from tkinter import *
from tkinter import filedialog


# **全局变量**
folder_path = ""
TRIAL_END_TIME = datetime.datetime(2025, 3, 1, 12, 0, 0)  # 试用截止时间
stop_event = threading.Event()  # 用于停止线程


def cm_to_pixels(cm, dpi=96):
    """ 将厘米转换为像素（默认 96 DPI）"""
    return int(cm * dpi / 2.54)  # 1 英寸 = 2.54cm


def extract_dimensions_from_folder_name(folder_name):
    """ 从文件夹名称中提取尺寸（宽 X 高），支持多种格式 """
    match = re.search(r'(\d+)[xX](\d+)(CM|cm)?', folder_name)
    if match:
        width_cm = int(match.group(1))
        height_cm = int(match.group(2))
        return width_cm, height_cm
    return None


def process_images_in_folder(folder_path):
    """ 每隔 5 秒检查图片尺寸，不符合则调整 """
    while datetime.datetime.now() < TRIAL_END_TIME:
        if stop_event.is_set():
            print("❌ 停止了文件夹处理")
            break  # 停止文件夹处理

        for folder_name in os.listdir(folder_path):
            subfolder_path = os.path.join(folder_path, folder_name)

            if os.path.isdir(subfolder_path):
                dimensions = extract_dimensions_from_folder_name(folder_name)

                if not dimensions:
                    print(f"⚠️ 无法从文件夹 '{folder_name}' 提取尺寸，跳过处理")
                    continue

                width_cm, height_cm = dimensions
                target_width = cm_to_pixels(width_cm)
                target_height = cm_to_pixels(height_cm)

                print(f"📏 处理文件夹: {folder_name}, 目标尺寸: {target_width}x{target_height} 像素")

                for filename in os.listdir(subfolder_path):
                    if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
                        image_path = os.path.join(subfolder_path, filename)

                        try:
                            with PilImage.open(image_path) as image:
                                image = image.convert("RGB")  # 确保为 RGB 格式
                                if image.size != (target_width, target_height):
                                    print(f"📷 处理 {filename} (原尺寸: {image.size})...")
                                    resized_image = image.resize((target_width, target_height), PilImage.LANCZOS)
                                    resized_image.save(image_path)
                                    print(f"✅ 已调整并覆盖: {image_path}")

                        except Exception as e:
                            print(f"❌ 无法处理 {filename}: {e}")

        time.sleep(5)  # **每隔 5 秒检查一次**


def countdown_timer(label):
    """ 显示固定的试用截止时间倒计时 """
    while True:
        now = datetime.datetime.now()
        remaining_time = TRIAL_END_TIME - now
        if remaining_time.total_seconds() <= 0:
            label.config(text="试用时间已结束!", fg="red")
            return
        days, remainder = divmod(remaining_time.total_seconds(), 86400)
        hours, remainder = divmod(remainder, 3600)
        mins, secs = divmod(remainder, 60)
        label.config(text=f"试用剩余时间: {int(days)}天 {int(hours):02d}小时 {int(mins):02d}分钟 {int(secs):02d}秒", fg="red")
        time.sleep(1)


def browse_folder():
    """ 选择文件夹 """
    global folder_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=f"已选择文件夹: {folder_path}")
        start_button.config(state=NORMAL)  # 允许点击开始按钮


def start_processing():
    """ 启动文件夹处理 """
    stop_event.clear()  # 清除停止事件标志
    threading.Thread(target=process_images_in_folder, args=(folder_path,)).start()


def stop_processing():
    """ 停止文件夹处理 """
    stop_event.set()  # 设置停止事件标志
    print("✅ 停止文件夹处理")


# **GUI界面**
root = Tk()
root.title("批量调整图片尺寸小工具")
root.geometry("500x400")

# 选择文件夹按钮
folder_button = Button(root, text="选择文件夹", command=browse_folder)
folder_button.pack(pady=20)

# 显示文件夹路径
folder_label = Label(root, text="请选择文件夹", wraplength=350)
folder_label.pack()

# 开始按钮
start_button = Button(root, text="开始处理", state=DISABLED, command=start_processing)
start_button.pack(pady=20)

# 停止按钮
stop_button = Button(root, text="停止处理", command=stop_processing)
stop_button.pack(pady=20)

# 倒计时标签（红色字体）
time_label = Label(root, text="", font=("Arial", 14), fg="red")
time_label.pack(pady=20)

# 显示固定的试用截止时间
end_time_label = Label(root, text=f"试用截止时间: {TRIAL_END_TIME.strftime('%Y-%m-%d %H:%M:%S')}", font=("Arial", 12), fg="red")
end_time_label.pack()
threading.Thread(target=countdown_timer, args=(time_label,)).start()

# 运行界面
root.mainloop()