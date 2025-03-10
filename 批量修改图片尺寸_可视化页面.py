import os
import re
import time
import datetime
import logging
import json
from logging.handlers import RotatingFileHandler
import threading
import queue
from tkinter import *
from tkinter import filedialog, scrolledtext
from PIL import Image, ImageCms
Image.MAX_IMAGE_PIXELS = 1000000000  # 设置为5亿像素，适应你的大图

# 全局变量
folder_path = ""
TRIAL_END_TIME = datetime.datetime(2025, 3, 10, 17, 59, 59)
LOG_FILE = "processing_log.txt"
MAX_LOG_FILE_SIZE = 20 * 1024 * 1024
stop_processing = False
MAX_LOG_LINES = 500
scan_thread = None  # 后台线程
CONFIG_FILE = "config.json"
icc_profile = "USWebCoatedSWOP.icc"
log_queue = queue.Queue()

# 设置日志
def setup_logging():
    if not os.path.exists("logs"):
        os.makedirs("logs")
    handler = RotatingFileHandler(f"logs/{LOG_FILE}", maxBytes=MAX_LOG_FILE_SIZE, backupCount=5)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.basicConfig(level=logging.INFO, handlers=[handler])

def load_config():
    """ 读取配置文件，获取默认文件夹路径 """
    global folder_path
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            try:
                config = json.load(config_file)
                folder_path = config.get("folder_path", "")
                if folder_path == "":
                    write_log("⚠️ 配置文件格式错误，请重新配置或手动选择文件夹")
                else:
                    write_log(f"🔧 已加载配置文件，默认文件夹路径：{folder_path}")
                    folder_label.config(text=f"已加载默认配置文件夹: {folder_path}")  # 显示加载后的路径
                    start_button.config(state=NORMAL)  # 启用“开始处理”按钮
            except json.JSONDecodeError:
                write_log("⚠️ 配置文件格式错误，请重新配置或手动选择文件夹")
    else:
        write_log("⚠️ 配置文件不存在，请重新配置或手动选择文件夹")


def write_log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"[{timestamp}] {message}"
    log_queue.put(log_message)
    logging.info(log_message)

def update_log_window():
    while not log_queue.empty():
        log_text.insert(END, log_queue.get() + "\n")
        log_text.yview(END)
    log_text.after(500, update_log_window)

def cm_to_pixels(cm, dpi=72):
    return round(cm * dpi / 2.54)

def extract_dimensions_from_folder_name(folder_name):
    match = re.search(r'(\d+(\.\d+)?)[xX](\d+(\.\d+)?)(CM|cm)?', folder_name)
    if match:
        return float(match.group(1)), float(match.group(3))
    return None

def convert_rgb_to_cmyk(image, icc_profile_path):
    cmyk_profile = ImageCms.getOpenProfile(icc_profile_path)
    srgb_profile = ImageCms.createProfile("sRGB")
    return ImageCms.profileToProfile(image, srgb_profile, cmyk_profile, outputMode="CMYK")

def process_images_in_folder(root_folder):
    global stop_processing

    write_log(f"📏 扫描根目录: {root_folder} 开始 ")

    for current_folder, subfolders, filenames in os.walk(root_folder):
        if stop_processing:
            write_log("🚫 停止信号收到，提前终止扫描")
            return

        folder_name = os.path.basename(current_folder)

        dimensions = extract_dimensions_from_folder_name(folder_name)

        if not dimensions:
            write_log(f"⚠️ 文件夹 '{current_folder}' 名称不符合尺寸格式，跳过")
            continue

        width_cm, height_cm = dimensions
        target_width = cm_to_pixels(width_cm)
        target_height = cm_to_pixels(height_cm)

        write_log(f"📏 处理文件夹: {current_folder}, 目标尺寸: {target_width}x{target_height} 像素")

        for filename in filenames:
            if stop_processing:
                write_log("🚫 停止信号收到，提前终止图片处理")
                return

            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
                image_path = os.path.join(current_folder, filename)
                time.sleep(1)

                try:
                    with Image.open(image_path) as image:
                        image = image.convert("RGB")

                        if image.size == (target_width, target_height):
                            write_log(f"✅ 图片 '{image_path}' 尺寸已符合要求，跳过")
                            continue

                        resized_image = image.resize((target_width, target_height), Image.LANCZOS)
                        cmyk_image = convert_rgb_to_cmyk(resized_image, icc_profile)

                        # 强制保存为JPEG格式
                        jpg_image_path = os.path.splitext(image_path)[0] + ".jpg"
                        cmyk_image.save(jpg_image_path, 'JPEG', quality=90)

                        # 如果原文件不是jpg，则删除原文件
                        if not image_path.lower().endswith('.jpg'):
                            os.remove(image_path)

                        write_log(f"✅ 已处理并覆盖: {image_path}")

                except Exception as e:
                    write_log(f"❌ 处理失败: {image_path}, 错误: {e}")

    write_log("✅✅✅------------ 本次扫描处理图片完成！！！ ------------")
    write_log("✅✅✅------------ 本次扫描处理图片完成！！！ ------------")
    write_log("✅✅✅------------ 本次扫描处理图片完成！！！ ------------")
    write_log("✅✅✅------------ 本次扫描处理图片完成！！！ ------------")
    write_log("✅✅✅------------ 本次扫描处理图片完成！！！ ------------")
    start_button.config(state="normal")
    stop_button.config(state="disabled")

def start_threaded_processing():
    global scan_thread, stop_processing
    stop_processing = False
    write_log("🚀 开始扫描")
    scan_thread = threading.Thread(target=process_images_in_folder, args=(folder_path,))
    scan_thread.start()
    start_button.config(state="disabled")
    stop_button.config(state="normal")

def countdown_timer(label):
    while True:
        remaining = TRIAL_END_TIME - datetime.datetime.now()
        if remaining.total_seconds() <= 0:
            label.config(text="试用时间已结束!", fg="red")
            folder_button.config(state="disabled")
            start_button.config(state="disabled")
            stop_button.config(state="disabled")
            break

        days, rem = divmod(remaining.total_seconds(), 86400)
        hours, rem = divmod(rem, 3600)
        mins, secs = divmod(rem, 60)
        label.config(text=f"试用剩余时间: {int(days)}天 {int(hours):02d}:{int(mins):02d}:{int(secs):02d}", fg="red")
        time.sleep(1)

def browse_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=f"已选择文件夹: {folder_path}")
        start_button.config(state=NORMAL)

def stop_processing_function():
    global stop_processing
    stop_processing = True
    write_log("🚫 已请求停止处理")

# GUI界面
root = Tk()
root.title("图片尺寸调整小工具-试用版V5.1")
root.geometry("800x600")

folder_button = Button(root, text="选择文件夹", command=browse_folder)
folder_button.pack(pady=10)

folder_label = Label(root, text="请选择文件夹")
folder_label.pack()

start_button = Button(root, text="开始处理", state=DISABLED, command=start_threaded_processing)
start_button.pack()

stop_button = Button(root, text="停止处理", command=stop_processing_function)
stop_button.pack()
stop_button.config(state="disabled")

time_label = Label(root, text="", font=("Arial", 14), fg="red")
time_label.pack()

end_time_label = Label(root, text=f"试用截止时间: {TRIAL_END_TIME.strftime('%Y-%m-%d %H:%M:%S')}", fg="red")
end_time_label.pack()

log_text = scrolledtext.ScrolledText(root, width=90, height=30, wrap=WORD)
log_text.pack()

update_log_window()
threading.Thread(target=countdown_timer, args=(time_label,), daemon=True).start()

setup_logging()
load_config()

root.mainloop()
