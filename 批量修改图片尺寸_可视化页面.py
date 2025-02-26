import os
import re
import time
import datetime
import logging
import json
from logging.handlers import RotatingFileHandler
import threading
from tkinter import *
from tkinter import filedialog, scrolledtext
from PIL import Image

# **全局变量**
folder_path = ""
TRIAL_END_TIME = datetime.datetime(2025, 2, 26, 16, 59, 59)  # 试用截止时间
LOG_FILE = "processing_log.txt"  # 日志文件路径
MAX_LOG_FILE_SIZE = 20 * 1024 * 1024  # 5 MB 日志文件大小限制
stop_processing = False  # 停止处理的标志
MAX_LOG_LINES = 500  # 日志最多显示 1000 行
processing_thread = None  # 处理线程
scan_timer = None  # 定时器
CONFIG_FILE = "config.json"  # 配置文件路径


# **读取配置文件**
def load_config():
    """ 读取配置文件，获取默认文件夹路径 """
    global folder_path
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            try:
                config = json.load(config_file)
                folder_path = config.get("folder_path", "")
                write_log(f"🔧 已加载配置文件，默认文件夹路径：{folder_path}")
                folder_label.config(text=f"已加载默认配置文件夹: {folder_path}")  # 显示加载后的路径
                start_button.config(state=NORMAL)  # 启用“开始处理”按钮
            except json.JSONDecodeError:
                write_log("⚠️ 配置文件格式错误，请重新配置或手动选择文件夹")
    else:
        write_log("⚠️ 配置文件不存在，请重新配置或手动选择文件夹")


# **设置日志文件**
def setup_logging():
    """ 设置日志输出，包含实时显示和写入文件 """
    if not os.path.exists("logs"):
        os.makedirs("logs")

    log_file_path = os.path.join("logs", LOG_FILE)

    handler = RotatingFileHandler(log_file_path, maxBytes=MAX_LOG_FILE_SIZE, backupCount=5)
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)


# **可视化页面 - 更新日志显示**
def write_log(message):
    """ 在日志窗口和文件中同时显示日志 """
    # 添加当前时间信息
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_message = f"[{current_time}] {message}"

    # 检查日志行数，超过 1000 行则清空
    if int(log_text.index('end-1c').split('.')[0]) >= MAX_LOG_LINES:
        log_text.delete(1.0, END)  # 清空日志框

    # 显示日志
    log_text.insert(END, formatted_message + '\n')
    log_text.yview(END)  # 滚动到最底部
    logging.info(formatted_message)


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
    """ 读取文件夹名称提取尺寸，并批量调整图片大小（不保持比例，直接拉伸变形），转换为CMYK颜色模式 """
    global stop_processing

    try:
        os.listdir(folder_path)
    except Exception as e:
        write_log(f"❌ 文件目录 {folder_path} 读取失败: {e}")

    for folder_name in os.listdir(folder_path):
        subfolder_path = os.path.join(folder_path, folder_name)

        if os.path.isdir(subfolder_path):
            dimensions = extract_dimensions_from_folder_name(folder_name)

            if not dimensions:
                write_log(f"⚠️ 无法从文件夹 '{folder_name}' 提取尺寸，跳过处理")
                continue

            width_cm, height_cm = dimensions
            target_width = cm_to_pixels(width_cm)
            target_height = cm_to_pixels(height_cm)

            write_log(f"📏 处理文件夹: {folder_name}, 目标尺寸: {target_width}x{target_height} 像素")

            for filename in os.listdir(subfolder_path):
                if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
                    image_path = os.path.join(subfolder_path, filename)

                    try:
                        with Image.open(image_path) as image:
                            image = image.convert("CMYK")  # **转换为CMYK模式**

                            # **判断图片尺寸是否已经符合要求**
                            if image.size == (target_width, target_height):
                                write_log(f"📷 图片 '{image_path}' 尺寸已经符合要求，跳过处理")
                                continue  # 跳过该图片

                            write_log(f"📷 处理 {image_path} (原尺寸: {image.size})...")

                            # **拉伸变形缩放**
                            resized_image = image.resize((target_width, target_height), Image.LANCZOS)

                            # **保存**
                            resized_image.save(image_path)
                            write_log(f"✅ 已调整并覆盖: {image_path}")

                    except Exception as e:
                        write_log(f"❌ 处理 {image_path} 失败: {e}")

    # 每次扫描结束后加一个分隔线
    write_log("-------------------------------------------------------------------------")
    write_log("-------------------------------------------------------------------------")
    write_log("-------------------------------------------------------------------------")

    if not stop_processing:
        # 继续扫描，设置定时器每5秒调用一次
        global scan_timer
        scan_timer = threading.Timer(5, process_images_in_folder, args=(folder_path,))
        scan_timer.start()
    else:
        write_log("🚫 已停止文件扫描和处理")


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
        label.config(text=f"试用剩余时间: {int(days)}天 {int(hours):02d}:{int(mins):02d}:{int(secs):02d}", fg="red")
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
    global stop_processing, scan_timer

    # 停止之前的扫描
    if scan_timer and scan_timer.is_alive():
        write_log("🚫 已有扫描在进行中，等待当前扫描完成")
        return

    stop_processing = False
    # 启动扫描
    write_log("🚀 开始处理文件夹")
    process_images_in_folder(folder_path)


def stop_processing_function():
    """ 停止文件夹处理 """
    global stop_processing, scan_timer

    stop_processing = True
    if scan_timer and scan_timer.is_alive():
        scan_timer.cancel()  # 取消定时器
        write_log("🚫 已停止文件扫描和处理")
    else:
        write_log("🚫 没有正在运行的扫描")


# **GUI界面**
root = Tk()
root.title("图片尺寸调整小工具-试用版V1.0")
root.geometry("600x600")

# 选择文件夹按钮
folder_button = Button(root, text="选择文件夹", command=browse_folder)
folder_button.pack(pady=20)

# 显示文件夹路径
folder_label = Label(root, text="请选择文件夹", wraplength=350)
folder_label.pack()

# 开始按钮
start_button = Button(root, text="开始处理文件", state=DISABLED, command=start_processing)
start_button.pack(pady=10)

# 停止按钮
stop_button = Button(root, text="停止处理文件", state=NORMAL, command=stop_processing_function)
stop_button.pack(pady=10)

# 倒计时标签（红色字体）
time_label = Label(root, text="", font=("Arial", 14), fg="red")
time_label.pack(pady=10)

# 显示固定的试用截止时间
end_time_label = Label(root, text=f"试用截止时间: {TRIAL_END_TIME.strftime('%Y-%m-%d %H:%M:%S')}", font=("Arial", 12), fg="red")
end_time_label.pack()

# 日志显示框
log_text = scrolledtext.ScrolledText(root, width=90, height=30, wrap=WORD, font=("Arial", 12))
log_text.pack(pady=10)

# 启动倒计时线程
threading.Thread(target=countdown_timer, args=(time_label,)).start()

# 启动日志配置
setup_logging()

# 加载配置文件
load_config()

# 运行界面
root.mainloop()
