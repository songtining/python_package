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

# **全局变量**
folder_path = ""
TRIAL_END_TIME = datetime.datetime(2025, 3, 4, 17, 59, 59)  # 试用截止时间
LOG_FILE = "processing_log.txt"  # 日志文件路径
MAX_LOG_FILE_SIZE = 20 * 1024 * 1024  # 5 MB 日志文件大小限制
stop_processing = False  # 停止处理的标志
MAX_LOG_LINES = 500  # 日志最多显示 1000 行
processing_thread = None  # 处理线程
scan_timer = None  # 定时器
CONFIG_FILE = "config.json"  # 配置文件路径
# 示例调用
icc_profile = "USWebCoatedSWOP.icc"  # 替换为 Photoshop 使用的 CMYK ICC 颜色配置文件路径

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
    log_queue.put(formatted_message)

    # 检查日志行数，超过 1000 行则清空
    if int(log_text.index('end-1c').split('.')[0]) >= MAX_LOG_LINES:
        log_text.delete(1.0, END)  # 清空日志框

    # 显示日志
    #log_text.insert(END, formatted_message + '\n')
    #log_text.yview(END)  # 滚动到最底部
    logging.info(formatted_message)

def update_log_window():
    """ 在主线程更新日志 """
    while not log_queue.empty():
        message = log_queue.get()
        log_text.insert('end', message + '\n')
        log_text.yview('end')  # 滚动到最底部
        log_text.update_idletasks()  # 刷新 GUI 更新
    log_text.after(500, update_log_window)  # 每100毫秒检查一次更新

def cm_to_pixels(cm, dpi=72):
    """ 将厘米转换为像素（默认 72 DPI），并四舍五入保留整数 """
    pixels = cm * dpi / 2.54
    return round(pixels)  # 四舍五入返回整数

def extract_dimensions_from_folder_name(folder_name):
    """
    从文件夹名称中提取尺寸（宽 X 高），支持整数和小数格式
    支持格式：
    - 125X215CM
    - 125.5X215.1CM
    - 125x215
    - 125.5x215.1
    """
    match = re.search(r'(\d+(\.\d+)?)[xX](\d+(\.\d+)?)(CM|cm)?', folder_name)
    if match:
        width_cm = float(match.group(1))   # 直接改为float支持小数
        height_cm = float(match.group(3))  # group(3)是高度部分
        return width_cm, height_cm
    return None


def convert_rgb_to_cmyk(image, icc_profile_path):
    """
    将 RGB 图像转换为 CMYK 并应用 ICC 颜色配置文件，最终保存为高质量 JPEG
    """
    cmyk_profile = ImageCms.getOpenProfile(icc_profile_path)
    srgb_profile = ImageCms.createProfile("sRGB")

    # 颜色管理转换 (RGB -> CMYK)
    cmyk_image = ImageCms.profileToProfile(image, srgb_profile, cmyk_profile, outputMode="CMYK")

    return cmyk_image

def process_images_in_folder(root_folder):
    """
    遍历根目录下的所有文件夹（包括多级子目录），
    发现文件夹名符合尺寸格式的，就对该文件夹下的图片做处理。
    """
    global stop_processing

    write_log(f"📏 扫描根目录: {root_folder} 开始 ******************************** ")

    # 使用os.walk递归遍历所有目录
    for current_folder, subfolders, filenames in os.walk(root_folder):
        folder_name = os.path.basename(current_folder)

        dimensions = extract_dimensions_from_folder_name(folder_name)

        if not dimensions:
            write_log(f"⚠️ 文件夹 '{current_folder}' 名称不符合尺寸格式，跳过")
            continue  # 跳过不符合尺寸格式的文件夹

        width_cm, height_cm = dimensions
        target_width = cm_to_pixels(width_cm)
        target_height = cm_to_pixels(height_cm)

        write_log(f"📏 处理文件夹: {current_folder}, 目标尺寸: {target_width}x{target_height} 像素")

        for filename in filenames:
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
                image_path = os.path.join(current_folder, filename)

                try:
                    with Image.open(image_path) as image:
                        image = image.convert("RGB")  # 确保为RGB模式

                        # 判断图片尺寸是否已经符合要求
                        if image.size == (target_width, target_height):
                            write_log(f"📷 图片 '{image_path}' 尺寸已经符合要求，跳过处理")
                            continue

                        write_log(f"📷 处理 {image_path} (原尺寸: {image.size})...")

                        # 拉伸变形缩放
                        resized_image = image.resize((target_width, target_height), Image.LANCZOS)

                        # 转换为CMYK
                        cmyk_image = convert_rgb_to_cmyk(resized_image, icc_profile)

                        # 保存为JPEG，覆盖原图
                        cmyk_image.save(image_path, 'JPEG', quality=90)
                        write_log(f"✅ 已调整并覆盖: {image_path}")

                except Exception as e:
                    write_log(f"❌ 处理 {image_path} 失败: {e}")

    # 每次扫描结束后加分隔线
    write_log("-------------------------------------------------------------------------")
    write_log("-------------------------------------------------------------------------")
    write_log("-------------------------------------------------------------------------")

    if not stop_processing:
        global scan_timer
        scan_timer = threading.Timer(10, process_images_in_folder, args=(root_folder,))
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
            folder_button.config(state="disabled")  # 禁用按钮
            start_button.config(state="disabled")  # 禁用按钮
            stop_button.config(state="disabled")  # 禁用按钮
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

# 创建队列
log_queue = queue.Queue()

# **GUI界面**
root = Tk()
root.title("图片尺寸调整小工具-试用版V4.0")
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

# 启动日志更新线程
log_text.after(500, update_log_window)

# 启动倒计时线程
threading.Thread(target=countdown_timer, args=(time_label,)).start()

# 启动日志配置
setup_logging()

# 加载配置文件
load_config()

# 运行界面
root.mainloop()
