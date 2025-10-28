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
from tkinter import filedialog, scrolledtext, messagebox
from PIL import Image, ImageDraw
import math
import win32com.client
import pythoncom
import functools
Image.MAX_IMAGE_PIXELS = 1000000000  # 设置为10亿像素，适应大图

# 全局变量
folder_path = ""
LOG_FILE = "processing_log.txt"
MAX_LOG_FILE_SIZE = 20 * 1024 * 1024
stop_processing = False
MAX_LOG_LINES = 500
scan_thread = None  # 后台线程
CONFIG_FILE = "config.json"
log_queue = queue.Queue()
line_color = "white"  # 新增全局变量用于存储画线颜色
line_width = 0.06
horizontal_offset_options = ["6", "7"]

# =============== 试用期检查 ===============
def check_trial_period():
    """检查试用期是否过期"""
    # 设置试用期到期时间（精确到时分秒）
    # ⚠️ 请按实际需要修改下面的日期时间（例如 2025-12-31 23:59:59）
    expire_time = datetime.datetime(2025, 10, 30, 23, 59, 59)
    
    # 获取当前系统时间
    now = datetime.datetime.now()
    
    # 如果超过试用期
    if now > expire_time:
        root = Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror("试用期已结束", 
                           f"软件试用期已到期（{expire_time.strftime('%Y-%m-%d %H:%M:%S')}），\n"
                           f"请联系开发者获取正式版本。\n\n"
                           f"当前时间：{now.strftime('%Y-%m-%d %H:%M:%S')}")
        root.destroy()
        return False
    
    return True


# =============== 装饰器 ===============
def com_thread(func):
    """保证线程内自动初始化/释放 COM"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()
    return wrapper

# 设置日志
def setup_logging():
    if not os.path.exists("logs"):
        os.makedirs("logs")
    handler = RotatingFileHandler(f"logs/{LOG_FILE}", maxBytes=MAX_LOG_FILE_SIZE, backupCount=5)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.basicConfig(level=logging.INFO, handlers=[handler])

def load_config():
    """ 读取配置文件，获取默认文件夹路径和画线颜色 """
    global folder_path, line_color, line_width
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            try:
                config = json.load(config_file)
                folder_path = config.get("folder_path", "")
                # 读取画线颜色
                line_color = config.get("line_color", "white")
                line_width = config.get("line_width", 0.06)
                if folder_path == "":
                    write_log("⚠️ 配置文件格式错误，请重新配置或手动选择文件夹")
                else:
                    write_log(f"🔧 已加载配置文件，默认文件夹路径：{folder_path}")
                    write_log(f"🔧 已加载配置文件，画线颜色：{line_color}")
                    write_log(f"🔧 已加载配置文件，画线宽度：{line_width}mm")
                    folder_label.config(text=f"已加载默认配置文件夹: {folder_path}")  # 显示加载后的路径
                    start_button.config(state=NORMAL)  # 启用"开始处理"按钮
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


def convert_rgb_to_cmyk_jpeg(input_tif, output_jpg):
    """
    使用 Photoshop 将 CMYK TIF 转换为 CMYK JPEG，并保持 CMYK 颜色空间
    :param input_tif: 输入的 TIF 文件路径
    :param output_jpg: 输出的 JPEG 文件路径
    """
    # 启动 Photoshop
    psApp = win32com.client.Dispatch("Photoshop.Application")
    psApp.DisplayDialogs = 3  # 设为静默模式，不弹出对话框

    # 打开 TIF 文件
    doc = psApp.Open(input_tif)

    # 确保文档颜色模式为 CMYK
    if doc.Mode != 3:  # 3 = psCMYKMode
        doc.ChangeMode(3)  # 转为 CMYK
        doc.Save()

    # 设置 JPEG 保存选项
    options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
    options.Quality = 12  # 最高质量 (1-12)
    options.Matte = 1  # 1 = psNoMatte，保持透明区域

    # 保存为 JPEG
    doc.SaveAs(output_jpg, options, True)

    # 关闭文档
    doc.Close()

def draw_lines_on_image(image, draw_line_color, horizontal_offset_cm=7, dpi=72):
    """ 在图片上方指定厘米处绘制水平线，并在中央绘制垂直线 """
    draw = ImageDraw.Draw(image)
    width, height = image.size
    horizontal_offset_px = cm_to_pixels(horizontal_offset_cm, dpi)
    y_horizontal = min(horizontal_offset_px, height - 1)
    x_vertical = width // 2
    line_width_px = mm_to_pixels(line_width, dpi)
    line_width_px = math.ceil(line_width_px) if line_width_px - math.floor(line_width_px) >= 0.5 else math.floor(line_width_px)
    
    # 画水平线 (从 (0, y) 到 (width, y))
    draw.line([(0, y_horizontal), (width, y_horizontal)], fill=draw_line_color, width=line_width_px)
    # 画垂直线 (从 (x, 0) 到 (x, height))
    draw.line([(x_vertical, 0), (x_vertical, height)], fill=draw_line_color, width=line_width_px)
    
    return image

def draw_holes_on_image(image, hole_count=6, hole_diameter_cm=1, margin_cm=2, dpi=72):
    """在图片上绘制打孔点，保证左右上下对称、间距均匀"""
    draw = ImageDraw.Draw(image)
    width_px, height_px = image.size
    width_cm = width_px * 2.54 / dpi
    height_cm = height_px * 2.54 / dpi

    hole_radius_cm = hole_diameter_cm / 2
    hole_radius_px = cm_to_pixels(hole_radius_cm, dpi)

    # 上下行数量
    if hole_count == 6:
        per_row = 3
    elif hole_count == 8:
        per_row = 4
    else:
        raise ValueError("打孔数量只能是6或8")

    # === X方向：均匀分布（左右留边 + 半径） ===
    x1_cm = margin_cm + hole_radius_cm
    xN_cm = width_cm - margin_cm - hole_radius_cm
    spacing_cm = (xN_cm - x1_cm) / (per_row - 1) if per_row > 1 else 0
    x_positions_px = [cm_to_pixels(x1_cm + i * spacing_cm, dpi) for i in range(per_row)]

    # === Y方向：上下边距同理（加半径） ===
    top_y_px = cm_to_pixels(height_cm - margin_cm - hole_radius_cm, dpi)
    bottom_y_px = cm_to_pixels(margin_cm + hole_radius_cm, dpi)

    # 绘制红色圆点（顶部+底部）
    for y in [top_y_px, bottom_y_px]:
        for x in x_positions_px:
            draw.ellipse(
                [x - hole_radius_px, y - hole_radius_px, x + hole_radius_px, y + hole_radius_px],
                fill='red', outline='red'
            )

    return image

def mm_to_pixels(mm_value, dpi):
    """将毫米转换为像素"""
    return mm_value * (dpi / 25.4)

@com_thread
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

        jpg_seq = 1
        for filename in filenames:
            if stop_processing:
                write_log("🚫 停止信号收到，提前终止图片处理")
                return

            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif')):
                image_path = os.path.join(current_folder, filename)
                time.sleep(1)

                try:
                    with Image.open(image_path) as image:
                        write_log(f"✅ 图片 '{image_path}' 开始处理...")

                        # 调整尺寸
                        resized_image = image.resize((target_width, target_height), Image.LANCZOS)
                        write_log(f"✅ 第一步：尺寸调整成功...")

                        # 绘制线条
                        if draw_lines.get() == True:
                            draw_line_color = ""
                            if draw_lines_color_1.get() == True:
                                draw_line_color = "white"
                            if draw_lines_color_2.get() == True:
                                draw_line_color = "gray"
                            if draw_lines_color_3.get() == True:
                                draw_line_color = "black"
                            write_log(f"✅ 第二步：画线开始, 线条颜色: {draw_line_color}, 线条宽度: {line_width}, 画线偏移量: {selected_horizontal_offset.get()}CM...")
                            resized_image = draw_lines_on_image(resized_image, draw_line_color, horizontal_offset_cm=int(selected_horizontal_offset.get()), dpi=72)
                            write_log(f"✅ 第二步：画线成功...")
                        else:
                            write_log(f"✅ 第二步：不画线, 跳过...")

                        # 绘制打孔点
                        if draw_holes.get() == True:
                            try:
                                hole_count = int(hole_count_var.get())
                                hole_diameter = float(hole_diameter_entry.get())
                                hole_margin = float(hole_margin_entry.get())
                                write_log(f"✅ 第三步：打孔开始, 打孔数量: {hole_count}, 孔直径: {hole_diameter}cm, 边距: {hole_margin}cm...")
                                resized_image = draw_holes_on_image(resized_image, hole_count, hole_diameter, hole_margin, dpi=72)
                                write_log(f"✅ 第三步：打孔成功...")
                            except Exception as e:
                                write_log(f"❌ 打孔失败: {e}")
                        else:
                            write_log(f"✅ 第三步：不打孔, 跳过...")

                        tif_image_path = os.path.splitext(image_path)[0] + ".tif"
                        # 以无损 LZW 压缩方式保存为 TIF
                        resized_image.save(tif_image_path, "TIFF", compression="tiff_lzw")
                        write_log(f"✅ 第四步：保存调整尺寸后的图片成功...")

                        jpg_image_path = os.path.splitext(image_path)[0] + "(" + folder_name + ")" + ".jpg"
                        convert_rgb_to_cmyk_jpeg(tif_image_path, jpg_image_path)
                        write_log(f"✅ 第五步：调用PS -> 图片转CMYK模式成功, 文件保存到本地成功...")
                        jpg_seq += 1

                        # 如果原文件不是jpg，则删除原文件
                        os.remove(tif_image_path)
                        os.remove(image_path)
                        write_log(f"✅ 第六步：删除原图片文件成功...")
                        write_log(f"✅ 图片处理完成！！！")

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

# =============== 程序入口 ===============
# 检查试用期
if not check_trial_period():
    exit(0)

# GUI界面
root = Tk()
root.title("自动调图软件V2.0_试用版_20251028")
root.geometry("900x800")

folder_button = Button(root, text="选择文件夹", command=browse_folder)
folder_button.pack(pady=10)

folder_label = Label(root, text="请选择文件夹")
folder_label.pack()


# 画线设置
line_frame = Frame(root)
line_frame.pack(pady=10)

draw_lines = BooleanVar(root)
check_button = Checkbutton(line_frame, text="是否绘制线条", variable=draw_lines)
check_button.pack(side="left", padx=5)

# 水平偏移量设置
offset_label = Label(line_frame, text="请输入上方水平画线偏移量（CM）:")
offset_label.pack(side="left", padx=5)
selected_horizontal_offset = StringVar()
selected_horizontal_offset.set(horizontal_offset_options[0])
offset_entry = Entry(line_frame, textvariable=selected_horizontal_offset, width=5)
offset_entry.pack(side="left", padx=5)

# 线条颜色设置
color_frame = Frame(root)
color_frame.pack(pady=10)

draw_lines_color_1 = BooleanVar(root)
check_button1 = Checkbutton(color_frame, text="线条颜色-白色", variable=draw_lines_color_1)
check_button1.pack(side="left", padx=5)

draw_lines_color_2 = BooleanVar(root)
check_button2 = Checkbutton(color_frame, text="线条颜色-灰色", variable=draw_lines_color_2)
check_button2.pack(side="left", padx=5)

draw_lines_color_3 = BooleanVar(root)
check_button3 = Checkbutton(color_frame, text="线条颜色-黑色", variable=draw_lines_color_3)
check_button3.pack(side="left", padx=5)

# 打孔设置
hole_frame = Frame(root)
hole_frame.pack(pady=10)

draw_holes = BooleanVar(root)
hole_check_button = Checkbutton(hole_frame, text="是否绘制打孔点", variable=draw_holes)
hole_check_button.pack(side="left", padx=5)

# 打孔数量
hole_count_label = Label(hole_frame, text="打孔数量:")
hole_count_label.pack(side="left", padx=5)
hole_count_var = StringVar(value="6")
hole_count_frame = Frame(hole_frame)
hole_count_frame.pack(side="left")
Radiobutton(hole_count_frame, text="6个", variable=hole_count_var, value="6").pack(side="left")
Radiobutton(hole_count_frame, text="8个", variable=hole_count_var, value="8").pack(side="left")

# 打孔参数设置
hole_params_frame = Frame(root)
hole_params_frame.pack(pady=10)

hole_diameter_label = Label(hole_params_frame, text="孔直径(cm):")
hole_diameter_label.pack(side="left", padx=5)
hole_diameter_entry = Entry(hole_params_frame, width=6)
hole_diameter_entry.insert(0, "1")
hole_diameter_entry.pack(side="left")

hole_margin_label = Label(hole_params_frame, text="边距(cm):")
hole_margin_label.pack(side="left", padx=5)
hole_margin_entry = Entry(hole_params_frame, width=6)
hole_margin_entry.insert(0, "1.5")
hole_margin_entry.pack(side="left")

# 控制按钮
button_frame = Frame(root)
button_frame.pack(pady=10)

start_button = Button(button_frame, text="开始处理", state=DISABLED, command=start_threaded_processing)
start_button.pack(side="left", padx=5)

stop_button = Button(button_frame, text="停止处理", command=stop_processing_function)
stop_button.pack(side="left", padx=5)
stop_button.config(state="disabled")

# 日志显示
log_text = scrolledtext.ScrolledText(root, width=100, height=25, wrap=WORD)
log_text.pack()

update_log_window()

setup_logging()
load_config()

root.mainloop()
