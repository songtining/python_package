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
from PIL import Image, ImageCms, ImageDraw
import math
import win32com.client
Image.MAX_IMAGE_PIXELS = 1000000000  # 设置为5亿像素，适应你的大图

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

def convert_rgb_to_cmyk(image):
    cmyk_profile = ImageCms.getOpenProfile("CMYK.icc")
    # srgb_profile = ImageCms.createProfile("sRGB.icc")
    srgb_profile = ImageCms.getOpenProfile("sRGB.icc")
    return ImageCms.profileToProfile(image, srgb_profile, cmyk_profile, outputMode="CMYK")

def process_images_in_folder(root_folder):
    global stop_processing

    write_log(f"📏 扫描根目录: {root_folder} 开始 ")

    # 初始化 Photoshop 实例（仅一次）
    ps_app = win32com.client.Dispatch("Photoshop.Application")
    ps_app.DisplayDialogs = 2  # 全局静默模式
    success_count = 0

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
                        # image = image.convert("RGB")

                        if image.size == (target_width, target_height):
                            write_log(f"✅ 图片 '{image_path}' 尺寸已符合要求，跳过")
                            continue

                        write_log(f"✅ 图片 '{image_path}' 开始处理...")

                        # 调整尺寸
                        resized_image = image.resize((target_width, target_height), Image.LANCZOS)
                        write_log(f"✅ 第一步：尺寸调整成功...")

                        if (draw_lines.get() == True):
                            draw_line_color = ""
                            if draw_lines_color_1.get() == True:
                                draw_line_color = "white"
                            if draw_lines_color_2.get() == True:
                                draw_line_color = "gray"
                            if draw_lines_color_3.get() == True:
                                draw_line_color = "black"
                            write_log(f"✅ 第二步：画线开始, 线条颜色: {draw_line_color}, 线条宽度: {line_width}, 画线偏移量: {selected_horizontal_offset.get()}CM...")
                            # 画线
                            resized_image = draw_lines_on_image(resized_image, draw_line_color, horizontal_offset_cm=int(selected_horizontal_offset.get()), dpi=72)
                            write_log(f"✅ 第二步：画线成功...")
                        else:
                            write_log(f"✅ 第二步：不画线, 跳过...")

                        tif_image_path = os.path.splitext(image_path)[0] + ".tif"
                        # 以无损 LZW 压缩方式保存为 TIF
                        resized_image.save(tif_image_path, "TIFF", compression="tiff_lzw")
                        write_log(f"✅ 第三步：保存调整尺寸后的图片成功...")

                        jpg_image_path = os.path.splitext(image_path)[0] + "(" + folder_name + ")" + ".jpg"
                        convert_rgb_to_cmyk_jpeg(tif_image_path, jpg_image_path, ps_app)
                        write_log(f"✅ 第四步：调用PS -> 图片转CMYK模式成功, 文件保存到本地成功...")
                        jpg_seq += 1

                        # 如果原文件不是jpg，则删除原文件
                        os.remove(tif_image_path)
                        os.remove(image_path)
                        write_log(f"✅ 第五步：删除原图片文件成功...")
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

def draw_lines_on_image(image, draw_line_color, horizontal_offset_cm=7, dpi=72):
    """ 在图片上方指定厘米处绘制水平线，并在中央绘制垂直线 """

    # 打开图片
    # image = Image.open(image_path)
    draw = ImageDraw.Draw(image)

    # 获取图片尺寸
    width, height = image.size

    # 计算水平线位置（7cm 对应的像素）
    horizontal_offset_px = cm_to_pixels(horizontal_offset_cm, dpi)

    # 确保线条不会超出图片范围
    y_horizontal = min(horizontal_offset_px, height - 1)
    x_vertical = width // 2

    # 将0.1毫米转换为像素
    line_width_px = mm_to_pixels(line_width, dpi)
    # 四舍五入取整，因为线条宽度一般为整数像素
    line_width_px = math.ceil(line_width_px) if line_width_px - math.floor(line_width_px) >= 0.5 else math.floor(line_width_px)

    # 画水平线 (从 (0, y) 到 (width, y))
    draw.line([(0, y_horizontal), (width, y_horizontal)], fill=draw_line_color, width=line_width_px)

    # 画垂直线 (从 (x, 0) 到 (x, height))
    draw.line([(x_vertical, 0), (x_vertical, height)], fill=draw_line_color, width=line_width_px)

    return image

    # 保存新图片
    # image.save(output_path)
    # write_log(f"✅ 图片画线完成'{image_path}'")

def mm_to_pixels(mm_value, dpi):
    """将毫米转换为像素"""
    return mm_value * (dpi / 25.4)


def convert_rgb_to_cmyk_jpeg(input_tif, output_jpg, ps_app=None):
    """
    使用已初始化的 Photoshop 实例处理单张图片
    :param input_tif: 输入 TIF 文件路径
    :param output_jpg: 输出 JPEG 文件路径
    :param ps_app: 预初始化的 Photoshop 实例（可选）
    :return: 是否处理成功
    """
    # 初始化 Photoshop（仅当未传入实例时）
    if ps_app is None:
        ps_app = win32com.client.Dispatch("Photoshop.Application")
        ps_app.DisplayDialogs = 2  # 静默模式（关键！原代码是3）
        ps_app.UserPreferences = {
            "rasterizationColorSpace": 3,  # 确保栅格化颜色空间为CMYK
            "jpgQuality": 12               # 默认JPEG质量
        }

    try:
        # 打开文档并强制关闭旧文档（防残留）
        if ps_app.Documents.Count > 0:
            ps_app.Documents.Close(SaveChanges=False)
        doc = ps_app.Open(input_tif)
        if not doc:
            print(f"无法打开文件: {input_tif}")
            return False

        # 强制转换为CMYK模式（如果输入非CMYK）
        if doc.Mode != 3:  # 3 = psCMYKMode
            doc.ChangeMode(3)
            doc.Save()  # 保存模式变更

        # JPEG保存配置（复用对象提升性能）
        options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        options.Quality = ps_app.UserPreferences.get("jpgQuality", 12)
        options.Matte = 1  # 无蒙版（透明区域填充白色）

        # 处理输出路径（避免文件占用）
        # output_dir = os.path.dirname(output_jpg)
        # os.makedirs(output_dir, exist_ok=True)
        # if os.path.exists(output_jpg):
        #     os.remove(output_jpg)

        # 保存并关闭文档（不退出PS）
        doc.SaveAs(output_jpg, options, overwrite=True)
        doc.Close(SaveChanges=False)
        return True

    except Exception as e:
        print(f"处理 {input_tif} 失败: {str(e)}")
        return False
    finally:
        # 防御性关闭文档（确保资源释放）
        if 'doc' in locals() and doc.IsOpen:
            doc.Close(SaveChanges=False)

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

# GUI界面
root = Tk()
root.title("自动调图软件V1.0")
root.geometry("800x600")

folder_button = Button(root, text="选择文件夹", command=browse_folder)
folder_button.pack(pady=10)

folder_label = Label(root, text="请选择文件夹")
folder_label.pack()

# 新增：水平偏移量选择项
selected_horizontal_offset = StringVar()
selected_horizontal_offset.set(horizontal_offset_options[0])  # 默认选择7CM

# 创建一个 Frame 容器（用于存放同一行的组件）
frame = Frame(root)
frame.pack(pady=10)  # 设置一点垂直间距

draw_lines = BooleanVar(root)  # 记录是否绘制线条，默认不绘制
# 复选框（是否绘制线条）
check_button = Checkbutton(frame, text="是否绘制线条", variable=draw_lines)
check_button.pack(side="left", padx=5)  # `side="left"` 让它放在左侧

# 创建输入框代替下拉选择框
offset_label = Label(frame, text="请输入上方水平画线偏移量（CM）:")
offset_label.pack(side="left", padx=5)
offset_entry = Entry(frame, textvariable=selected_horizontal_offset, width=5)
offset_entry.pack(side="left", padx=5)

# 创建一个 Frame 容器（用于存放同一行的组件）
frame3 = Frame(root)
frame3.pack(pady=10)  # 设置一点垂直间距
draw_lines_color_1 = BooleanVar(root)
check_button1 = Checkbutton(frame3, text="线条颜色-白色", variable=draw_lines_color_1)
check_button1.pack(side="left", padx=5)  # `side="left"` 让它放在左侧

draw_lines_color_2 = BooleanVar(root)
check_button2 = Checkbutton(frame3, text="线条颜色-灰色", variable=draw_lines_color_2)
check_button2.pack(side="left", padx=5)  # `side="left"` 让它放在左侧

draw_lines_color_3 = BooleanVar(root)
check_button3 = Checkbutton(frame3, text="线条颜色-黑色", variable=draw_lines_color_3)
check_button3.pack(side="left", padx=5)  # `side="left"` 让它放在左侧

frame2 = Frame(root)
frame2.pack(pady=10)  # 设置一点垂直间距
start_button = Button(frame2, text="开始处理", state=DISABLED, command=start_threaded_processing)
start_button.pack(side="left", padx=5)

stop_button = Button(frame2, text="停止处理", command=stop_processing_function)
stop_button.pack(side="left", padx=5)
stop_button.config(state="disabled")

log_text = scrolledtext.ScrolledText(root, width=90, height=30, wrap=WORD)
log_text.pack()

update_log_window()

setup_logging()
load_config()

root.mainloop()
