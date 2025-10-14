#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
from pathlib import Path
from PIL import Image, ImageDraw
import win32com.client
import sys
import pythoncom
import functools

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

# =============== 工具函数 ===============

def parse_folder_dimensions(folder_name):
    """从文件夹名称解析尺寸，支持多种格式：
    - 30x40cm, 30x40, 30*40cm, 30*40
    - 30x40CM, 30X40cm 等
    """
    # 清理文件夹名称
    name = folder_name.strip()
    
    # 尝试多种匹配模式
    patterns = [
        r'(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*cm?',  # 30x40cm, 30*40CM
        r'(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)',        # 30x40, 30*40
    ]
    
    for pattern in patterns:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            width = float(match.group(1))
            height = float(match.group(2))
            return width, height
    
    return None, None

def cm_to_px(cm, dpi=300):
    """厘米转像素"""
    return int(round(cm * dpi / 2.54))

def draw_holes(image, hole_count=6, hole_diameter_cm=1, margin_cm=2, dpi=300):
    """在图片上绘制打孔点，保证左右对称、间距均匀"""
    draw = ImageDraw.Draw(image)
    width_px, height_px = image.size
    width_cm = width_px * 2.54 / dpi
    height_cm = height_px * 2.54 / dpi

    hole_radius_cm = hole_diameter_cm / 2
    hole_radius_px = cm_to_px(hole_radius_cm, dpi)
    margin_px = cm_to_px(margin_cm, dpi)

    # 上下行数量
    if hole_count == 6:
        per_row = 3
    elif hole_count == 8:
        per_row = 4
    else:
        raise ValueError("打孔数量只能是6或8")

    # === 均匀分布算法（你想要的逻辑） ===
    x1_cm = margin_cm + hole_radius_cm
    xN_cm = width_cm - margin_cm - hole_radius_cm
    if per_row > 1:
        spacing_cm = (xN_cm - x1_cm) / (per_row - 1)
    else:
        spacing_cm = 0

    x_positions_px = [cm_to_px(x1_cm + i * spacing_cm, dpi) for i in range(per_row)]

    # 顶部和底部 y 坐标
    top_y_px = margin_px
    bottom_y_px = height_px - margin_px

    # 绘制红色圆点（顶部+底部）
    for y in [top_y_px, bottom_y_px]:
        for x in x_positions_px:
            draw.ellipse(
                [x - hole_radius_px, y - hole_radius_px, x + hole_radius_px, y + hole_radius_px],
                fill='red', outline='red'
            )

    return image

def get_photoshop_app(log_func=print):
    """健壮获取 Photoshop COM 对象"""
    if not sys.platform.startswith('win'):
        raise RuntimeError("当前系统不是 Windows，无法使用 Photoshop COM 接口")

    progids = [
        # 通用/较新版本
        "Photoshop.Application",
        "Photoshop.Application.2025",
        "Photoshop.Application.2024",
        "Photoshop.Application.2023",
        "Photoshop.Application.2022",
        # 旧版本/CS 系列
        "Photoshop.Application.CS6",
        "Photoshop.Application.60",
    ]
    last_err = None
    for pid in progids:
        try:
            log_func(f"尝试使用 ProgID: {pid}")
            try:
                app = win32com.client.gencache.EnsureDispatch(pid)
            except Exception:
                app = win32com.client.Dispatch(pid)
            # 设置静默模式
            try:
                app.DisplayDialogs = 3
            except Exception:
                pass
            return app
        except Exception as e:
            last_err = e
    raise RuntimeError(f"无法启动 Photoshop COM，请确认已安装并可正常启动。原始错误: {last_err}")

def convert_to_cmyk(input_path, output_path, ps_app=None, log_func=print):
    """使用Photoshop转换为CMYK格式"""
    try:
        input_path = str(Path(input_path).resolve())
        output_path = str(Path(output_path).resolve())

        if ps_app is None:
            log_func("🚀 启动 Photoshop...")
            ps_app = get_photoshop_app(log_func)

        log_func(f"➡ 打开文件: {input_path}")
        doc = ps_app.Open(input_path)

        if doc is None:
            log_func(f"❌ 无法打开文件: {input_path}")
            return False

        # 确保转换为 CMYK
        if doc.Mode != 3:  # 3 = psCMYKMode
            log_func("🎨 转换为 CMYK 模式")
            doc.ChangeMode(3)
            doc.Save()

        # JPEG 保存选项
        log_func(f"💾 保存为 JPEG: {output_path}")
        options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        options.Quality = 12
        options.Matte = 1

        doc.SaveAs(output_path, options, True)
        doc.Close()
        log_func(f"✅ CMYK 转换完成: {output_path}")
        return True

    except Exception as e:
        log_func(f"❌ CMYK 转换失败: {str(e)}")
        return False

# =============== 主应用 ===============

class ImageHoleProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片打孔处理工具 V1.0")
        self.root.geometry("800x700")

        self.stop_flag = False
        self.psApp = None

        # 输入输出目录
        row1 = tk.Frame(root)
        row1.pack(fill="x", padx=10, pady=6)
        tk.Label(row1, text="输入目录:").pack(side="left")
        self.in_entry = tk.Entry(row1)
        self.in_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row1, text="选择目录", command=self.choose_in_dir).pack(side="right")

        row2 = tk.Frame(root)
        row2.pack(fill="x", padx=10, pady=6)
        tk.Label(row2, text="输出目录:").pack(side="left")
        self.out_entry = tk.Entry(row2)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row2, text="选择目录", command=self.choose_out_dir).pack(side="right")

        # 参数设置
        row3 = tk.Frame(root)
        row3.pack(fill="x", padx=10, pady=6)
        
        tk.Label(row3, text="DPI:").pack(side="left")
        self.dpi_entry = tk.Entry(row3, width=6)
        self.dpi_entry.insert(0, "300")
        self.dpi_entry.pack(side="left", padx=4)
        
        tk.Label(row3, text="打孔数量:").pack(side="left", padx=(16, 4))
        self.hole_count_var = tk.StringVar(value="6")
        hole_frame = tk.Frame(row3)
        hole_frame.pack(side="left")
        tk.Radiobutton(hole_frame, text="6个", variable=self.hole_count_var, value="6").pack(side="left")
        tk.Radiobutton(hole_frame, text="8个", variable=self.hole_count_var, value="8").pack(side="left")
        
        tk.Label(row3, text="圆点直径(cm):").pack(side="left", padx=(16, 4))
        self.diameter_entry = tk.Entry(row3, width=6)
        self.diameter_entry.insert(0, "1")
        self.diameter_entry.pack(side="left")
        
        tk.Label(row3, text="边距(cm):").pack(side="left", padx=(16, 4))
        self.margin_entry = tk.Entry(row3, width=6)
        self.margin_entry.insert(0, "2")
        self.margin_entry.pack(side="left")

        # 是否转换为 CMYK
        row4 = tk.Frame(root)
        row4.pack(fill="x", padx=10, pady=6)
        self.cmyk_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row4, text="转换为CMYK模式", variable=self.cmyk_var).pack(side="left")

        # 按钮
        row5 = tk.Frame(root)
        row5.pack(fill="x", padx=10, pady=10)
        tk.Button(row5, text="开始处理", command=self.start).pack(side="left", padx=6)
        tk.Button(row5, text="停止处理", command=self.stop).pack(side="left", padx=6)

        # 日志
        tk.Label(root, text="处理日志:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=15)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # 进度条
        tk.Label(root, text="进度:").pack(anchor="w", padx=10)
        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=600)
        self.progress.pack(fill="x", padx=10, pady=8)

    def choose_in_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.in_entry.delete(0, tk.END)
            self.in_entry.insert(0, p)

    def choose_out_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, p)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def stop(self):
        self.stop_flag = True
        self.log("⚠️ 用户请求停止...")

    def start(self):
        in_dir = Path(self.in_entry.get().strip())
        out_dir = Path(self.out_entry.get().strip())
        
        if not in_dir or not out_dir:
            messagebox.showwarning("提示", "请先选择输入目录和输出目录")
            return

        try:
            dpi = int(self.dpi_entry.get().strip() or "300")
        except:
            dpi = 300
            
        try:
            hole_count = int(self.hole_count_var.get())
        except:
            hole_count = 6
            
        try:
            diameter = float(self.diameter_entry.get().strip() or "1")
        except:
            diameter = 1.0
            
        try:
            margin = float(self.margin_entry.get().strip() or "2")
        except:
            margin = 2.0

        self.stop_flag = False
        
        # 启动处理线程
        t = threading.Thread(target=self.process_images,
                           args=(in_dir, out_dir, dpi, hole_count, diameter, margin),
                           daemon=True)
        t.start()

    @com_thread
    def process_images(self, in_dir, out_dir, dpi, hole_count, diameter, margin):
        """处理图片的主函数"""
        self.log("🚀 开始处理图片...")
        
        # 获取所有子文件夹
        folders = [f for f in in_dir.iterdir() if f.is_dir()]
        if not folders:
            self.log("❌ 输入目录中没有找到子文件夹")
            return
        
        total_folders = len(folders)
        processed_folders = 0
        
        for folder in folders:
            if self.stop_flag:
                self.log("🚫 用户请求停止处理")
                break
                
            try:
                self.log(f"📁 处理文件夹: {folder.name}")
                
                # 解析文件夹名称获取尺寸
                width_cm, height_cm = parse_folder_dimensions(folder.name)
                if width_cm is None or height_cm is None:
                    self.log(f"⚠️ 无法从文件夹名称解析尺寸: {folder.name}")
                    continue
                
                self.log(f"📏 解析尺寸: {width_cm}cm x {height_cm}cm")
                
                # 获取文件夹中的所有图片
                image_files = []
                for ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                    image_files.extend(folder.glob(f'*{ext}'))
                    image_files.extend(folder.glob(f'*{ext.upper()}'))
                
                if not image_files:
                    self.log(f"⚠️ 文件夹中没有找到图片: {folder.name}")
                    continue
                
                self.log(f"🖼️ 找到 {len(image_files)} 张图片")
                
                # 创建输出文件夹
                output_folder = out_dir / folder.name
                output_folder.mkdir(parents=True, exist_ok=True)
                
                # 处理每张图片
                for img_file in image_files:
                    if self.stop_flag:
                        break
                        
                    try:
                        self.log(f"🔄 处理图片: {img_file.name}")
                        
                        # 打开图片
                        with Image.open(img_file) as img:
                            # 转换为RGB模式
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            
                            # 调整尺寸
                            target_width = cm_to_px(width_cm, dpi)
                            target_height = cm_to_px(height_cm, dpi)
                            resized_img = img.resize((target_width, target_height), Image.LANCZOS)
                            
                            # 绘制打孔点
                            hole_img = draw_holes(resized_img, hole_count, diameter, margin, dpi)
                            
                            # 保存RGB版本
                            rgb_output = output_folder / f"{img_file.stem}_rgb.jpg"
                            hole_img.save(rgb_output, "JPEG", quality=95, dpi=(dpi, dpi))
                            self.log(f"✅ 已保存RGB版本: {rgb_output.name}")
                            
                            # 如果需要CMYK转换
                            if self.cmyk_var.get():
                                try:
                                    if self.psApp is None:
                                        self.log("🚀 启动 Photoshop...")
                                        self.psApp = get_photoshop_app(self.log)
                                    
                                    cmyk_output = output_folder / f"{img_file.stem}_cmyk.jpg"
                                    convert_to_cmyk(rgb_output, cmyk_output, self.psApp, self.log)
                                    
                                except Exception as e:
                                    self.log(f"❌ CMYK转换失败: {e}")
                    
                    except Exception as e:
                        self.log(f"❌ 处理图片失败 {img_file.name}: {e}")
                
                processed_folders += 1
                self.progress["value"] = int(processed_folders * 100 / total_folders)
                self.root.update_idletasks()
                
            except Exception as e:
                self.log(f"❌ 处理文件夹失败 {folder.name}: {e}")
        
        self.log("🎉 所有图片处理完成！")

# =============== 入口 ===============
if __name__ == "__main__":
    root = tk.Tk()
    app = ImageHoleProcessorApp(root)
    root.mainloop()
