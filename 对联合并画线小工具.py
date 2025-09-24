#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
from pathlib import Path
from PIL import Image, ImageDraw
import win32com.client

# =============== 工具函数 ===============

def parse_pair_name(filename: str):
    stem = Path(filename).stem.strip()
    m = re.match(r'^(?P<key>.+?)\s*\((?P<part>[12])\)_?(?P<sub>\d+)?$', stem)
    if m: return m.group("key"), int(m.group("part")), int(m.group("sub") or 0)
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])_?(?P<sub>\d+)?$', stem)
    if m: return m.group("key"), int(m.group("part")), int(m.group("sub") or 0)
    m = re.match(r'^(?P<key>.+?)\s*\((?P<part>[12])\)$', stem)
    if m: return m.group("key"), int(m.group("part")), 0
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])$', stem)
    if m: return m.group("key"), int(m.group("part")), 0
    m = re.match(r'^(?P<key>.+?)_(?P<sub>\d+)$', stem)
    if m: return m.group("key"), None, int(m.group("sub"))
    return stem, None, 0

def get_image_dpi(img: Image.Image, default_dpi=300):
    dpi = img.info.get('dpi')
    if isinstance(dpi, tuple) and len(dpi) >= 2:
        return dpi[0] or default_dpi, dpi[1] or default_dpi
    return default_dpi, default_dpi

def cm_to_px(cm: float, dpi: float) -> int:
    return int(round(cm * dpi / 2.54))

def draw_guides(img: Image.Image, top_cm=2.5, line_width=3,
                color=(255, 255, 255), default_dpi=300):
    draw = ImageDraw.Draw(img)
    dpi_x, dpi_y = get_image_dpi(img, default_dpi)
    y = cm_to_px(top_cm, dpi_y)
    y = max(0, min(img.height - 1, y))
    x = img.width // 2
    draw.line([(0, y), (img.width, y)], fill=color, width=line_width)
    draw.line([(x, 0), (x, img.height)], fill=color, width=line_width)
    return img

def format_cm(value: float) -> str:
    value = round(value, 1)
    if value.is_integer(): return str(int(value))
    return f"{value:.1f}"

def get_size_cm(img: Image.Image, default_dpi=300, top_margin_cm=0.0):
    dpi_x, dpi_y = get_image_dpi(img, default_dpi)
    w_cm = img.width * 2.54 / dpi_x
    h_cm = img.height * 2.54 / dpi_y - top_margin_cm
    return format_cm(w_cm), format_cm(h_cm)

def ensure_folder(p: Path):
    p.mkdir(parents=True, exist_ok=True)
    return p

def resize_to_target(img: Image.Image, target_w_cm: float,
                     target_h_cm: float, dpi: int):
    target_w = cm_to_px(target_w_cm, dpi)
    target_h = cm_to_px(target_h_cm, dpi)
    return img.resize((target_w, target_h), Image.LANCZOS)

# =============== Photoshop 转换函数 ===============
def convert_rgb_to_cmyk_jpeg(input_jpg, output_jpg, ps_app=None, log_func=print):
    try:
        if ps_app is None:
            log_func("🚀 启动 Photoshop...")
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            ps_app.DisplayDialogs = 2  # 静默模式

        log_func(f"➡ 打开文件: {input_jpg}")
        doc = ps_app.Open(str(input_jpg))

        if doc is None:
            log_func(f"❌ 无法打开文件: {input_jpg}")
            return False

        if doc.Mode != 3:  # 3 = psCMYKMode
            log_func("🎨 转换为 CMYK 模式")
            doc.ChangeMode(3)
            doc.Save()

        log_func(f"💾 保存为 JPEG: {output_jpg}")
        options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        options.Quality = 12
        options.Matte = 1
        doc.SaveAs(str(output_jpg), options, True)

        doc.Close(SaveChanges=False)
        log_func(f"✅ CMYK 转换完成: {output_jpg}")
        return True

    except Exception as e:
        log_func(f"❌ CMYK 转换失败: {str(e)}")
        return False

# =============== 主应用 ===============

class CoupletProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片合并处理小工具 V1.3")
        self.root.geometry("1100x750")

        self.stop_flag = False
        self.psApp = None

        # 输入输出目录
        row1 = tk.Frame(root); row1.pack(fill="x", padx=10, pady=6)
        tk.Label(row1, text="输入目录:").pack(side="left")
        self.in_entry = tk.Entry(row1); self.in_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row1, text="选择目录", command=self.choose_in_dir).pack(side="right")

        row2 = tk.Frame(root); row2.pack(fill="x", padx=10, pady=6)
        tk.Label(row2, text="输出目录:").pack(side="left")
        self.out_entry = tk.Entry(row2); self.out_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row2, text="选择目录", command=self.choose_out_dir).pack(side="right")

        # 选项
        row3 = tk.Frame(root); row3.pack(fill="x", padx=10, pady=6)
        self.merge_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row3, text="先合并上下联（场景1）",
                       variable=self.merge_var).pack(side="left", padx=4)

        tk.Label(row3, text="DPI:").pack(side="left", padx=(16, 4))
        self.dpi_entry = tk.Entry(row3, width=6); self.dpi_entry.insert(0, "300"); self.dpi_entry.pack(side="left")

        tk.Label(row3, text="横线距离(cm):").pack(side="left", padx=(16, 4))
        self.cm_entry = tk.Entry(row3, width=6); self.cm_entry.insert(0, "2.5"); self.cm_entry.pack(side="left")

        tk.Label(row3, text="线宽(px):").pack(side="left", padx=(16, 4))
        self.width_entry = tk.Entry(row3, width=6); self.width_entry.insert(0, "2"); self.width_entry.pack(side="left")

        tk.Label(row3, text="目标宽(cm):").pack(side="left", padx=(16, 4))
        self.target_w_entry = tk.Entry(row3, width=6); self.target_w_entry.insert(0, "30"); self.target_w_entry.pack(side="left")

        tk.Label(row3, text="目标高(cm):").pack(side="left", padx=(16, 4))
        self.target_h_entry = tk.Entry(row3, width=6); self.target_h_entry.insert(0, "180"); self.target_h_entry.pack(side="left")

        # 是否转换为 CMYK
        self.cmyk_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row3, text="是否转换为CMYK模式",
                       variable=self.cmyk_var).pack(side="left", padx=(20, 4))

        # 按钮
        row4 = tk.Frame(root); row4.pack(fill="x", padx=10, pady=10)
        tk.Button(row4, text="开始处理", command=self.start).pack(side="left", padx=6)
        tk.Button(row4, text="停止处理", command=self.stop).pack(side="left", padx=6)

        # 日志
        tk.Label(root, text="过程日志:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=20)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # 进度条
        tk.Label(root, text="进度:").pack(anchor="w", padx=10)
        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=1000)
        self.progress.pack(fill="x", padx=10, pady=8)

    # ---------- 工具函数 ----------
    def choose_in_dir(self):
        p = filedialog.askdirectory()
        if p: self.in_entry.delete(0, tk.END); self.in_entry.insert(0, p)

    def choose_out_dir(self):
        p = filedialog.askdirectory()
        if p: self.out_entry.delete(0, tk.END); self.out_entry.insert(0, p)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n"); self.log_text.see(tk.END)
        self.root.update_idletasks()

    def stop(self):
        self.stop_flag = True; self.log("⚠️ 用户请求停止...")

    # ---------- 启动 ----------
    def start(self):
        in_dir = Path(self.in_entry.get().strip())
        out_dir = Path(self.out_entry.get().strip())
        if not in_dir or not out_dir:
            messagebox.showwarning("提示", "请先选择输入目录和输出目录")
            return

        try: dpi = int(self.dpi_entry.get().strip() or "300")
        except: dpi = 300
        try: top_cm = float(self.cm_entry.get().strip() or "2.5")
        except: top_cm = 2.5
        try: line_w = int(self.width_entry.get().strip() or "3")
        except: line_w = 3
        try: target_w_cm = float(self.target_w_entry.get().strip() or "30")
        except: target_w_cm = 30
        try: target_h_cm = float(self.target_h_entry.get().strip() or "180")
        except: target_h_cm = 180

        self.stop_flag = False
        if self.psApp is None and self.cmyk_var.get():
            self.log("🚀 启动 Photoshop...")
            self.psApp = win32com.client.Dispatch("Photoshop.Application")
            self.psApp.DisplayDialogs = 2

        if self.merge_var.get():
            t = threading.Thread(target=self.process_pairs,
                                 args=(in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm),
                                 daemon=True)
        else:
            t = threading.Thread(target=self.process_single,
                                 args=(in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm),
                                 daemon=True)
        t.start()

    # ---------- 成对处理 ----------
    def process_pairs(self, in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm):
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]
        groups = {}
        for f in files:
            key, part, sub = parse_pair_name(f.name)
            groups.setdefault((key, sub), {})
            if part is None: part = 1 if 2 in groups[(key, sub)] else 2
            groups[(key, sub)][part] = f

        # 打印分组结果
        self.log("📂 文件分组结果：")
        unpaired = []
        for (key, sub), parts in groups.items():
            files_info = ", ".join([f"part{p}:{f.name}" for p, f in parts.items()])
            if 1 in parts and 2 in parts:
                self.log(f"  ✅ 配对完成 ▶ {key}_{sub} => {files_info}")
            else:
                self.log(f"  ❌ 未配对 ▶ {key}_{sub} => {files_info}")
                unpaired.append((key, sub, files_info))

        # 如果有未配对文件，单独打印总结
        if unpaired:
            self.log("⚠️ 未成对文件列表：")
            for key, sub, files_info in unpaired:
                self.log(f"   - {key}_{sub}: {files_info}")

        pairs = [(k, v) for k, v in groups.items() if 1 in v and 2 in v]
        total = len(pairs); done = 0

        for (key, sub), pair in pairs:
            if self.stop_flag: break
            try:
                img1 = resize_to_target(Image.open(pair[1]).convert("RGB"), target_w_cm, target_h_cm, dpi)
                img2 = resize_to_target(Image.open(pair[2]).convert("RGB"), target_w_cm, target_h_cm, dpi)

                merged_w = cm_to_px(target_w_cm * 2, dpi)
                merged_h = cm_to_px(target_h_cm + top_cm, dpi)
                merged = Image.new("RGB", (merged_w, merged_h), (255, 255, 255))

                offset_y = cm_to_px(top_cm, dpi)
                merged.paste(img1, (0, offset_y))
                merged.paste(img2, (img1.width, offset_y))
                draw_guides(merged, top_cm=top_cm, line_width=line_w, color=(128, 128, 128), default_dpi=dpi)

                w_cm, h_cm = get_size_cm(merged, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                out_name = f"{key}_{sub}.jpg" if sub != 0 else f"{key}.jpg"
                out_path = bucket_dir / out_name
                merged.save(out_path, format="JPEG", quality=95, dpi=(dpi, dpi))
                self.log(f"✅ 已输出: {out_path}")

                if self.cmyk_var.get():
                    bucket_dir_cmyk = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm_cmyk")
                    cmyk_path = bucket_dir_cmyk / out_name
                    convert_rgb_to_cmyk_jpeg(out_path, cmyk_path, self.psApp, self.log)

            except Exception as e:
                self.log(f"❌ {key}_{sub} 处理失败: {e}")
            done += 1
            self.progress["value"] = int(done * 100 / max(1, total))
            self.root.update_idletasks()
        self.log("🎉 成对合并任务完成")

    # ---------- 单图处理 ----------
    def process_single(self, in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm):
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]
        total = len(files); done = 0

        for p in files:
            if self.stop_flag: break
            try:
                img = resize_to_target(Image.open(p).convert("RGB"), target_w_cm, target_h_cm, dpi)
                merged_w = cm_to_px(target_w_cm, dpi)
                merged_h = cm_to_px(target_h_cm + top_cm, dpi)
                canvas = Image.new("RGB", (merged_w, merged_h), (255, 255, 255))
                offset_y = cm_to_px(top_cm, dpi)
                canvas.paste(img, (0, offset_y))
                draw_guides(canvas, top_cm=top_cm, line_width=line_w, color=(128, 128, 128), default_dpi=dpi)

                w_cm, h_cm = get_size_cm(canvas, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                out_path = bucket_dir / f"{p.stem}.jpg"
                canvas.save(out_path, dpi=(dpi, dpi))
                self.log(f"✅ 已输出: {out_path}")

                if self.cmyk_var.get():
                    bucket_dir_cmyk = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm_cmyk")
                    cmyk_path = bucket_dir_cmyk / out_path.name
                    convert_rgb_to_cmyk_jpeg(out_path, cmyk_path, self.psApp, self.log)

            except Exception as e:
                self.log(f"❌ {p.name} 处理失败: {e}")
            done += 1
            self.progress["value"] = int(done * 100 / max(1, total))
            self.root.update_idletasks()
        self.log("🎉 单图缩放任务完成")

# =============== 入口 ===============
if __name__ == "__main__":
    root = tk.Tk()
    app = CoupletProcessorApp(root)
    root.mainloop()
