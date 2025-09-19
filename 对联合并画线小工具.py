#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
from pathlib import Path
from PIL import Image, ImageDraw
import datetime
import sys

# =============== 试用期限制 ===============
TRIAL_EXPIRE_DATE = datetime.date(2025, 9, 20)  # 设置试用期截止日期

def check_trial():
    today = datetime.date.today()
    if today > TRIAL_EXPIRE_DATE:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror(
            "试用期已到期",
            f"您的试用期已于 {TRIAL_EXPIRE_DATE} 到期，无法继续使用。\n请联系管理员获取正式版本。"
        )
        sys.exit(0)

# 先检查试用期
check_trial()

# =============== 工具函数 ===============

def parse_pair_name(filename: str):
    """
    支持：
    1) xxx (1)_1.jpg / xxx (2)_1.jpg   -> key=xxx_sub, part=1/2
    2) xxx-1.jpg / xxx_2.jpg           -> key=xxx, part=1/2
    3) xxx (1).jpg / xxx (2).jpg       -> key=xxx, part=1/2
    """
    stem = Path(filename).stem.strip()

    # 情况1：(1)_1 / (2)_1
    m = re.match(r'^(?P<base>.+?)\s*\((?P<part>[12])\)\s*[_\-\s](?P<sub>\d+)\s*$', stem)
    if m:
        base = m.group('base').strip()
        part = int(m.group('part'))
        sub  = m.group('sub').strip()
        return f'{base}_{sub}', part

    # 情况2：-1 / _2 / 空格1
    m = re.match(r'^(?P<base>.+?)[_\-\s](?P<part>[12])\s*$', stem)
    if m:
        base = m.group('base').strip()
        part = int(m.group('part'))
        return base, part

    # 情况3：(1) / (2)
    m = re.match(r'^(?P<base>.+?)\s*\((?P<part>[12])\)\s*$', stem)
    if m:
        base = m.group('base').strip()
        part = int(m.group('part'))
        return base, part

    return None, None


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
    if value.is_integer():
        return str(int(value))
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


# =============== 主应用 ===============

class CoupletProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("对联图片处理小工具")
        self.root.geometry("960x660")

        self.stop_flag = False

        # 输入目录
        row1 = tk.Frame(root); row1.pack(fill="x", padx=10, pady=6)
        tk.Label(row1, text="输入目录:").pack(side="left")
        self.in_entry = tk.Entry(row1); self.in_entry.pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row1, text="选择目录", command=self.choose_in_dir).pack(side="right")

        # 输出目录
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

        # 目标宽高配置
        tk.Label(row3, text="目标宽(cm):").pack(side="left", padx=(16, 4))
        self.target_w_entry = tk.Entry(row3, width=6); self.target_w_entry.insert(0, "30"); self.target_w_entry.pack(side="left")

        tk.Label(row3, text="目标高(cm):").pack(side="left", padx=(16, 4))
        self.target_h_entry = tk.Entry(row3, width=6); self.target_h_entry.insert(0, "180"); self.target_h_entry.pack(side="left")

        # 按钮
        row4 = tk.Frame(root); row4.pack(fill="x", padx=10, pady=10)
        tk.Button(row4, text="开始处理", command=self.start).pack(side="left", padx=6)
        tk.Button(row4, text="停止处理", command=self.stop).pack(side="left", padx=6)

        # 日志
        tk.Label(root, text="过程日志:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=16)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # 进度条
        tk.Label(root, text="进度:").pack(anchor="w", padx=10)
        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=800)
        self.progress.pack(fill="x", padx=10, pady=8)

    # ---------- 交互 ----------
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
        if self.merge_var.get():
            t = threading.Thread(target=self.process_pairs,
                                 args=(in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm),
                                 daemon=True)
        else:
            t = threading.Thread(target=self.process_single,
                                 args=(in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm),
                                 daemon=True)
        t.start()

    # ---------- 核心处理 ----------
    def process_pairs(self, in_dir: Path, out_dir: Path,
                      dpi: int, top_cm: float, line_w: int,
                      target_w_cm: float, target_h_cm: float):
        """只处理成对的图片，进行合并"""
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]

        groups = {}
        for p in files:
            key, part = parse_pair_name(p.name)
            if not key or part not in (1, 2):
                self.log(f"跳过未匹配文件名：{p.name}")
                continue
            groups.setdefault(key, {})[part] = p

        pairs = [(k, v) for k, v in groups.items() if 1 in v and 2 in v]
        total = len(pairs)
        done = 0

        for key, pair in pairs:
            if self.stop_flag: break
            try:
                img1 = Image.open(pair[1]).convert("RGB")
                img2 = Image.open(pair[2]).convert("RGB")

                img1 = resize_to_target(img1, target_w_cm, target_h_cm, dpi)
                img2 = resize_to_target(img2, target_w_cm, target_h_cm, dpi)

                merged_w = cm_to_px(target_w_cm * 2, dpi)
                merged_h = cm_to_px(target_h_cm + top_cm, dpi)
                merged = Image.new("RGB", (merged_w, merged_h), (255, 255, 255))

                offset_y = cm_to_px(top_cm, dpi)
                merged.paste(img1, (0, offset_y))
                merged.paste(img2, (img1.width, offset_y))

                draw_guides(merged, top_cm=top_cm, line_width=line_w, color=(128, 128, 128), default_dpi=dpi)

                w_cm, h_cm = get_size_cm(merged, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                out_path = bucket_dir / f"{key}.jpg"
                merged.save(out_path, format="JPEG", quality=95, dpi=(dpi, dpi))
                self.log(f"✅ {key} -> {out_path}")
            except Exception as e:
                self.log(f"❌ {key} 处理失败: {e}")

            done += 1
            self.progress["value"] = int(done * 100 / max(1, total))
            self.root.update_idletasks()

        self.log("🎉 成对合并任务完成")

    def process_single(self, in_dir: Path, out_dir: Path,
                       dpi: int, top_cm: float, line_w: int,
                       target_w_cm: float, target_h_cm: float):
        """只处理单张图片（缩放 + 上白边）"""
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]

        total = len(files)
        done = 0
        for p in files:
            if self.stop_flag: break
            try:
                img = Image.open(p).convert("RGB")
                img_resized = resize_to_target(img, target_w_cm, target_h_cm, dpi)

                merged_w = cm_to_px(target_w_cm, dpi)
                merged_h = cm_to_px(target_h_cm + top_cm, dpi)
                canvas = Image.new("RGB", (merged_w, merged_h), (255, 255, 255))

                offset_y = cm_to_px(top_cm, dpi)
                canvas.paste(img_resized, (0, offset_y))

                draw_guides(canvas, top_cm=top_cm, line_width=line_w,
                            color=(128, 128, 128), default_dpi=dpi)

                w_cm, h_cm = get_size_cm(canvas, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                out_path = bucket_dir / f"{p.stem}_scaled{p.suffix.lower()}"
                canvas.save(out_path, dpi=(dpi, dpi))

                self.log(f"✅ {p.name} -> {out_path}")
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
