#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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
import os
import time
import traceback

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

def parse_pair_name(filename: str):
    stem = Path(filename).stem.strip()

    # 1. 处理括号格式: xxx(1)_2
    m = re.match(r'^(?P<key>.+?)\s*\((?P<part>[12])\)_?(?P<sub>\d+)?$', stem)
    if m:
        return m.group("key"), int(m.group("part")), int(m.group("sub") or 0)

    # 2. 处理 -1_2 这种: xxx-1_2
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])_?(?P<sub>\d+)?$', stem)
    if m:
        return m.group("key"), int(m.group("part")), int(m.group("sub") or 0)

    # 3. 处理括号无子图: xxx(1)
    m = re.match(r'^(?P<key>.+?)\s*\((?P<part>[12])\)$', stem)
    if m:
        return m.group("key"), int(m.group("part")), 0

    # 4. 处理 -1 这种: xxx-1
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])$', stem)
    if m:
        return m.group("key"), int(m.group("part")), 0

    # 5. 处理 xxx_1
    m = re.match(r'^(?P<key>.+?)_(?P<sub>\d+)$', stem)
    if m:
        return m.group("key"), None, int(m.group("sub"))

    # 6. 特殊处理：任意前缀 + -<part>--<任意数字>--<sub>
    m = re.match(r'^(?P<key>.+)-(?P<part>[12])--\d+--(?P<sub>\d+)$', stem)
    if m:
        return m.group("key"), int(m.group("part")), int(m.group("sub"))

    # 7. 特殊处理: key-<part>--<sub>
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])--(?P<sub>\d+)$', stem)
    if m:
        return m.group("key"), int(m.group("part")), int(m.group("sub"))

    # 8. 特殊处理: key-<part>--1--<sub>
    m = re.match(r'^(?P<key>.+?)-(?P<part>[12])--1--(?P<sub>\d+)$', stem)
    if m:
        return m.group("key"), int(m.group("part")), int(m.group("sub"))

    # 默认
    return stem, None, 0

def parse_triplet_pattern(filename: str):
    """解析末尾三段型 ...-a--b--c，返回 (base, a, b, c)，否则返回 (base, None, None, 0)
    适配两种规律：
    - 规律A：...-1--x--x / ...-2--x--x  → 按 a(1/2) 配对，c 常为 1
    - 规律B：...-x--x--1 / ...-x--x--2  → 按 c(1/2) 配对，a 常为 1
    """
    stem = Path(filename).stem.strip()
    if "--" in stem:
        parts = stem.split("--")
        if len(parts) >= 3:
            c_str = parts[-1]
            b_str = parts[-2]
            left = "--".join(parts[:-2])
            m_a = re.search(r"-(?P<a>[0-9]+)$", left)
            if m_a and c_str.isdigit():
                a_val = int(m_a.group("a"))
                b_val = int(b_str) if b_str.isdigit() else None
                c_val = int(c_str)
                base = left[: -(len(m_a.group(0)))]  # 去掉 -<a>
                return base, a_val, b_val, c_val
    return stem, None, None, 0

def get_image_dpi(img: Image.Image, default_dpi=300):
    dpi = img.info.get('dpi')
    if isinstance(dpi, tuple) and len(dpi) >= 2:
        return dpi[0] or default_dpi, dpi[1] or default_dpi
    return default_dpi, default_dpi

def cm_to_px(cm: float, dpi: float) -> int:
    return int(round(cm * dpi / 2.54))

def draw_guides(img: Image.Image, top_cm=2.5, line_width=3,
                color=(255, 255, 255), default_dpi=300, log_func=print):
    draw = ImageDraw.Draw(img)
    dpi_x, dpi_y = get_image_dpi(img, default_dpi)
    
    # 绘制横线
    y = cm_to_px(top_cm, dpi_y)
    y = max(0, min(img.height - 1, y))
    log_func(f"📏 横线坐标: y={y}px (top_cm={top_cm}cm, dpi_y={dpi_y})")
    draw.line([(0, y), (img.width, y)], fill=color, width=line_width)

    # 绘制竖线（上下各十分之二）
    x = img.width // 2
    
    # 计算上下各十分之二的高度
    segment_height = img.height * 2 // 10
    segment_height = max(1, segment_height)  # 确保至少为1像素
    
    log_func(f"📏 图片尺寸: {img.width}x{img.height}px, DPI: ({dpi_x}, {dpi_y})")
    log_func(f"📏 竖线坐标: x={x}px (图片宽度中心)")
    log_func(f"📏 竖线段高度: {segment_height}px (图片高度的2/10)")
    
    # 上部 2/10：从顶部开始画
    top_end = segment_height
    log_func(f"📏 上部竖线: 从 (x={x}, y=0) 到 (x={x}, y={top_end})")
    draw.line([(x, 0), (x, top_end)], fill=color, width=line_width)

    # 下部 2/10：从底部向上画
    bottom_start = img.height - segment_height
    bottom_end = img.height
    log_func(f"📏 下部竖线: 从 (x={x}, y={bottom_start}) 到 (x={x}, y={bottom_end})")
    draw.line([(x, bottom_start), (x, bottom_end)], fill=color, width=line_width)
    
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

# =============== 导出辅助 ===============
def save_as_tif(image: Image.Image, tif_path, dpi: int):
    """以无损 LZW 压缩方式保存为 TIF，并写入 DPI"""
    Path(tif_path).parent.mkdir(parents=True, exist_ok=True)
    image.save(tif_path, format="TIFF", compression="tiff_lzw", dpi=(dpi, dpi))

# =============== Photoshop 转换函数 ===============
def convert_rgb_to_cmyk_jpeg(input_jpg, output_jpg, ps_app=None, log_func=print):
    """在同一线程内复用传入的 ps_app；不跨线程复用"""
    try:
        # 绝对路径 & 输出目录
        input_path = str(Path(input_jpg).resolve())
        output_path = str(Path(output_jpg).resolve())
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        if ps_app is None:
            log_func("🚀 启动 Photoshop...")
            ps_app = get_photoshop_app(log_func)

        log_func(f"➡ 打开文件: {input_path}")
        doc = ps_app.Open(input_path)
        if doc is None:
            log_func(f"❌ 无法打开文件: {input_path}")
            return False

        # 3 = psCMYKMode（避免魔法数可用 constants.psCMYK）
        if getattr(doc, "Mode", None) != 3:
            log_func("🎨 转换为 CMYK 模式")
            doc.ChangeMode(3)
            # 有些版本切换模式后必须先保存一次
            try:
                doc.Save()
            except Exception:
                pass

        log_func(f"💾 保存为 JPEG: {output_path}")
        options = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        options.Quality = 12
        try:
            options.Matte = 1  # 1 = psNoMatte（部分版本可无此属性）
        except Exception:
            pass
        try:
            options.EmbedColorProfile = True
        except Exception:
            pass

        # SaveAs(SaveIn, Options, AsCopy)
        doc.SaveAs(output_path, options, True)
        # 2 = psDoNotSaveChanges
        try:
            doc.Close(2)
        except Exception:
            doc.Close()

        log_func(f"✅ CMYK 转换完成: {output_path}")
        return True

    except Exception as e:
        log_func(f"❌ CMYK 转换失败: {e}\n{traceback.format_exc()}")
        return False

def get_photoshop_app(log_func=print):
    """健壮获取 Photoshop COM 对象（每线程独立实例）"""
    if not sys.platform.startswith('win'):
        raise RuntimeError("当前系统不是 Windows，无法使用 Photoshop COM 接口")

    progids = [
        "Photoshop.Application",
        "Photoshop.Application.2025",
        "Photoshop.Application.2024",
        "Photoshop.Application.2023",
        "Photoshop.Application.2022",
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
            try:
                app.DisplayDialogs = 3  # psDisplayNoDialogs
            except Exception:
                pass
            # 给 Photoshop 一点时间完全就绪
            time.sleep(5)
            return app
        except Exception as e:
            last_err = e
    raise RuntimeError(f"无法启动 Photoshop COM，请确认已安装并可正常启动。原始错误: {last_err}")

# =============== 主应用 ===============

class CoupletProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("自动调图软件（图片合并 & CMYK模式转换）V26.01.18")
        self.root.geometry("1100x750")

        self.stop_flag = False

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
    @com_thread
    def process_pairs(self, in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm):
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]
        groups = {}
        # 预扫描：以最后一段 1/2 的前缀做直配候选
        last_pair_candidates = {}
        stems = [p.stem for p in files]
        for stem in stems:
            if "--" in stem:
                base_last, last = stem.rsplit("--", 1)
                if last.isdigit() and int(last) in (1, 2):
                    d = last_pair_candidates.setdefault(base_last, set())
                    d.add(int(last))

        for f in files:
            base, a, b, c = parse_triplet_pattern(f.name)

            # 第一优先：若存在“仅最后一位不同 1/2”的前缀，按该前缀直配
            stem = f.stem
            if "--" in stem:
                base_last, last = stem.rsplit("--", 1)
                if last.isdigit() and int(last) in (1, 2):
                    present = last_pair_candidates.get(base_last, set())
                    if present == {1, 2}:
                        groups.setdefault((base_last, 'last_c'), {})[int(last)] = f
                        continue

            # 第二优先：先规律A，再规律B
            if a in (1, 2) and c == 1:
                groups.setdefault((base, 'vary_a', b, c), {})[a] = f
                continue
            if c in (1, 2) and a == 1:
                groups.setdefault((base, 'vary_c', a, b), {})[c] = f
                continue
            # 兜底：退回旧解析
            key, part, sub = parse_pair_name(f.name)
            groups.setdefault((key, sub), {})
            if part is None: part = 1 if 2 in groups[(key, sub)] else 2
            groups[(key, sub)][part] = f

        # 打印分组结果
        self.log("📂 文件分组结果：")
        unpaired = []

        def sort_key(item):
            k = item[0]
            if isinstance(k, tuple) and len(k) >= 2 and k[1] == 'last_c':
                return (str(k[0]), 'last_c')
            if isinstance(k, tuple) and len(k) >= 2 and k[1] in ('vary_a','vary_c'):
                return (str(k[0]), str(k[1]), str(k[2]), str(k[3]) if len(k) > 3 else '')
            if isinstance(k, tuple) and len(k) == 2:
                return (str(k[0]), str(k[1]))
            return (str(k), '')
        sorted_items = sorted(groups.items(), key=sort_key)

        group_index = 1
        for gk, parts in sorted_items:
            has1 = 1 in parts
            has2 = 2 in parts
            if has1 and has2:
                rule = 'last_c' if (isinstance(gk, tuple) and len(gk)>=2 and gk[1]=='last_c') else ('vary_a' if (isinstance(gk, tuple) and len(gk)>=2 and gk[1]=='vary_a') else ('vary_c' if (isinstance(gk, tuple) and len(gk)>=2 and gk[1]=='vary_c') else 'fallback'))
                self.log(f"第{group_index}组：[{rule}]")
                self.log(parts[1].name)
                self.log(parts[2].name)
                group_index += 1
            else:
                files_info = ", ".join([f"part{p}:{f.name}" for p, f in parts.items()])
                unpaired.append((gk, files_info))

        if unpaired:
            self.log("⚠️ 未成对文件列表：")
            for gk, files_info in unpaired:
                self.log(f"   - {gk}: {files_info}")

        pairs = [(k, v) for k, v in groups.items() if 1 in v and 2 in v]
        total = len(pairs); done = 0
        saved_jpgs = []  # 第一阶段保存完成的 RGB JPG 列表

        for gk, pair in pairs:
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
                draw_guides(merged, top_cm=top_cm, line_width=line_w, color=(128, 128, 128), default_dpi=dpi, log_func=self.log)

                w_cm, h_cm = get_size_cm(merged, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                first_file = pair[1] if 1 in pair else pair[2]
                base_stem = Path(first_file.name).stem
                out_name = f"{base_stem}.jpg"
                out_path = bucket_dir / out_name
                merged.save(out_path, format="JPEG", quality=95, dpi=(dpi, dpi))
                self.log(f"✅ 已输出: {out_path}")
                saved_jpgs.append((out_path, w_cm, h_cm))

            except Exception as e:
                self.log(f"❌ 处理失败: group={gk}, 错误: {e}")
            done += 1
            self.progress["value"] = int(done * 100 / max(1, total))
            self.root.update_idletasks()
        self.log("🎉 成对合并任务完成")

        # 第二阶段：批量转换为 CMYK（如勾选）
        if self.cmyk_var.get() and saved_jpgs:
            try:
                self.log("🚀 启动 Photoshop...")
                psApp = get_photoshop_app(self.log)  # ✅ 每线程内局部实例
            except Exception as e:
                self.log(f"❌ 启动 Photoshop 失败: {e}")
                return

            self.log("▶ 开始第二阶段：批量转换 CMYK...")
            total2 = len(saved_jpgs); done2 = 0
            for out_path, w_cm, h_cm in saved_jpgs:
                if self.stop_flag: break
                try:
                    # 生成中间 TIF
                    tif_path = out_path.with_suffix(".tif")
                    self.log(f"📝 生成中间 TIF: {tif_path}")
                    with Image.open(out_path) as _img:
                        save_as_tif(_img, tif_path, dpi)

                    bucket_dir_cmyk = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm_cmyk")
                    cmyk_path = bucket_dir_cmyk / out_path.name

                    convert_rgb_to_cmyk_jpeg(tif_path, cmyk_path, psApp, self.log)

                    # 清理中间文件
                    try:
                        Path(tif_path).unlink(missing_ok=True)
                        self.log(f"🧹 已删除中间 TIF: {tif_path}")
                    except Exception as e:
                        self.log(f"⚠️ 删除中间 TIF 失败: {e}")
                except Exception as e:
                    self.log(f"❌ CMYK 转换失败: {e}")
                done2 += 1
                self.progress["value"] = int(done2 * 100 / max(1, total2))
                self.root.update_idletasks()
            self.log("🎉 第二阶段 CMYK 转换完成")

    # ---------- 单图处理 ----------
    @com_thread
    def process_single(self, in_dir, out_dir, dpi, top_cm, line_w, target_w_cm, target_h_cm):
        files = [p for p in in_dir.iterdir() if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png")]
        total = len(files); done = 0
        saved_jpgs = []

        for p in files:
            if self.stop_flag: break
            try:
                img = resize_to_target(Image.open(p).convert("RGB"), target_w_cm, target_h_cm, dpi)
                merged_w = cm_to_px(target_w_cm, dpi)
                merged_h = cm_to_px(target_h_cm + top_cm, dpi)
                canvas = Image.new("RGB", (merged_w, merged_h), (255, 255, 255))
                offset_y = cm_to_px(top_cm, dpi)
                canvas.paste(img, (0, offset_y))
                draw_guides(canvas, top_cm=top_cm, line_width=line_w, color=(128, 128, 128), default_dpi=dpi, log_func=self.log)

                w_cm, h_cm = get_size_cm(canvas, dpi, top_margin_cm=top_cm)
                bucket_dir = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm")
                out_path = bucket_dir / f"{p.stem}.jpg"
                canvas.save(out_path, dpi=(dpi, dpi))
                self.log(f"✅ 已输出: {out_path}")
                saved_jpgs.append((out_path, w_cm, h_cm))

            except Exception as e:
                self.log(f"❌ {p.name} 处理失败: {e}")
            done += 1
            self.progress["value"] = int(done * 100 / max(1, total))
            self.root.update_idletasks()
        self.log("🎉 单图缩放任务完成")

        # 第二阶段：批量转换为 CMYK（如勾选）
        if self.cmyk_var.get() and saved_jpgs:
            try:
                self.log("🚀 启动 Photoshop...")
                psApp = get_photoshop_app(self.log)  # ✅ 每线程内局部实例
            except Exception as e:
                self.log(f"❌ 启动 Photoshop 失败: {e}")
                return

            self.log("▶ 开始第二阶段：批量转换 CMYK...")
            total2 = len(saved_jpgs); done2 = 0
            for out_path, w_cm, h_cm in saved_jpgs:
                if self.stop_flag: break
                try:
                    tif_path = out_path.with_suffix(".tif")
                    self.log(f"📝 生成中间 TIF: {tif_path}")
                    with Image.open(out_path) as _img:
                        save_as_tif(_img, tif_path, dpi)

                    bucket_dir_cmyk = ensure_folder(out_dir / f"{w_cm}x{h_cm}cm_cmyk")
                    cmyk_path = bucket_dir_cmyk / out_path.name

                    convert_rgb_to_cmyk_jpeg(tif_path, cmyk_path, psApp, self.log)

                    try:
                        Path(tif_path).unlink(missing_ok=True)
                        self.log(f"🧹 已删除中间 TIF: {tif_path}")
                    except Exception as e:
                        self.log(f"⚠️ 删除中间 TIF 失败: {e}")
                except Exception as e:
                    self.log(f"❌ CMYK 转换失败: {e}")
                done2 += 1
                self.progress["value"] = int(done2 * 100 / max(1, total2))
                self.root.update_idletasks()
            self.log("🎉 第二阶段 CMYK 转换完成")

# =============== 入口 ===============
if __name__ == "__main__":
    root = tk.Tk()
    app = CoupletProcessorApp(root)
    root.mainloop()
