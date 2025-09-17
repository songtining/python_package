import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import datetime

class FileSearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("文件搜索小工具V1.1")
        self.root.geometry("700x600")

        self.stop_flag = False

        # 文件名输入区域
        tk.Label(root, text="请输入文件名 & 数量（每行一个）:").pack(anchor="w", padx=10, pady=5)
        self.filename_text = scrolledtext.ScrolledText(root, height=8)
        self.filename_text.pack(fill="x", padx=10)

        # 搜索目录
        frame1 = tk.Frame(root)
        frame1.pack(fill="x", padx=10, pady=5)
        tk.Label(frame1, text="搜索目录:").pack(side="left")
        self.search_entry = tk.Entry(frame1)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame1, text="选择目录", command=self.choose_search_dir).pack(side="right")

        # 保存目录
        frame2 = tk.Frame(root)
        frame2.pack(fill="x", padx=10, pady=5)
        tk.Label(frame2, text="保存目录:").pack(side="left")
        self.save_entry = tk.Entry(frame2)
        self.save_entry.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame2, text="选择目录", command=self.choose_save_dir).pack(side="right")

        # 按钮
        frame3 = tk.Frame(root)
        frame3.pack(fill="x", padx=10, pady=10)
        tk.Button(frame3, text="开始搜索", command=self.start_search).pack(side="left", padx=5)
        tk.Button(frame3, text="停止搜索", command=self.stop_search).pack(side="left", padx=5)

        # 日志输出
        tk.Label(root, text="日志输出:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=15)
        self.log_text.pack(fill="both", padx=10, pady=5, expand=True)

        # 进度条
        tk.Label(root, text="进度:").pack(anchor="w", padx=10)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(padx=10, pady=5)

    def choose_search_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.search_entry.delete(0, tk.END)
            self.search_entry.insert(0, path)

    def choose_save_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, path)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def stop_search(self):
        self.stop_flag = True
        self.log("停止任务...")


    def start_search(self):
        # 解析输入，每行格式：文件名 数量
        raw_lines = self.filename_text.get("1.0", tk.END).strip().splitlines()
        filenames = []
        for line in raw_lines:
            parts = line.strip().split()
            if not parts:
                continue
            fname = parts[0]
            count = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 1
            filenames.append((fname, count))

        search_dir = self.search_entry.get().strip()
        save_dir = self.save_entry.get().strip()

        if not filenames or not search_dir or not save_dir:
            messagebox.showwarning("警告", "请输入文件名并选择目录！")
            return

        self.stop_flag = False
        self.progress["value"] = 0
        self.progress["maximum"] = len(filenames)

        t = threading.Thread(target=self.search_files, args=(filenames, search_dir, save_dir))
        t.start()

    def search_files(self, filenames, search_dir, save_dir):
        global_index = 1  # 全局序号前缀（每一行关键字一个序号）

        for i, (fname, count) in enumerate(filenames, start=1):
            if self.stop_flag:
                self.log("搜索已中断！")
                break

            # 收集所有匹配到的文件（前缀匹配）
            matches = []
            for root, dirs, files in os.walk(search_dir):
                for file in files:
                    if file.startswith(fname):  # 依然使用前缀匹配
                        matches.append(os.path.join(root, file))

            if not matches:
                self.log(f"[{i}/{len(filenames)}] ❌ 未找到: {fname}")
                self.progress["value"] = i
                self.root.update()
                continue

            # 日志：找到多少个候选文件
            self.log(f"[{i}/{len(filenames)}] ✅ 找到 {len(matches)} 个匹配: {fname} (每个复制 {count} 份)")

            # 对同一关键字下的所有匹配文件，逐个复制；该关键字的全局前缀保持一致
            for src in sorted(matches):
                name, ext = os.path.splitext(os.path.basename(src))
                for c in range(1, count + 1):
                    new_name = f"{global_index}_{name}_{c}{ext}"
                    dst = os.path.join(save_dir, new_name)
                    shutil.copy2(src, dst)
                    self.log(f"   → 复制: {src} -> {dst}")

            # 这一行（一个关键字）处理完后，再递增全局前缀
            global_index += 1

            self.progress["value"] = i
            self.root.update()

        self.log("任务完成！")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileSearchTool(root)
    root.mainloop()
