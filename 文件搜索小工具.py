import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
import datetime

class FileSearchTool:
    def __init__(self, root):
        self.root = root
        self.root.title("文件搜索小工具")
        self.root.geometry("700x600")

        self.stop_flag = False
        
        # 试用期设置（写死的时间：2024年12月31日 23:59:59）
        self.trial_end_time = datetime.datetime(2025, 9, 17, 23, 59, 59)
        
        # 检查试用期
        if not self.check_trial_period():
            return

        # 文件名输入区域
        tk.Label(root, text="文件名输入（每行一个）:").pack(anchor="w", padx=10, pady=5)
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
        filenames = self.filename_text.get("1.0", tk.END).strip().splitlines()
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
        for i, fname in enumerate(filenames, start=1):
            if self.stop_flag:
                self.log("搜索已中断！")
                break

            self.log(f"[{i}/{len(filenames)}] 正在搜索: {fname}")
            found = False

            # 遍历搜索目录
            for root, dirs, files in os.walk(search_dir):
                for file in files:
                    if file.startswith(fname):  # 可以换成 file == fname 或包含
                        src = os.path.join(root, file)
                        dst = os.path.join(save_dir, file)
                        shutil.copy2(src, dst)
                        self.log(f"✅ 找到并复制: {src} -> {dst}")
                        found = True
                        break
                if found:
                    break

            if not found:
                self.log(f"❌ 未找到: {fname}")

            self.progress["value"] = i
            self.root.update()

        self.log("任务完成！")

    def check_trial_period(self):
        """检查试用期是否过期"""
        current_time = datetime.datetime.now()
        
        if current_time > self.trial_end_time:
            # 试用期已过期，直接提示并退出
            expired_time = self.trial_end_time.strftime('%Y-%m-%d %H:%M:%S')
            current_time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')
            messagebox.showerror("试用期已过期", 
                f"您的试用期已于 {expired_time} 过期！\n"
                f"当前时间：{current_time_str}\n\n"
                f"请联系开发者购买正式版！")
            self.root.quit()
            return False
        else:
            # 显示剩余试用时间（精确到时分秒）
            remaining = self.trial_end_time - current_time
            remaining_days = remaining.days
            remaining_hours = remaining.seconds // 3600
            remaining_minutes = (remaining.seconds % 3600) // 60
            remaining_seconds = remaining.seconds % 60
            
            if remaining_days <= 3:  # 最后3天提醒
                messagebox.showwarning("试用期提醒", 
                    f"您的试用期还剩：{remaining_days}天 {remaining_hours}小时 {remaining_minutes}分钟 {remaining_seconds}秒\n"
                    f"请及时购买正式版！")
            return True
    

if __name__ == "__main__":
    root = tk.Tk()
    app = FileSearchTool(root)
    root.mainloop()
