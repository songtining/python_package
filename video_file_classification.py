import os
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict

class FileClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件分类工具")
        self.root.geometry("500x400")

        # 文件夹选择部分
        self.folder_path = tk.StringVar()
        tk.Label(root, text="选择文件夹:").pack(pady=5)

        self.folder_frame = tk.Frame(root)
        self.folder_frame.pack(pady=5)

        self.entry_folder = tk.Entry(self.folder_frame, textvariable=self.folder_path, width=40)
        self.entry_folder.pack(side=tk.LEFT, padx=5)

        self.btn_browse = tk.Button(self.folder_frame, text="浏览", command=self.browse_folder)
        self.btn_browse.pack(side=tk.LEFT)

        # 最小文件数量输入
        tk.Label(root, text="最小分类文件数量:").pack(pady=5)

        self.min_files_var = tk.StringVar()
        self.min_files_var.set("5")  # 默认值
        self.spin_min_files = tk.Spinbox(root, from_=1, to=100, textvariable=self.min_files_var, width=5)
        self.spin_min_files.pack()

        # 处理分类按钮
        self.btn_process = tk.Button(root, text="处理分类", command=self.process_files, state=tk.DISABLED)
        self.btn_process.pack(pady=10)

        # 替换标题：选择txt文本路径
        tk.Label(root, text="选择替换标题txt文件:").pack(pady=5)

        self.txt_path = tk.StringVar()
        self.txt_frame = tk.Frame(root)
        self.txt_frame.pack(pady=5)

        self.entry_txt = tk.Entry(self.txt_frame, textvariable=self.txt_path, width=40)
        self.entry_txt.pack(side=tk.LEFT, padx=5)

        self.btn_choose_txt = tk.Button(self.txt_frame, text="浏览", command=self.choose_txt)
        self.btn_choose_txt.pack(side=tk.LEFT)

        # 替换按钮
        self.btn_rename = tk.Button(root, text="替换标题", command=self.replace_titles, state=tk.DISABLED)
        self.btn_rename.pack(pady=10)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(root, textvariable=self.status_var, fg="blue")
        self.status_label.pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.btn_process.config(state=tk.NORMAL)
            self.status_var.set("已选择文件夹，可以开始处理")

    def choose_txt(self):
        txt_file = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if txt_file:
            self.txt_path.set(txt_file)
            self.btn_rename.config(state=tk.NORMAL)
            self.status_var.set("已选择标题替换文本")

    def process_files(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("错误", "请先选择文件夹")
            return

        try:
            min_files = int(self.min_files_var.get())
            if min_files <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "请输入有效的正整数")
            return

        files = [
            f for f in os.listdir(folder)
            if self.is_valid_filename(os.path.join(folder, f))
        ]

        groups = defaultdict(list)
        for filename in files:
            prefix = filename.split('_')[0]
            groups[prefix].append(filename)

        valid_groups = {k: v for k, v in groups.items() if len(v) >= min_files}

        if not valid_groups:
            messagebox.showinfo("提示", "没有找到符合条件的文件分组")
            return

        self.status_var.set("正在处理...")
        self.root.update()

        created_count = 0
        for prefix, filenames in valid_groups.items():
            target_dir = os.path.join(folder, prefix)
            os.makedirs(target_dir, exist_ok=True)

            for filename in filenames:
                src = os.path.join(folder, filename)
                dst = os.path.join(target_dir, filename)
                os.rename(src, dst)

            created_count += 1

        self.status_var.set(f"处理完成！创建了 {created_count} 个分类文件夹")
        messagebox.showinfo("完成", f"成功创建 {created_count} 个分类文件夹")

    def replace_titles(self):
        folder = self.folder_path.get()
        txt_path = self.txt_path.get()

        if not folder or not txt_path:
            messagebox.showerror("错误", "请先选择文件夹和替换标题文件")
            return

        # 读取并按 ID 分组
        id_groups = defaultdict(list)
        with open(txt_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if '_' not in line:
                    continue
                parts = line.split('_', 1)
                if len(parts) == 2:
                    id_groups[parts[0]].append(parts[0] + '_' + parts[1])

        modified_count = 0

        for id_prefix, new_titles in id_groups.items():
            dir_path = os.path.join(folder, id_prefix)
            if not os.path.isdir(dir_path):
                continue

            files = sorted(
                [f for f in os.listdir(dir_path) if os.path.isfile(os.path.join(dir_path, f))]
            )

            if len(files) != len(new_titles):
                print(f"⚠️ {id_prefix} 文件数量不匹配，跳过")
                continue

            for i, old_name in enumerate(files):
                old_path = os.path.join(dir_path, old_name)
                ext = os.path.splitext(old_name)[1]
                new_name = f"{new_titles[i]}{ext}"
                new_path = os.path.join(dir_path, new_name)
                os.rename(old_path, new_path)

            modified_count += 1

        self.status_var.set(f"已完成重命名，成功处理 {modified_count} 个分组")
        messagebox.showinfo("完成", f"成功替换 {modified_count} 个文件夹的文件名")

    @staticmethod
    def is_valid_filename(file_path):
        if not os.path.isfile(file_path):
            return False
        filename = os.path.basename(file_path)
        parts = filename.split('_')
        return len(parts) == 2 and all(parts)


import datetime

if __name__ == "__main__":
    # 设置试用期截止时间（年, 月, 日, 时, 分, 秒）
    expiration_datetime = datetime.datetime(2025, 5, 1, 12, 0, 0)  # 2025-05-01 12:00:00

    # 获取当前时间
    now = datetime.datetime.now()

    if now > expiration_datetime:
        root = tk.Tk()
        root.withdraw()  # 不显示主界面
        messagebox.showerror("试用已过期", "该工具的试用期已结束（到期时间：2025-05-01 12:00:00）。\n请联系开发者获取正式授权。")
        root.destroy()
    else:
        root = tk.Tk()
        app = FileClassifierApp(root)
        root.mainloop()
