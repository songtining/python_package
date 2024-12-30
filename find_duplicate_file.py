import os
import sys
import hashlib
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


def check_usage_expiry():
    """
    检查程序是否在使用期限内。
    """
    expiry_date = datetime(2024, 12, 30, 16, 59)
    current_date = datetime.now()

    if current_date > expiry_date:
        messagebox.showerror("错误", "程序的使用期限已过，无法运行。请联系开发者以更新版本。")
        sys.exit(1)


def calculate_md5(file_path):
    """
    计算文件的 MD5 哈希值。
    """
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
    except Exception as e:
        print(f"[错误] 无法读取文件: {file_path}，错误: {e}")
        return None
    print(f"[信息] 计算 MD5 成功: {file_path}")
    return hash_md5.hexdigest()


def find_version_conflicts(root_dir, output_file):
    """
    找出文件名相同但 MD5 不同的文件，并导出到 Excel 文件。
    """
    file_info = []  # 保存文件信息

    print(f"[信息] 正在扫描文件夹: {root_dir}")

    # 遍历文件夹，收集所有文件的信息
    total_files = 0
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            total_files += 1
            file_path = os.path.join(dirpath, filename)
            print(f"[扫描] 正在处理文件: {file_path}")
            md5_hash = calculate_md5(file_path)
            if md5_hash:
                file_info.append({"filename": filename, "file_path": file_path, "md5": md5_hash})

    print(f"[信息] 扫描完成，共处理文件数: {total_files}")

    # 转换为 DataFrame 方便处理
    df = pd.DataFrame(file_info)

    # 按文件名分组，查找 MD5 不同的文件
    print(f"[信息] 正在分析文件名重复...")
    conflicts = []
    grouped = df.groupby("filename")
    for filename, group in grouped:
        if len(group["md5"].unique()) > 1:  # 如果同名文件有不同的 MD5 值
            print(f"[重复] 文件名: {filename}, 重复文件数: {len(group)}")
            conflicts.append(
                {"filename": filename, "file_paths": group["file_path"].tolist(), "md5s": group["md5"].tolist()})

    # 保存重复文件名到 Excel
    if conflicts:
        print(f"[信息] 检测到 {len(conflicts)} 个文件重复。正在保存到 Excel...")
        output_data = []
        for conflict in conflicts:
            for path, md5 in zip(conflict["file_paths"], conflict["md5s"]):
                output_data.append({"文件名": conflict["filename"], "文件路径": path, "MD5": md5})

        output_df = pd.DataFrame(output_data)
        output_df.to_excel(output_file, index=False)
        print(f"[成功] 重复文件已导出到: {output_file}")
        messagebox.showinfo("完成", f"重复文件已导出到: {output_file}")
    else:
        print("[信息] 未发现文件名相同但 MD5 不同的文件。")
        messagebox.showinfo("完成", "未发现文件名相同但 MD5 不同的文件。")


def start_scan():
    """
    启动扫描操作。
    """
    root_dir = folder_path.get()
    if not root_dir:
        messagebox.showerror("错误", "请选择要扫描的文件夹路径。")
        return

    if not os.path.exists(root_dir):
        messagebox.showerror("错误", "输入的文件夹路径不存在或无法访问，请检查后重试。")
        return

    output_file = "重复文件查找结果.xlsx"
    find_version_conflicts(root_dir, output_file)


def browse_folder():
    """
    选择文件夹。
    """
    path = filedialog.askdirectory()
    folder_path.set(path)


if __name__ == "__main__":
    # 检查使用期限
    check_usage_expiry()

    # 创建主窗口
    app = tk.Tk()
    app.title("重复文件查找工具")
    app.geometry("500x200")

    # 文件夹选择
    folder_path = tk.StringVar()
    tk.Label(app, text="请选择要扫描的文件夹路径:").pack(pady=10)
    tk.Entry(app, textvariable=folder_path, width=50).pack(pady=5)
    tk.Button(app, text="浏览", command=browse_folder).pack(pady=5)

    # 开始按钮
    tk.Button(app, text="开始扫描", command=start_scan, bg="green", fg="white").pack(pady=20)

    # 启动应用
    app.mainloop()
