import os
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from datetime import datetime


def check_trial_period():
    """检查试用时间是否过期"""
    trial_end = datetime(2025, 1, 9, 18, 0, 0)  # 设置试用截止时间
    current_time = datetime.now()
    if current_time > trial_end:
        messagebox.showerror("试用已结束", "该程序的试用期已结束，感谢您的使用！")
        logging.warning("试用期已过，程序退出。")
        exit()


def setup_logging():
    # 创建日志文件
    log_file = os.path.join(os.getcwd(), f"processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logging.info("日志系统已启动")
    return log_file


def process_file(file_path):
    all_data = []
    try:
        logging.info(f"开始处理文件: {file_path}")
        with open(file_path, 'r', encoding='iso-8859-1') as file:
            for line in file:
                clean_line = line.strip()
                data_array = clean_line.split(',')
                if len(data_array) == 12:
                    all_data.append(data_array)
        logging.info(f"文件 {file_path} 成功读取，包含 {len(all_data)} 行有效数据")
    except Exception as e:
        logging.error(f"处理文件 {file_path} 时发生错误: {e}")
    return all_data


def process_folder(folder_path):
    logging.info(f"开始处理文件夹: {folder_path}")
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".txt"):
                file_path = os.path.join(root, file)
                logging.info(f"正在处理文件: {file_path}")

                # 读取并处理文件
                result = process_file(file_path)
                new_array = [[data_array[0], float(data_array[4]), float(data_array[3]), 0]
                             for data_array in result if len(data_array) > 4]

                # 如果 new_array 为空，跳过该文件
                if not new_array:
                    logging.warning(f"文件 {file_path} 数据为空，跳过处理。")
                    continue

                try:
                    # 找到 A 列的最大值
                    all_index_1_values = [item[2] for item in new_array]
                    max_value = max(all_index_1_values)
                    logging.info(f"文件 {file_path} 的 A 列最大值为: {max_value}")

                    # 处理数据并更新
                    for item in new_array:
                        item[3] = round((max_value - item[2]) * 2 / 0.085, 6)

                    # 输出文件路径
                    logging.info(f"root: {root}")
                    # 统一路径分隔符为 `/`
                    normalized_path = root.replace("\\", "/")
                    # 获取最后一层目录名
                    last_folder = os.path.basename(normalized_path)
                    logging.info(f"last_folder: {last_folder}")
                    output_file = os.path.join(root, f"correct.{last_folder}{file}")

                    # 保存处理结果
                    with open(output_file, 'w', encoding='utf-8') as out_file:
                        for item in new_array:
                            out_file.write(f"{item[1]}\t{item[3]}\n")
                    logging.info(f"文件 {file_path} 已成功保存为 {output_file}")
                except Exception as e:
                    logging.error(f"处理文件 {file_path} 时发生错误: {e}")


def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_path_label.config(text=f"选择的文件夹: {folder_path}")
        global selected_folder
        selected_folder = folder_path
        logging.info(f"用户选择了文件夹: {folder_path}")


def start_processing():
    if not selected_folder:
        messagebox.showerror("错误", "请先选择一个文件夹！")
        logging.warning("未选择文件夹，处理中断。")
        return

    logging.info("用户点击了开始处理按钮")
    process_folder(selected_folder)
    logging.info("所有文件已处理完成")
    messagebox.showinfo("完成", "所有文件已处理完成！")


# 检查试用时间
check_trial_period()

# 创建日志系统
log_file_path = setup_logging()

# 创建 GUI
root = tk.Tk()
root.title("批量文件处理程序")
root.geometry("500x200")

selected_folder = ""

# 文件夹选择按钮
select_button = tk.Button(root, text="选择文件夹", command=select_folder, width=20)
select_button.pack(pady=10)

# 显示选择的文件夹路径
folder_path_label = tk.Label(root, text="未选择文件夹", fg="blue", wraplength=400)
folder_path_label.pack()

# 开始按钮
start_button = tk.Button(root, text="开始处理", command=start_processing, width=20)
start_button.pack(pady=10)

# 日志路径显示
log_label = tk.Label(root, text=f"日志保存到: {log_file_path}", fg="green", wraplength=400)
log_label.pack(pady=10)

# 运行 GUI
root.mainloop()
