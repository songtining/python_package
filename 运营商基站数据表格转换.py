import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from tkinter.ttk import Progressbar, Style
from datetime import datetime


def check_trial_expiry():
    """检查试用期是否已过期"""
    expiry_datetime = datetime(2025, 1, 13, 20, 59, 59)  # 设置试用到期日期和时间
    now = datetime.now()
    if now > expiry_datetime:
        messagebox.showerror("试用已过期", "试用期已结束，请联系管理员获取正式版本。")
        root.destroy()  # 关闭程序
    else:
        remaining_time = expiry_datetime - now
        days_left = remaining_time.days
        hours, remainder = divmod(remaining_time.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        trial_label.config(
            text=f"试用期剩余时间: {days_left}天 {hours}小时 {minutes}分 {seconds}秒"
        )
        root.after(1000, check_trial_expiry)  # 每秒钟更新一次


# 去重处理（按 LAC 和 CI 去重）
def deduplicate(df):
    """按 LAC 和 CI 去重，优先保留经纬度不为 0 的数据"""
    deduplicated = []
    grouped = df.groupby(["CGI（必填，CGI序列或运营商名称）", "LAC（必填）", "CI（必填）"])

    for _, group in grouped:
        # 检查是否存在经纬度不为 0 的行
        valid_rows = group[(group["基站经度"] != 0) & (group["基站纬度"] != 0)]
        if not valid_rows.empty:
            # 如果存在经纬度不为 0 的数据，优先保留第一条
            deduplicated.append(valid_rows.iloc[0])
        else:
            # 如果所有行的经纬度均为 0，随机保留一条
            deduplicated.append(group.iloc[0])

    return pd.DataFrame(deduplicated)

def process_excel(input_file, output_file, progress_var, progress_label, output_label):
    # 初始化一个空的 DataFrame 来存储表格2数据
    columns_table2 = ["CGI（必填，CGI序列或运营商名称）", "LAC（必填）", "CI（必填）",
                      "基站类型（必填）", "基站经度", "基站纬度", "基站名称", "基站地址"]
    result_df = pd.DataFrame(columns=columns_table2)

    # 读取表格1中的所有 sheet 页
    sheets = pd.ExcelFile(input_file)
    total_sheets = len(sheets.sheet_names)
    for idx, sheet_name in enumerate(sheets.sheet_names, start=1):
        if sheet_name == 'WIFI':
            continue
        # 更新进度条
        progress_var.set(int((idx / total_sheets) * 100))
        progress_label.config(text=f"正在处理: {sheet_name} ({idx}/{total_sheets})")
        root.update_idletasks()

        # 读取每个 sheet 页
        sheet_df = pd.read_excel(input_file, sheet_name=sheet_name)

        # CGI 值来源于 sheet 页名称的前两个字符
        cgi_value = sheet_name[:2]

        # 检查列的数量是否足够
        if sheet_df.shape[1] < 3:
            print(f"Sheet 页 {sheet_name} 列数量不足，跳过处理。")
            continue

        # 提取 LAC（第二列）和 CI（第三列）
        lac_col = sheet_df.iloc[:, 1]  # 第二列
        ci_col = sheet_df.iloc[:, 2]  # 第三列

        # 提取经纬度列（列名固定为 lng 和 lat）
        if "lng" not in sheet_df.columns or "lat" not in sheet_df.columns:
            print(f"Sheet 页 {sheet_name} 缺少经纬度列，跳过处理。")
            continue

        # 提取经纬度列
        lng_col = sheet_df["lng"]
        lat_col = sheet_df["lat"]

        # 将数据按照规则转换为表格2格式
        transformed_df = pd.DataFrame({
            "CGI（必填，CGI序列或运营商名称）": cgi_value,
            "LAC（必填）": lac_col,
            "CI（必填）": ci_col,
            "基站类型（必填）": "运营商基站",  # 默认填写（可根据需要修改）
            "基站经度": lng_col,
            "基站纬度": lat_col,
            "基站名称": None,  # 没有对应列，填充为空值
            "基站地址": None  # 没有对应列，填充为空值
        })

        # 汇总到结果 DataFrame
        result_df = pd.concat([result_df, transformed_df], ignore_index=True)

    # 去重处理（按 LAC 和 CI 去重）
    result_df = deduplicate(result_df)

    # 对最终的数据按 LAC（必填） 排序
    result_df = result_df.sort_values(by=["CGI（必填，CGI序列或运营商名称）"], ascending=True)

    # 将经纬度列转为字符串格式，保留原始小数位数
    result_df["基站经度"] = result_df["基站经度"].apply(lambda x: str(x) if pd.notnull(x) else "")
    result_df["基站纬度"] = result_df["基站纬度"].apply(lambda x: str(x) if pd.notnull(x) else "")

    # 写入到本地文件
    result_df.to_excel(output_file, index=False, engine="openpyxl")

    # 设置输出的列宽
    workbook = load_workbook(output_file)
    worksheet = workbook.active

    # 设置每列的宽度为 35
    for i, column in enumerate(columns_table2, start=1):
        col_letter = get_column_letter(i)
        worksheet.column_dimensions[col_letter].width = 35  # 宽度设置为 35

    # 保存文件
    workbook.save(output_file)
    print(f"数据已成功处理并保存到 {output_file}，并设置了列宽。")
    progress_var.set(100)
    progress_label.config(text="处理完成！")
    output_label.config(text=f"文件已保存到: {output_file}")
    messagebox.showinfo("完成", f"处理后的文件已成功处理并保存到 {output_file}")


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_label.config(text=f"已选择文件: {file_path}")
        process_button.config(state=tk.NORMAL)
        return file_path
    else:
        file_label.config(text="未选择文件")
        process_button.config(state=tk.DISABLED)


def start_processing():
    input_file = file_label.cget("text").replace("已选择文件: ", "")
    if not input_file or not os.path.isfile(input_file):
        messagebox.showwarning("警告", "未选择有效的输入文件！")
        return

    # 输出文件路径：与输入文件同级目录，文件名格式为“输入文件名_基站信息导入.xlsx”
    output_file = os.path.join(
        os.path.dirname(input_file),
        f"{os.path.splitext(os.path.basename(input_file))[0]}_基站信息导入.xlsx"
    )

    try:
        process_excel(input_file, output_file, progress_var, progress_label, output_label)
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时出错: {e}")


# 创建 GUI 界面
root = tk.Tk()
root.title("Excel文件处理工具")
root.geometry("600x450")
root.resizable(False, False)

# 样式设置
style = Style()
style.configure("TProgressbar", thickness=15)

# 文件选择部分
file_label = tk.Label(root, text="点击下方按钮选择Excel文件", font=("Arial", 12))
file_label.pack(pady=10)

select_button = tk.Button(root, text="选择文件", command=select_file)
select_button.pack(pady=10)

# 进度条
progress_var = tk.IntVar()
progress_bar = Progressbar(root, orient="horizontal", length=400, mode="determinate", variable=progress_var)
progress_bar.pack(pady=10)

progress_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
progress_label.pack(pady=5)

# 输出文件路径显示
output_label = tk.Label(root, text="", font=("Arial", 10), fg="blue", wraplength=500, justify="center")
output_label.pack(pady=10)

# 剩余试用期显示
trial_label = tk.Label(root, text="", font=("Arial", 10), fg="red")
trial_label.pack(pady=5)

# 检查试用期
check_trial_expiry()

# 处理按钮
process_button = tk.Button(root, text="开始处理", state=tk.DISABLED, command=start_processing)
process_button.pack(pady=10)

# 主循环
root.mainloop()
