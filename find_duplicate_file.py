import os
import sys
import hashlib
import pandas as pd


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
        print(f"无法读取文件: {file_path}，错误: {e}")
        return None
    return hash_md5.hexdigest()


def find_version_conflicts(root_dir, output_file):
    """
    找出文件名相同但 MD5 不同的文件，并导出到 Excel 文件。

    参数:
    - root_dir: 要扫描的文件夹路径。
    - output_file: 输出的 Excel 文件路径。
    """
    file_info = []  # 保存文件信息

    # 遍历文件夹，收集所有文件的信息
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            md5_hash = calculate_md5(file_path)
            if md5_hash:
                file_info.append({"filename": filename, "file_path": file_path, "md5": md5_hash})

    # 转换为 DataFrame 方便处理
    df = pd.DataFrame(file_info)

    # 按文件名分组，查找 MD5 不同的文件
    conflicts = []
    grouped = df.groupby("filename")
    for filename, group in grouped:
        if len(group["md5"].unique()) > 1:  # 如果同名文件有不同的 MD5 值
            conflicts.append({"filename": filename, "file_paths": group["file_path"].tolist(), "md5s": group["md5"].tolist()})

    # 保存冲突文件名到 Excel
    if conflicts:
        output_data = []
        for conflict in conflicts:
            for path, md5 in zip(conflict["file_paths"], conflict["md5s"]):
                output_data.append({"文件名": conflict["filename"], "文件路径": path, "MD5": md5})

        output_df = pd.DataFrame(output_data)
        output_df.to_excel(output_file, index=False)
        print(f"版本冲突文件已导出到: {output_file}")
    else:
        print("未发现文件名相同但 MD5 不同的文件。")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: version_conflicts_finder.exe <要扫描的文件夹路径>")
        sys.exit(1)

    root_dir = sys.argv[1]
    output_file = "version_conflicts_with_md5.xlsx"

    if os.path.exists(root_dir):
        print("开始扫描文件夹...")
        find_version_conflicts(root_dir, output_file)
        print("扫描完成！")
    else:
        print("输入的文件夹路径不存在，请检查后重试。")
