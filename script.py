import cv2
import numpy as np
import sys

def detect_and_overlay(image_path, overlay_image_path, output_path):
    """
    检测图像中最大的红色矩形区域，缩小至一半，只保留右侧半部分，
    并在该区域上叠加指定的图片，保存处理后的图像。
    """
    print(">>> 开始执行 detect_and_overlay 函数")
    print(f">>> 输入图片路径: {image_path}")
    print(f">>> 叠加图片路径: {overlay_image_path}")
    print(f">>> 输出图片路径: {output_path}")

    # 加载输入图片
    print(">>> 加载输入图片...")
    image = cv2.imread(image_path)
    if image is None:
        print(f"错误: 无法加载输入图片 {image_path}")
        sys.exit(1)

    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    print(">>> 转换输入图片为 HSV 颜色空间完成")

    # 定义红色的HSV范围
    print(">>> 定义红色的 HSV 范围...")
    lower_red1 = np.array([0, 120, 70])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([170, 120, 70])
    upper_red2 = np.array([180, 255, 255])

    # 创建红色掩膜
    print(">>> 创建红色掩膜...")
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = mask1 + mask2
    print(">>> 红色掩膜创建完成")

    # 找到红色区域的轮廓
    print(">>> 查找红色区域轮廓...")
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    print(f">>> 找到 {len(contours)} 个轮廓")

    # 初始化变量，用于存储最大的矩形
    largest_rectangle = None
    largest_area = 0

    # 遍历所有轮廓，找到面积最大的矩形
    print(">>> 遍历轮廓，寻找最大红色矩形区域...")
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = w * h
        print(f"检测到矩形: x={x}, y={y}, 宽度={w}, 高度={h}, 面积={area}")
        if area > largest_area:
            largest_rectangle = (x, y, w, h)
            largest_area = area

    if not largest_rectangle:
        print("错误: 未检测到红色矩形区域")
        sys.exit(1)

    x, y, w, h = largest_rectangle
    print(f"最大矩形区域: x={x}, y={y}, 宽度={w}, 高度={h}, 面积={largest_area}")

    # 缩小矩形区域至宽度的一半，只保留右侧部分
    print(">>> 缩小矩形区域，保留右侧部分...")
    new_width = w // 2
    x = x + new_width
    w = new_width
    print(f"缩小后的矩形区域: x={x}, y={y}, 宽度={w}, 高度={h}")

    # 加载需要叠加的图片
    print(f">>> 加载叠加图片 {overlay_image_path}...")
    overlay = cv2.imread(overlay_image_path, cv2.IMREAD_UNCHANGED)
    if overlay is None:
        print(f"错误: 无法加载叠加图片 {overlay_image_path}")
        sys.exit(1)
    print(">>> 叠加图片加载完成")

    # 调整叠加图片的大小与检测到的矩形区域匹配
    print(">>> 调整叠加图片的大小与矩形区域匹配...")
    overlay_resized = cv2.resize(overlay, (w, h))
    print(f"叠加图片调整完成: 新宽度={w}, 新高度={h}")

    # 如果叠加图片有透明通道
    if overlay_resized.shape[2] == 4:
        print(">>> 检测到叠加图片包含透明通道，进行透明度处理...")
        alpha_channel = overlay_resized[:, :, 3] / 255.0
        alpha_overlay = alpha_channel[:, :, np.newaxis]
        alpha_overlay = np.repeat(alpha_overlay, 3, axis=2)
        alpha_background = 1.0 - alpha_overlay

        # 提取叠加图片的RGB通道
        overlay_rgb = overlay_resized[:, :, :3]

        # 将叠加图片与原图融合
        print(">>> 叠加图片与原图融合...")
        for c in range(3):
            image[y:y + h, x:x + w, c] = (
                alpha_overlay[:, :, c] * overlay_rgb[:, :, c] +
                alpha_background[:, :, c] * image[y:y + h, x:x + w, c]
            )
        print(">>> 融合完成")
    else:
        print(">>> 叠加图片没有透明通道，直接覆盖...")
        image[y:y + h, x:x + w] = overlay_resized

    # 保存输出图片
    print(">>> 保存处理后的图片...")
    cv2.imwrite(output_path, image)
    print(f"处理后的图片已保存到 {output_path}")

if __name__ == "__main__":
    print(">>> 脚本开始执行...")
    if len(sys.argv) != 4:
        print("用法: python script.py <输入图片路径> <叠加图片路径> <输出图片路径>")
        sys.exit(1)

    input_image_path = sys.argv[1]
    overlay_image_path = sys.argv[2]
    output_image_path = sys.argv[3]

    print(">>> 参数检查完成")
    print(f">>> 输入图片路径: {input_image_path}")
    print(f">>> 叠加图片路径: {overlay_image_path}")
    print(f">>> 输出图片路径: {output_image_path}")

    detect_and_overlay(input_image_path, overlay_image_path, output_image_path)
    print(">>> 脚本执行完成")
