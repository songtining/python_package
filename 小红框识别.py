import cv2
import numpy as np
import sys

def detect_and_overlay(image_path, overlay_image_path, output_path):
    """
    检测图像中最大的红色矩形区域，缩小至一半，只保留右侧半部分，
    并在该区域上叠加指定的图片，保存处理后的图像。
    """
    # 加载输入图片
    image = cv2.imread(image_path)
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

    # 定义红色的HSV范围
    lower_red1 = np.array([0, 120, 70])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([170, 120, 70])
    upper_red2 = np.array([180, 255, 255])

    # 创建红色掩膜
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = mask1 + mask2

    # 找到红色区域的轮廓
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 初始化变量，用于存储最大的矩形
    largest_rectangle = None
    largest_area = 0

    # 遍历所有轮廓，找到面积最大的矩形
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = w * h

        # 只考虑面积较大的矩形
        if area > largest_area:
            largest_rectangle = (x, y, w, h)
            largest_area = area

    # 如果检测到最大的红色矩形区域
    if largest_rectangle:
        x, y, w, h = largest_rectangle

        # 缩小矩形区域至宽度的一半，只保留右侧部分
        new_width = w // 2
        x = x + new_width  # 右移起始点
        w = new_width  # 更新宽度

        # 打印检测到的矩形坐标
        print(f"检测到的红色矩形区域（缩小后右侧部分）: x={x}, y={y}, 宽度={w}, 高度={h}")

        # 加载需要叠加的图片
        overlay = cv2.imread(overlay_image_path, cv2.IMREAD_UNCHANGED)

        # 调整叠加图片的大小与检测到的矩形区域匹配
        overlay_resized = cv2.resize(overlay, (w, h))

        # 如果叠加图片有透明通道
        if overlay_resized.shape[2] == 4:  # 检查是否有alpha通道
            alpha_channel = overlay_resized[:, :, 3] / 255.0  # 归一化alpha值
            alpha_overlay = cv2.merge((alpha_channel, alpha_channel, alpha_channel))
            alpha_background = 1.0 - alpha_overlay

            # 提取叠加图片的RGB通道
            overlay_rgb = overlay_resized[:, :, :3]

            # 将叠加图片与原图融合
            for c in range(3):  # 分别处理每个通道
                image[y:y + h, x:x + w, c] = (
                    alpha_overlay * overlay_rgb[:, :, c] +
                    alpha_background * image[y:y + h, x:x + w, c]
                )
        else:
            # 如果叠加图片没有透明通道，直接覆盖
            image[y:y + h, x:x + w] = overlay_resized

    # 保存输出图片
    cv2.imwrite(output_path, image)
    print(f"处理后的图片已保存到 {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("用法: python script.py <输入图片路径> <叠加图片路径> <输出图片路径>")
        sys.exit(1)

    input_image_path = sys.argv[1]
    overlay_image_path = sys.argv[2]
    output_image_path = sys.argv[3]

    detect_and_overlay(input_image_path, overlay_image_path, output_image_path)