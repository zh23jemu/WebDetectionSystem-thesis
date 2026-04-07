import cv2
import os


def resize_image(image_path, output_path, max_width=800):
    """
    调整图片大小，避免Web端显示过大
    """
    img = cv2.imread(image_path)
    if img is None:
        return False

    h, w = img.shape[:2]

    if w > max_width:
        ratio = max_width / w
        new_h = int(h * ratio)
        new_w = max_width
        resized_img = cv2.resize(img, (new_w, new_h))
        cv2.imwrite(output_path, resized_img)
        return True

    # 如果不需要缩放，直接保存原图
    cv2.imwrite(output_path, img)
    return True


def convert_to_jpg(image_path):
    """
    将非jpg格式转换为jpg
    """
    filename, ext = os.path.splitext(image_path)
    if ext.lower() not in ['.jpg', '.jpeg']:
        img = cv2.imread(image_path)
        new_path = filename + ".jpg"
        cv2.imwrite(new_path, img)
        return new_path
    return image_path