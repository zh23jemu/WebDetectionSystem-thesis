import cv2
from ultralytics import YOLO
import os
from datetime import datetime

# 全局变量缓存模型
_model = None
_current_model_path = None


def load_model(model_path):
    global _model, _current_model_path
    if _model is None or model_path != _current_model_path:
        print(f"Loading Model: {model_path} ...")
        try:
            _model = YOLO(model_path)
            _current_model_path = model_path
        except Exception as e:
            print(f"Error loading model: {e}")
            return None
    return _model


def process_image(image_path, save_dir, model_path, conf_threshold=0.25):
    """
    统一接口：处理单张图片
    """
    model = load_model(model_path)
    if not model:
        return {'error': '模型加载失败'}

    # 读取图片
    img = cv2.imread(image_path)
    if img is None:
        return {'error': '无法读取图片'}

    # 确保 conf 是 float 类型
    try:
        conf = float(conf_threshold)
    except:
        conf = 0.25

    # 推理
    results = model(img, conf=conf)

    detected_objects = []
    has_defect = False
    max_conf = 0.0

    res_plotted = img

    for r in results:
        res_plotted = r.plot()
        for box in r.boxes:
            cls_id = int(box.cls[0])
            cls_name = model.names[cls_id]
            conf_val = float(box.conf[0])

            # 直接添加类别名，后续去重
            detected_objects.append(cls_name)

            if conf_val > max_conf:
                max_conf = conf_val
            has_defect = True

    # 保存结果
    filename = os.path.basename(image_path)
    save_name = f"res_{int(datetime.now().timestamp())}_{filename}"
    save_path = os.path.join(save_dir, save_name)
    cv2.imwrite(save_path, res_plotted)

    # 格式化缺陷字符串：
    # 1. set() 去重 (一张图可能有多个同类框)
    # 2. join() 用英文逗号连接
    if detected_objects:
        defects_str = ",".join(list(set(detected_objects)))
    else:
        defects_str = "无缺陷"

    return {
        'filepath': save_name,
        'defects': defects_str,
        'count': len(detected_objects),  # 这里返回的是框的总数量
        'is_clean': not has_defect,
        'confidence': round(max_conf, 2)
    }