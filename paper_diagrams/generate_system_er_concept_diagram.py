#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成木质板材缺陷检测系统的概念 E-R 图。

说明：
1. 图形风格参考用户提供的示例：黑底、白色实体框、白色关系菱形。
2. 图中既体现系统真实数据库实体，也保留论文中更容易讲清业务的角色层。
3. 本图适合直接作为论文中的“系统数据库 E-R 图”候选版本。
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


OUT_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\图4-7_系统数据库E-R图_参考样式.png")

FONT_PATHS = [
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
]


def load_font(size: int):
    """优先加载常见中文字体，保证论文图中的中文显示稳定。"""
    for font_path in FONT_PATHS:
        path = Path(font_path)
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


FONT_BOX = load_font(34)
FONT_DIAMOND = load_font(30)
FONT_TITLE = load_font(40)

W, H = 1400, 920
BG = "black"
FG = "white"
TEXT = "black"
LINE_W = 4


def draw_centered_text(draw, box, text, font, fill=TEXT):
    """在指定区域中居中绘制文字。"""
    x0, y0, x1, y1 = box
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = x0 + (x1 - x0 - tw) / 2
    ty = y0 + (y1 - y0 - th) / 2 - 2
    draw.text((tx, ty), text, font=font, fill=fill)


def draw_entity(draw, center, label, w=220, h=84):
    """绘制实体矩形。"""
    cx, cy = center
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    draw.rectangle(box, outline=FG, fill=FG, width=LINE_W)
    draw_centered_text(draw, box, label, FONT_BOX)
    return box


def draw_relation(draw, center, label, w=120, h=92):
    """绘制关系菱形。"""
    cx, cy = center
    points = [
        (cx, cy - h // 2),
        (cx + w // 2, cy),
        (cx, cy + h // 2),
        (cx - w // 2, cy),
    ]
    draw.polygon(points, outline=FG, fill=FG, width=LINE_W)
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    draw_centered_text(draw, box, label, FONT_DIAMOND)
    return points


def mid_left(box):
    return (box[0], (box[1] + box[3]) // 2)


def mid_right(box):
    return (box[2], (box[1] + box[3]) // 2)


def mid_top(box):
    return ((box[0] + box[2]) // 2, box[1])


def mid_bottom(box):
    return ((box[0] + box[2]) // 2, box[3])


def diamond_left(points):
    return points[3]


def diamond_right(points):
    return points[1]


def diamond_top(points):
    return points[0]


def diamond_bottom(points):
    return points[2]


def connect(draw, p1, p2):
    """绘制关系连接线。"""
    draw.line([p1, p2], fill=FG, width=LINE_W)


def connect_polyline(draw, points):
    """绘制由水平/垂直线段组成的折线，避免出现斜线。"""
    draw.line(points, fill=FG, width=LINE_W)


def main():
    img = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # 标题留在画布顶部，论文裁图时也更方便保留。
    title_box = (0, 18, W, 72)
    draw_centered_text(draw, title_box, "系统数据库 E-R 图", FONT_TITLE, fill=FG)

    # 顶部角色层
    admin_box = draw_entity(draw, (240, 130), "管理员")
    role_diamond = draw_relation(draw, (700, 130), "角色")
    user_role_box = draw_entity(draw, (1160, 130), "普通用户")

    # 中心主实体
    left_match = draw_relation(draw, (400, 280), "对应")
    right_match = draw_relation(draw, (1000, 280), "对应")
    user_info_box = draw_entity(draw, (700, 320), "用户信息")

    # 中部业务实体
    log_rel = draw_relation(draw, (420, 500), "记录")
    exec_rel = draw_relation(draw, (700, 500), "执行")
    config_rel = draw_relation(draw, (980, 500), "配置")
    work_log_box = draw_entity(draw, (220, 660), "工作日志")
    detect_box = draw_entity(draw, (700, 660), "检测记录")
    settings_box = draw_entity(draw, (1180, 660), "系统设置")

    # 底部业务结果实体
    affect_rel = draw_relation(draw, (700, 800), "更新")
    deduct_rel = draw_relation(draw, (980, 800), "扣减")
    inventory_box = draw_entity(draw, (360, 800), "库存信息")
    sales_box = draw_entity(draw, (1220, 800), "销售记录")

    # 连接线：角色到用户信息
    connect(draw, mid_right(admin_box), diamond_left(role_diamond))
    connect(draw, diamond_right(role_diamond), mid_left(user_role_box))
    connect_polyline(draw, [mid_bottom(admin_box), (mid_bottom(admin_box)[0], diamond_left(left_match)[1]), diamond_left(left_match)])
    connect_polyline(draw, [diamond_right(left_match), (diamond_right(left_match)[0], mid_left(user_info_box)[1]), mid_left(user_info_box)])
    connect_polyline(draw, [mid_bottom(user_role_box), (mid_bottom(user_role_box)[0], diamond_right(right_match)[1]), diamond_right(right_match)])
    connect_polyline(draw, [diamond_left(right_match), (diamond_left(right_match)[0], mid_right(user_info_box)[1]), mid_right(user_info_box)])

    # 连接线：用户信息到日志、检测、设置
    connect_polyline(draw, [mid_left(user_info_box), (490, mid_left(user_info_box)[1]), (490, diamond_right(log_rel)[1]), diamond_right(log_rel)])
    connect_polyline(draw, [diamond_left(log_rel), (diamond_left(log_rel)[0], mid_top(work_log_box)[1]), mid_top(work_log_box)])
    connect(draw, mid_bottom(user_info_box), diamond_top(exec_rel))
    connect(draw, diamond_bottom(exec_rel), mid_top(detect_box))
    connect_polyline(draw, [mid_right(detect_box), (840, mid_right(detect_box)[1]), (840, diamond_left(config_rel)[1]), diamond_left(config_rel)])
    connect_polyline(draw, [diamond_right(config_rel), (1040, diamond_right(config_rel)[1]), (1040, mid_left(settings_box)[1]), mid_left(settings_box)])

    # 连接线：检测记录到库存，库存到销售
    connect(draw, mid_bottom(detect_box), diamond_top(affect_rel))
    connect(draw, diamond_left(affect_rel), mid_right(inventory_box))
    connect(draw, diamond_right(affect_rel), diamond_left(deduct_rel))
    connect(draw, diamond_right(deduct_rel), mid_left(sales_box))

    img.save(OUT_PATH, dpi=(180, 180))
    print(f"已生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
