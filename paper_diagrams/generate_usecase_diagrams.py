#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成标准 UML 用例图：
- 系统边界矩形
- 带箭头的关联线（actor → usecase）
- 紧凑布局，无多余空白
- 字体大小与正文一致（宋体 28pt）
"""

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT_DIR = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\usecase_generated")
OUT_DIR.mkdir(parents=True, exist_ok=True)

FONT_PATHS = [
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
]


def load_font(size: int):
    for path in FONT_PATHS:
        p = Path(path)
        if p.exists():
            return ImageFont.truetype(str(p), size)
    return ImageFont.load_default()


# 字体大小：28 与正文近似一致
FONT_LABEL = load_font(28)
FONT_SYSTEM = load_font(26)


def centered_text(draw, cx, cy, text, font, fill="black"):
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    draw.text((cx - w / 2, cy - h / 2), text, font=font, fill=fill)


def draw_actor(draw, cx, top_y, label):
    """绘制小人，top_y 是头顶 y 坐标。"""
    head_r = 20
    # 头
    draw.ellipse(
        (cx - head_r, top_y, cx + head_r, top_y + head_r * 2),
        outline="black", width=2
    )
    body_top = top_y + head_r * 2
    body_bot = body_top + 60
    # 身体
    draw.line((cx, body_top, cx, body_bot), fill="black", width=2)
    # 手臂
    arm_y = body_top + 20
    draw.line((cx - 28, arm_y, cx + 28, arm_y), fill="black", width=2)
    # 腿
    draw.line((cx, body_bot, cx - 24, body_bot + 36), fill="black", width=2)
    draw.line((cx, body_bot, cx + 24, body_bot + 36), fill="black", width=2)
    # 标签
    centered_text(draw, cx, body_bot + 36 + 22, label, FONT_LABEL)


def draw_usecase_ellipse(draw, cx, cy, rx, ry, text):
    """绘制用例椭圆+文字。"""
    draw.ellipse(
        (cx - rx, cy - ry, cx + rx, cy + ry),
        outline="black", width=2, fill="white"
    )
    centered_text(draw, cx, cy, text, FONT_LABEL)


def draw_arrow_line(draw, x0, y0, x1, y1, ell_rx, ell_ry, ell_cx, ell_cy):
    """
    从 (x0,y0) → 用例椭圆靠 actor 一侧边缘，带箭头。
    正确公式：找从椭圆中心指向 actor 方向的参数角，
    即 t = atan2(rx*(y0-cy), ry*(x0-cx))
    """
    # actor 相对于椭圆中心的方向 → 这一侧才是箭头终点
    t = math.atan2(ell_rx * (y0 - ell_cy), ell_ry * (x0 - ell_cx))
    ex = ell_cx + ell_rx * math.cos(t)
    ey = ell_cy + ell_ry * math.sin(t)

    # 画线
    draw.line((x0, y0, ex, ey), fill="black", width=2)

    # 箭头头部（朝向椭圆方向）
    head_len = 14
    head_angle = 0.45
    line_angle = math.atan2(ey - y0, ex - x0)
    for sign in (+1, -1):
        ax = ex - head_len * math.cos(line_angle - sign * head_angle)
        ay = ey - head_len * math.sin(line_angle - sign * head_angle)
        draw.line((ex, ey, ax, ay), fill="black", width=2)


def user_usecase():
    # 画布参数
    W, H = 900, 680
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # 系统边界
    box_x1, box_y1, box_x2, box_y2 = 310, 30, 860, 650
    draw.rectangle((box_x1, box_y1, box_x2, box_y2), outline="black", width=2)
    centered_text(draw, (box_x1 + box_x2) / 2, box_y1 + 14, "木质板材缺陷检测系统", FONT_SYSTEM)

    # Actor
    actor_cx = 120
    actor_top = 250
    actor_body_bot = actor_top + 20 * 2 + 60
    actor_connect_y = actor_top + 20 * 2 + 20   # 手臂高度（腰部）
    draw_actor(draw, actor_cx, actor_top, "用户")

    # 用例椭圆参数
    uc_cx = 580
    ell_rx, ell_ry = 130, 32
    labels = ["注册登录", "缺陷检测", "检测历史", "日志报表", "销售管理", "帮助中心", "个人中心"]
    n = len(labels)
    # 均匀分布在系统边界内
    y_start = box_y1 + 60
    y_end = box_y2 - 60
    step = (y_end - y_start) / (n - 1)

    for i, label in enumerate(labels):
        cy = y_start + i * step
        draw_usecase_ellipse(draw, uc_cx, cy, ell_rx, ell_ry, label)
        draw_arrow_line(draw, actor_cx, actor_connect_y, uc_cx, cy,
                        ell_rx, ell_ry, uc_cx, cy)

    out = OUT_DIR / "图3-1_普通用户用例图.png"
    img.save(out, dpi=(150, 150))
    return out


def admin_usecase():
    W, H = 900, 560
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # 系统边界
    box_x1, box_y1, box_x2, box_y2 = 310, 30, 860, 530
    draw.rectangle((box_x1, box_y1, box_x2, box_y2), outline="black", width=2)
    centered_text(draw, (box_x1 + box_x2) / 2, box_y1 + 14, "木质板材缺陷检测系统", FONT_SYSTEM)

    # Actor
    actor_cx = 120
    actor_top = 190
    actor_connect_y = actor_top + 20 * 2 + 20
    draw_actor(draw, actor_cx, actor_top, "管理员")

    uc_cx = 580
    ell_rx, ell_ry = 130, 32
    labels = ["管理员登录", "用户管理", "系统设置", "模型切换", "日志查看"]
    n = len(labels)
    y_start = box_y1 + 60
    y_end = box_y2 - 60
    step = (y_end - y_start) / (n - 1)

    for i, label in enumerate(labels):
        cy = y_start + i * step
        draw_usecase_ellipse(draw, uc_cx, cy, ell_rx, ell_ry, label)
        draw_arrow_line(draw, actor_cx, actor_connect_y, uc_cx, cy,
                        ell_rx, ell_ry, uc_cx, cy)

    out = OUT_DIR / "图3-2_管理员用例图.png"
    img.save(out, dpi=(150, 150))
    return out


def main():
    p1 = user_usecase()
    p2 = admin_usecase()
    print(f"生成: {p1}")
    print(f"生成: {p2}")


if __name__ == "__main__":
    main()
