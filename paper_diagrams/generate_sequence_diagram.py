#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用 PIL 程序化生成缺陷检测顺序图（图4-3）。
保证所有6列参与者完整显示，无截断。
"""

import math
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\图4-3_缺陷检测顺序图.png")

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

FONT_HEAD  = load_font(24)   # 参与者标签
FONT_MSG   = load_font(20)   # 消息文字
FONT_NUM   = load_font(18)   # 序号

# ── 布局参数 ──────────────────────────────────────────────────────────────────
PARTICIPANTS = [
    "用户",
    "检测页面",
    "Flask后端",
    "YOLO模型",
    "数据库",
    "库存表",
]

MESSAGES = [
    # (from_idx, to_idx, label, is_return, is_self)
    (0, 1, "① 选择图片/拍照",           False, False),
    (1, 2, "② 上传图片",                False, False),
    (2, 2, "③ 保存图片文件",            False, True),
    (2, 3, "④ 执行缺陷检测",            False, False),
    (3, 2, "⑤ 返回检测结果",            True,  False),
    (2, 2, "⑥ 缺陷映射与等级判定",      False, True),
    (2, 4, "⑦ 写入检测历史",            False, False),
    (2, 5, "⑧ 更新库存数量",            False, False),
    (2, 1, "⑨ 返回结果图/类型/等级/置信度", True, False),
    (1, 0, "⑩ 显示检测结果",            True,  False),
]

# 画布
COL_SPACING = 210      # 列间距
LEFT_MARGIN  = 80
TOP_MARGIN   = 30
HEAD_HEIGHT  = 70      # 参与者框高度
HEAD_BOX_W   = 160
HEAD_BOX_H   = 44
MSG_STEP     = 58      # 消息间距（行高）
SELF_LOOP_W  = 50      # 自调用箭头宽度
TAIL_MARGIN  = 40      # 底部留白

N = len(PARTICIPANTS)
W = LEFT_MARGIN + (N - 1) * COL_SPACING + LEFT_MARGIN + 60
H = TOP_MARGIN + HEAD_HEIGHT + len(MESSAGES) * MSG_STEP + TAIL_MARGIN

# 各参与者的 x 坐标
col_x = [LEFT_MARGIN + i * COL_SPACING for i in range(N)]

# 生命线从参与者框底部延伸
LIFELINE_TOP    = TOP_MARGIN + HEAD_HEIGHT
LIFELINE_BOTTOM = H - TAIL_MARGIN

# ── 绘制工具 ──────────────────────────────────────────────────────────────────
def centered_text(draw, cx, cy, text, font, fill="black"):
    bbox = draw.textbbox((0, 0), text, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - w / 2, cy - h / 2), text, font=font, fill=fill)


def draw_arrowhead(draw, x, y, pointing_right: bool):
    """画水平箭头头部。"""
    size = 10
    d = 1 if pointing_right else -1
    draw.polygon([
        (x, y),
        (x - d * size, y - 5),
        (x - d * size, y + 5),
    ], fill="black")


def draw_message(draw, from_idx, to_idx, label, is_return, is_self, y):
    """在给定 y 坐标绘制一条消息。"""
    x0 = col_x[from_idx]
    x1 = col_x[to_idx]

    line_style = (6, 4) if is_return else None  # 虚线 or 实线

    if is_self:
        # 自调用：小矩形回路
        lx = x0 + SELF_LOOP_W
        draw.line((x0, y, lx, y), fill="black", width=2)
        draw.line((lx, y, lx, y + 28), fill="black", width=2)
        draw.line((lx, y + 28, x0, y + 28), fill="black", width=2)
        draw_arrowhead(draw, x0, y + 28, False)
        # 文字在右侧
        bbox = draw.textbbox((0, 0), label, font=FONT_MSG)
        tw = bbox[2] - bbox[0]
        draw.text((lx + 6, y + 4), label, font=FONT_MSG, fill="black")
        return

    # 计算方向
    pointing_right = (x1 > x0)
    arrow_end = x1

    # 画线（实线 or 虚线）
    if line_style:
        # 手动虚线
        dx = x1 - x0
        total = abs(dx)
        sign  = 1 if dx > 0 else -1
        dash, gap = line_style
        pos = 0
        while pos < total:
            seg_start = x0 + sign * pos
            seg_end   = x0 + sign * min(pos + dash, total)
            draw.line((seg_start, y, seg_end, y), fill="black", width=2)
            pos += dash + gap
    else:
        draw.line((x0, y, x1, y), fill="black", width=2)

    # 箭头
    draw_arrowhead(draw, arrow_end, y, pointing_right)

    # 文字（线条上方居中）
    mid_x = (x0 + x1) / 2
    bbox = draw.textbbox((0, 0), label, font=FONT_MSG)
    tw = bbox[2] - bbox[0]
    draw.text((mid_x - tw / 2, y - 22), label, font=FONT_MSG, fill="black")


# ── 主绘图 ────────────────────────────────────────────────────────────────────
def main():
    img  = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # 1. 参与者框 + 标签
    for i, name in enumerate(PARTICIPANTS):
        cx = col_x[i]
        y0 = TOP_MARGIN
        x0 = cx - HEAD_BOX_W / 2
        x1 = cx + HEAD_BOX_W / 2
        y1 = y0 + HEAD_BOX_H
        draw.rectangle((x0, y0, x1, y1), outline="black", width=2, fill="white")
        centered_text(draw, cx, (y0 + y1) / 2, name, FONT_HEAD)

    # 2. 生命线（虚线竖线）
    for cx in col_x:
        y = LIFELINE_TOP
        while y < LIFELINE_BOTTOM:
            draw.line((cx, y, cx, min(y + 10, LIFELINE_BOTTOM)), fill="black", width=1)
            y += 20

    # 3. 消息
    msg_y_start = LIFELINE_TOP + MSG_STEP
    for step_idx, (fi, ti, label, is_ret, is_self) in enumerate(MESSAGES):
        y = msg_y_start + step_idx * MSG_STEP
        draw_message(draw, fi, ti, label, is_ret, is_self, y)

    img.save(OUT_PATH, dpi=(150, 150))
    print(f"生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
