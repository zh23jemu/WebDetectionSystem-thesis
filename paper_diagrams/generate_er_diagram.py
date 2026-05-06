#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 Chen 符号法实体属性图（图4-6 主要实体属性图）。
- 实体：矩形
- 属性：等大椭圆
- 每个实体的属性分布在上下两侧，避免线条交叉
- 6 个实体按 3×2 网格布局
"""

import math
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\图4-6_主要实体属性图.png")

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

FONT_ENTITY = load_font(26)
FONT_ATTR   = load_font(20)

# 实体定义：名称 + 属性列表
ENTITIES = [
    ("User",              ["id", "用户名", "密码", "角色"]),
    ("DetectionHistory",  ["id", "文件名", "结果图片", "上传日期", "缺陷类型", "置信度", "用户id", "等级"]),
    ("WorkLog",           ["id", "用户id", "登录时间", "登出时间"]),
    ("SalesRecord",       ["id", "客户名", "产品等级", "数量", "总价", "日期"]),
    ("SystemSettings",    ["id", "键名", "值"]),
    ("Inventory",         ["等级", "数量"]),
]

# ── 画布参数 ──────────────────────────────────────────────────────────────────
COLS, ROWS = 3, 2          # 网格：3列2行
CELL_W     = 580           # 每个格子宽度
CELL_H     = 380           # 每个格子高度
MARGIN_X   = 40
MARGIN_Y   = 40
W = COLS * CELL_W + 2 * MARGIN_X
H = ROWS * CELL_H + 2 * MARGIN_Y

# 实体矩形最小尺寸（实际宽度会根据文字自动扩展）
ENT_MIN_W, ENT_H = 180, 50
ENT_PAD_X = 24   # 文字左右各留的内边距

# 属性椭圆（全部统一大小）
ATT_RX, ATT_RY = 68, 22

# ── 绘制工具 ──────────────────────────────────────────────────────────────────
def centered_text(draw, cx, cy, text, font, fill="black"):
    bbox = draw.textbbox((0, 0), text, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - w / 2, cy - h / 2), text, font=font, fill=fill)


def get_ent_w(draw, name):
    """根据文字宽度动态计算实体矩形宽度。"""
    bbox = draw.textbbox((0, 0), name, font=FONT_ENTITY)
    text_w = bbox[2] - bbox[0]
    return max(ENT_MIN_W, text_w + ENT_PAD_X * 2)


def draw_entity(draw, cx, cy, name):
    """绘制实体矩形+名称（宽度自适应文字）。"""
    ew = get_ent_w(draw, name)
    x0, y0 = cx - ew / 2, cy - ENT_H / 2
    x1, y1 = cx + ew / 2, cy + ENT_H / 2
    draw.rectangle((x0, y0, x1, y1), outline="black", width=2, fill="white")
    centered_text(draw, cx, cy, name, FONT_ENTITY)


def draw_attribute(draw, cx, cy, text):
    """绘制属性椭圆+名称。"""
    draw.ellipse(
        (cx - ATT_RX, cy - ATT_RY, cx + ATT_RX, cy + ATT_RY),
        outline="black", width=2, fill="white"
    )
    centered_text(draw, cx, cy, text, FONT_ATTR)


def connect(draw, ex, ey, ax, ay, ent_name):
    """实体边缘 → 属性椭圆边缘连线（直线）。"""
    dx, dy = ax - ex, ay - ey
    length = math.hypot(dx, dy)
    if length == 0:
        return

    ew = get_ent_w(draw, ent_name)
    hw, hh = ew / 2, ENT_H / 2
    t_rect = min(
        abs(hw / dx) if dx else float('inf'),
        abs(hh / dy) if dy else float('inf'),
    )
    sx = ex + dx * t_rect
    sy = ey + dy * t_rect

    t_ell = math.atan2(ATT_RX * (-dy), ATT_RY * (-dx))
    tx = ax + ATT_RX * math.cos(t_ell)
    ty = ay + ATT_RY * math.sin(t_ell)

    draw.line((sx, sy, tx, ty), fill="black", width=2)


def place_attrs(cx, cy, attrs):
    """
    为实体中心 (cx, cy) 计算属性位置列表。
    策略：前半部分排上方，后半部分排下方，左右分列。
    返回 [(ax, ay), ...]
    """
    n = len(attrs)
    top_count    = math.ceil(n / 2)
    bottom_count = n - top_count

    positions = []

    # ── 上方 ──────────────────────────────────────────────────────────────────
    # 在实体正上方均匀分布若干椭圆
    top_y = cy - ENT_H / 2 - ATT_RY * 2 - 18   # 椭圆中心 y
    if top_count == 1:
        positions.append((cx, top_y))
    else:
        total_w  = (top_count - 1) * (ATT_RX * 2 + 18)
        x_start  = cx - total_w / 2
        for i in range(top_count):
            positions.append((x_start + i * (ATT_RX * 2 + 18), top_y))

    # ── 下方 ──────────────────────────────────────────────────────────────────
    bot_y = cy + ENT_H / 2 + ATT_RY * 2 + 18
    if bottom_count == 1:
        positions.append((cx, bot_y))
    elif bottom_count > 1:
        total_w  = (bottom_count - 1) * (ATT_RX * 2 + 18)
        x_start  = cx - total_w / 2
        for i in range(bottom_count):
            positions.append((x_start + i * (ATT_RX * 2 + 18), bot_y))

    return positions


# ── 主绘图 ────────────────────────────────────────────────────────────────────
def main():
    img  = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    for idx, (name, attrs) in enumerate(ENTITIES):
        row = idx // COLS
        col = idx  % COLS
        # 实体中心
        cx = MARGIN_X + col * CELL_W + CELL_W / 2
        cy = MARGIN_Y + row * CELL_H + CELL_H / 2

        attr_positions = place_attrs(cx, cy, attrs)

        # 先画连线（在椭圆/矩形之下）
        for (ax, ay) in attr_positions:
            connect(draw, cx, cy, ax, ay, name)

        # 画实体
        draw_entity(draw, cx, cy, name)

        # 画属性椭圆
        for (ax, ay), attr in zip(attr_positions, attrs):
            draw_attribute(draw, ax, ay, attr)

    img.save(OUT_PATH, dpi=(150, 150))
    print(f"生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
