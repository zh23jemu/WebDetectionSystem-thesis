#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成标准功能结构图（图4-1 系统总体结构图）。
横向树形布局：根节点居中，子模块向下扩展。
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\图4-1_系统总体结构图.png")

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

FONT_ROOT   = load_font(28)
FONT_L2     = load_font(24)
FONT_L3     = load_font(20)

# ── 树结构定义 ─────────────────────────────────────────────────────────────────
# (父节点名, [(子节点, [孙节点, ...])])
TREE = {
    "root": "木质板材缺陷检测系统",
    "modules": [
        ("用户管理模块", ["用户注册", "用户登录", "个人信息管理"]),
        ("缺陷检测模块", ["图片上传", "智能检测", "结果展示"]),
        ("历史记录管理", ["检测历史查询", "结果图查看"]),
        ("日志报表模块", ["工作日志查看", "数据报表统计"]),
        ("销售库存管理", ["销售记录管理", "库存数据管理"]),
        ("系统管理模块", ["用户权限管理", "模型切换", "系统设置"]),
    ]
}

# ── 节点尺寸 ──────────────────────────────────────────────────────────────────
ROOT_W, ROOT_H   = 280, 50
MOD_W,  MOD_H    = 160, 44
SUB_W,  SUB_H    = 148, 38

# 间距
MOD_H_GAP  = 28    # 模块列之间水平间距
SUB_V_GAP  = 14    # 子节点之间垂直间距
ROW1_H     = 80    # 根节点底部 → 模块节点顶部
ROW2_H     = 70    # 模块节点底部 → 子节点顶部

MARGIN_X = 36
MARGIN_Y = 36

def calc_layout():
    """
    计算每个节点的 (cx, cy) 坐标。
    自下向上：先算叶子高度，再算模块 y，最后算根节点 y。
    x 方向：模块列均匀分布。
    """
    modules = TREE["modules"]
    n_mods  = len(modules)

    # 每个模块列宽 = max(SUB_W, MOD_W)
    col_w = max(SUB_W, MOD_W)

    # 每个模块列的子节点高度
    col_heights = []
    for _, subs in modules:
        h = len(subs) * SUB_H + (len(subs) - 1) * SUB_V_GAP
        col_heights.append(h)

    max_col_h = max(col_heights)

    # 总画布尺寸
    total_w = n_mods * col_w + (n_mods - 1) * MOD_H_GAP + 2 * MARGIN_X
    total_h = (MARGIN_Y
               + ROOT_H
               + ROW1_H
               + MOD_H
               + ROW2_H
               + max_col_h
               + MARGIN_Y)

    # 根节点中心
    root_cx = total_w / 2
    root_cy = MARGIN_Y + ROOT_H / 2

    # 模块中心 x
    mod_cx_list = []
    start_x = MARGIN_X + col_w / 2
    for i in range(n_mods):
        mod_cx_list.append(start_x + i * (col_w + MOD_H_GAP))

    mod_cy = root_cy + ROOT_H / 2 + ROW1_H + MOD_H / 2

    # 子节点
    sub_positions = []   # [(cx, cy), ...]
    mod_groups    = []   # [[(cx, cy), ...], ...]  按模块分组
    for i, (mod_name, subs) in enumerate(modules):
        mcx = mod_cx_list[i]
        group = []
        n_subs = len(subs)
        total_sub_h = n_subs * SUB_H + (n_subs - 1) * SUB_V_GAP
        sub_top_cy = mod_cy + MOD_H / 2 + ROW2_H + SUB_H / 2
        for j in range(n_subs):
            scy = sub_top_cy + j * (SUB_H + SUB_V_GAP)
            group.append((mcx, scy))
        mod_groups.append(group)

    return total_w, total_h, root_cx, root_cy, mod_cx_list, mod_cy, mod_groups


def draw_rect(draw, cx, cy, w, h, font, text, fill_color="white"):
    x0, y0 = cx - w / 2, cy - h / 2
    x1, y1 = cx + w / 2, cy + h / 2
    draw.rectangle((x0, y0, x1, y1), outline="black", width=2, fill=fill_color)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text((cx - tw / 2, cy - th / 2), text, font=font, fill="black")


def main():
    W, H, root_cx, root_cy, mod_cx_list, mod_cy, mod_groups = calc_layout()
    W, H = int(W), int(H)

    img  = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    modules = TREE["modules"]

    # ── 连线 ─────────────────────────────────────────────────────────────────
    # 根 → 各模块（先画线，后画框）

    # 从根节点底部引一根竖线到所有模块列的中间高度，再横向到各模块
    conn_y = root_cy + ROOT_H / 2 + ROW1_H / 2    # 水平汇聚线 y

    draw.line((root_cx, root_cy + ROOT_H / 2, root_cx, conn_y), fill="black", width=2)
    draw.line((mod_cx_list[0], conn_y, mod_cx_list[-1], conn_y), fill="black", width=2)
    for mcx in mod_cx_list:
        draw.line((mcx, conn_y, mcx, mod_cy - MOD_H / 2), fill="black", width=2)

    # 模块 → 子节点
    for i, ((mod_name, subs), group) in enumerate(zip(modules, mod_groups)):
        mcx = mod_cx_list[i]
        sub_conn_y = mod_cy + MOD_H / 2 + ROW2_H / 2    # 子节点汇聚线 y

        draw.line((mcx, mod_cy + MOD_H / 2, mcx, sub_conn_y), fill="black", width=2)
        if len(group) > 1:
            draw.line((group[0][0], sub_conn_y, group[-1][0], sub_conn_y), fill="black", width=2)

        for (scx, scy) in group:
            draw.line((scx, sub_conn_y, scx, scy - SUB_H / 2), fill="black", width=2)

    # ── 节点框 ───────────────────────────────────────────────────────────────
    # 根节点（浅灰背景突出）
    draw_rect(draw, int(root_cx), int(root_cy), ROOT_W, ROOT_H,
              FONT_ROOT, TREE["root"], fill_color="#f0f0f0")

    # 模块节点
    for i, (mod_name, _) in enumerate(modules):
        draw_rect(draw, int(mod_cx_list[i]), int(mod_cy), MOD_W, MOD_H,
                  FONT_L2, mod_name)

    # 子节点
    for i, ((_, subs), group) in enumerate(zip(modules, mod_groups)):
        for (scx, scy), sub_name in zip(group, subs):
            draw_rect(draw, int(scx), int(scy), SUB_W, SUB_H,
                      FONT_L3, sub_name)

    img.save(OUT_PATH, dpi=(150, 150))
    print(f"生成: {OUT_PATH}")


if __name__ == "__main__":
    main()
