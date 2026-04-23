from math import cos, pi, sin
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


OUT_DIR = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\reference_style_generated")
OUT_DIR.mkdir(parents=True, exist_ok=True)

FONT_PATHS = [
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simfang.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
]


def load_font(size: int):
    for path in FONT_PATHS:
        p = Path(path)
        if p.exists():
            return ImageFont.truetype(str(p), size)
    return ImageFont.load_default()


FONT_20 = load_font(20)
FONT_24 = load_font(24)
FONT_28 = load_font(28)
FONT_32 = load_font(32)
FONT_36 = load_font(36)
FONT_40 = load_font(40)


def canvas(size=(1600, 950)):
    return Image.new("RGB", size, "white"), ImageDraw.Draw(Image.new("RGB", size, "white"))


def make_canvas(size=(1600, 950)):
    img = Image.new("RGB", size, "white")
    return img, ImageDraw.Draw(img)


def centered_text(draw, xy, text, font, fill="black"):
    bbox = draw.textbbox((0, 0), text, font=font)
    x = xy[0] - (bbox[2] - bbox[0]) / 2
    y = xy[1] - (bbox[3] - bbox[1]) / 2
    draw.text((x, y), text, font=font, fill=fill)


def draw_box(draw, box, text, font=FONT_28, fill="white"):
    draw.rectangle(box, outline="black", width=3, fill=fill)
    centered_text(draw, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), text, font)


def draw_ellipse(draw, box, text, font=FONT_28):
    draw.ellipse(box, outline="black", width=3, fill="white")
    centered_text(draw, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), text, font)


def draw_diamond(draw, center, size, text, font=FONT_24):
    cx, cy = center
    w, h = size
    pts = [(cx, cy - h // 2), (cx + w // 2, cy), (cx, cy + h // 2), (cx - w // 2, cy)]
    draw.polygon(pts, outline="black", width=3, fill="white")
    centered_text(draw, center, text, font)


def draw_arrow(draw, p1, p2, text=None, dashed=False, font=FONT_24):
    if dashed:
        draw_dashed_line(draw, p1, p2, dash=14, gap=8, width=3)
    else:
        draw.line([p1, p2], fill="black", width=3)

    angle = pi / 6
    length = 18
    import math
    theta = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
    p3 = (p2[0] - length * cos(theta - angle), p2[1] - length * sin(theta - angle))
    p4 = (p2[0] - length * cos(theta + angle), p2[1] - length * sin(theta + angle))
    draw.line([p2, p3], fill="black", width=3)
    draw.line([p2, p4], fill="black", width=3)
    if text:
        centered_text(draw, ((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2 - 20), text, font)


def draw_dashed_line(draw, p1, p2, dash=12, gap=8, width=3):
    import math
    x1, y1 = p1
    x2, y2 = p2
    dx = x2 - x1
    dy = y2 - y1
    dist = math.hypot(dx, dy)
    if dist == 0:
        return
    vx = dx / dist
    vy = dy / dist
    pos = 0
    while pos < dist:
        end = min(pos + dash, dist)
        sx = x1 + vx * pos
        sy = y1 + vy * pos
        ex = x1 + vx * end
        ey = y1 + vy * end
        draw.line([(sx, sy), (ex, ey)], fill="black", width=width)
        pos += dash + gap


def draw_actor(draw, x, y, label):
    draw.ellipse((x - 18, y, x + 18, y + 36), outline="black", width=3)
    draw.line((x, y + 36, x, y + 96), fill="black", width=3)
    draw.line((x - 26, y + 58, x + 26, y + 58), fill="black", width=3)
    draw.line((x, y + 96, x - 24, y + 124), fill="black", width=3)
    draw.line((x, y + 96, x + 24, y + 124), fill="black", width=3)
    centered_text(draw, (x, y - 18), label, FONT_28)


def draw_lifeline(draw, x, top, bottom, label, actor=False):
    if actor:
        draw_actor(draw, x, top, label)
        start_y = top + 138
    else:
        draw_box(draw, (x - 92, top, x + 92, top + 68), label, FONT_28)
        start_y = top + 68
    draw_dashed_line(draw, (x, start_y), (x, bottom), dash=12, gap=10, width=2)
    draw.rectangle((x - 14, start_y + 20, x + 14, bottom - 40), outline="black", width=3, fill="white")


def draw_alt_box(draw, box, sections):
    x1, y1, x2, y2 = box
    draw.rectangle(box, outline="black", width=2)
    draw.text((x1 + 12, y1 + 10), "alt", font=FONT_24, fill="black")
    section_height = (y2 - y1 - 50) / len(sections)
    cy = y1 + 40
    for idx, title in enumerate(sections):
        if idx > 0:
            draw_dashed_line(draw, (x1, cy), (x2, cy), dash=10, gap=6, width=2)
        centered_text(draw, ((x1 + x2) / 2, cy + 24), title, FONT_24)
        cy += section_height


def save(img, name):
    img.save(OUT_DIR / name)


def fig4_1():
    img, draw = make_canvas((1600, 980))
    draw_box(draw, (250, 70, 470, 150), "普通用户")
    draw_box(draw, (1130, 70, 1350, 150), "管理员")
    draw_box(draw, (540, 165, 1060, 275), "Web前端界面", FONT_40)

    y1 = 360
    boxes = [
        ((70, y1, 300, y1 + 90), "登录与权限模块"),
        ((330, y1, 560, y1 + 90), "缺陷检测模块"),
        ((590, y1, 820, y1 + 90), "历史记录模块"),
        ((850, y1, 1080, y1 + 90), "日志报表模块"),
        ((1110, y1, 1340, y1 + 90), "销售与库存模块"),
        ((1370, y1, 1580, y1 + 90), "帮助中心与系统设置"),
    ]
    for box, text in boxes:
        draw_box(draw, box, text, FONT_28)

    bottom = [
        ((250, 640, 500, 730), "YOLO检测模型"),
        ((670, 640, 930, 730), "SQLite数据库"),
        ((1090, 640, 1370, 730), "静态资源文件"),
    ]
    for box, text in bottom:
        draw_box(draw, box, text, FONT_32)

    for x in [360, 1240]:
        draw.line((x, 150, 800, 170), fill="black", width=3)
    draw.line((800, 270, 800, 330), fill="black", width=3)
    for box, _ in boxes:
        cx = (box[0] + box[2]) // 2
        draw.line((800, 330, cx, box[1]), fill="black", width=2)

    draw.line((445, 450, 375, 640), fill="black", width=2)
    draw.line((705, 450, 800, 640), fill="black", width=2)
    draw.line((965, 450, 800, 640), fill="black", width=2)
    draw.line((1225, 450, 800, 640), fill="black", width=2)
    draw.line((445, 450, 1230, 640), fill="black", width=2)
    draw.line((1475, 450, 1230, 640), fill="black", width=2)

    save(img, "图4-1_系统总体结构图_参考风格.png")


def fig4_2():
    img, draw = make_canvas((1600, 980))
    xs = [90, 640, 1040, 1480]
    labels = ["用户", "登录页面", "Flask后端", "数据库"]
    draw_lifeline(draw, xs[0], 70, 860, labels[0], actor=True)
    for x, label in zip(xs[1:], labels[1:]):
        draw_lifeline(draw, x, 70, 860, label, actor=False)

    draw_arrow(draw, (xs[0] + 14, 250), (xs[1] - 14, 250), "点击登录入口", font=FONT_28)
    draw_arrow(draw, (xs[1] - 16, 340), (xs[0] + 16, 340), "显示登录界面", font=FONT_28)
    draw_arrow(draw, (xs[0] + 14, 430), (xs[1] - 14, 430), "输入账号和密码", font=FONT_28)
    draw_arrow(draw, (xs[1] + 14, 530), (xs[2] - 14, 530), "提交登录请求", font=FONT_28)
    draw_arrow(draw, (xs[2] + 14, 620), (xs[3] - 14, 620), "查询用户信息", font=FONT_28)
    draw_arrow(draw, (xs[3] - 14, 700), (xs[2] + 14, 700), "返回用户记录", dashed=True, font=FONT_28)

    draw_alt_box(draw, (120, 730, 980, 930), ["[验证成功]", "[验证失败]"])
    draw_arrow(draw, (xs[2] - 14, 790), (xs[1] + 14, 790), "返回成功结果", dashed=True, font=FONT_28)
    draw_arrow(draw, (xs[1] - 14, 840), (xs[0] + 14, 840), "进入系统主界面", font=FONT_28)
    draw_arrow(draw, (xs[2] - 14, 900), (xs[1] + 14, 900), "返回失败结果", dashed=True, font=FONT_28)
    draw_arrow(draw, (xs[1] - 14, 920), (xs[0] + 14, 920), "显示错误提示", font=FONT_28)

    save(img, "图4-2_登录顺序图_参考风格.png")


def fig4_3():
    img, draw = make_canvas((1760, 980))
    xs = [80, 430, 760, 1090, 1410, 1670]
    labels = ["用户", "检测页面", "Flask后端", "YOLO模型", "数据库", "库存表"]
    draw_lifeline(draw, xs[0], 70, 860, labels[0], actor=True)
    for x, label in zip(xs[1:], labels[1:]):
        draw_lifeline(draw, x, 70, 860, label)

    ys = [240, 320, 400, 480, 560, 640, 720, 800]
    draw_arrow(draw, (xs[0] + 14, ys[0]), (xs[1] - 14, ys[0]), "选择图片或拍照", font=FONT_28)
    draw_arrow(draw, (xs[1] + 14, ys[1]), (xs[2] - 14, ys[1]), "上传检测图片", font=FONT_28)
    draw_arrow(draw, (xs[2] + 14, ys[2]), (xs[3] - 14, ys[2]), "调用模型检测", font=FONT_28)
    draw_arrow(draw, (xs[3] - 14, ys[3]), (xs[2] + 14, ys[3]), "返回类别与置信度", dashed=True, font=FONT_28)
    centered_text(draw, (760, 545), "缺陷映射与等级判定", FONT_28)
    draw_arrow(draw, (xs[2] + 14, ys[4]), (xs[4] - 14, ys[4]), "写入检测历史", font=FONT_28)
    draw_arrow(draw, (xs[2] + 14, ys[5]), (xs[5] - 14, ys[5]), "更新库存数量", font=FONT_28)
    draw_arrow(draw, (xs[2] - 14, ys[6]), (xs[1] + 14, ys[6]), "返回结果图和结果数据", dashed=True, font=FONT_28)
    draw_arrow(draw, (xs[1] - 14, ys[7]), (xs[0] + 14, ys[7]), "显示检测结果", font=FONT_28)

    save(img, "图4-3_缺陷检测顺序图_参考风格.png")


def fig4_4():
    img, draw = make_canvas((1600, 980))
    xs = [100, 520, 980, 1450]
    labels = ["用户", "历史/报表页面", "Flask后端", "数据库"]
    draw_lifeline(draw, xs[0], 70, 860, labels[0], actor=True)
    for x, label in zip(xs[1:], labels[1:]):
        draw_lifeline(draw, x, 70, 860, label)

    draw_arrow(draw, (xs[0] + 14, 250), (xs[1] - 14, 250), "进入页面", font=FONT_28)
    draw_arrow(draw, (xs[1] + 14, 340), (xs[2] - 14, 340), "发送查询请求", font=FONT_28)
    draw_arrow(draw, (xs[2] + 14, 430), (xs[3] - 14, 430), "查询历史与日志", font=FONT_28)
    draw_arrow(draw, (xs[3] - 14, 520), (xs[2] + 14, 520), "返回记录数据", dashed=True, font=FONT_28)
    centered_text(draw, (980, 605), "统计占比、趋势和工作时长", FONT_28)
    draw_arrow(draw, (xs[2] - 14, 700), (xs[1] + 14, 700), "返回列表和图表数据", dashed=True, font=FONT_28)
    draw_arrow(draw, (xs[1] - 14, 790), (xs[0] + 14, 790), "显示查询和统计结果", font=FONT_28)

    save(img, "图4-4_历史与报表顺序图_参考风格.png")


def fig4_5():
    img, draw = make_canvas((1680, 980))
    xs = [90, 470, 900, 1270, 1580]
    labels = ["用户", "销售页面", "Flask后端", "库存表", "数据库"]
    draw_lifeline(draw, xs[0], 70, 860, labels[0], actor=True)
    for x, label in zip(xs[1:], labels[1:]):
        draw_lifeline(draw, x, 70, 860, label)

    draw_arrow(draw, (xs[0] + 14, 240), (xs[1] - 14, 240), "填写订单信息", font=FONT_28)
    draw_arrow(draw, (xs[1] + 14, 330), (xs[2] - 14, 330), "提交销售订单", font=FONT_28)
    draw_arrow(draw, (xs[2] + 14, 420), (xs[3] - 14, 420), "查询库存数量", font=FONT_28)
    draw_arrow(draw, (xs[3] - 14, 510), (xs[2] + 14, 510), "返回库存结果", dashed=True, font=FONT_28)

    # 仅保留分支说明文字，不再绘制外层 alt 框和左上角标签。
    divider_y = 720
    centered_text(draw, (760, 645), "[库存充足]", FONT_28)
    centered_text(draw, (760, 760), "[库存不足]", FONT_28)
    draw_arrow(draw, (xs[2] + 14, 650), (xs[3] - 14, 650), "扣减库存", font=FONT_28)

    # Separate these two labels vertically to avoid overlap in the alt frame.
    draw.line((xs[2] + 14, 725, xs[4] - 14, 725), fill="black", width=3)
    centered_text(draw, ((xs[2] + xs[4]) / 2, 700), "写入销售记录", FONT_28)
    # arrow head for rightward solid line
    draw.line([(xs[4] - 14, 725), (xs[4] - 32, 715)], fill="black", width=3)
    draw.line([(xs[4] - 14, 725), (xs[4] - 32, 735)], fill="black", width=3)

    draw_dashed_line(draw, (xs[2] - 14, 790), (xs[1] + 14, 790), dash=14, gap=8, width=3)
    draw.line([(xs[1] + 14, 790), (xs[1] + 32, 780)], fill="black", width=3)
    draw.line([(xs[1] + 14, 790), (xs[1] + 32, 800)], fill="black", width=3)
    # Move this label further upward to avoid overlapping with the alt divider line.
    centered_text(draw, ((xs[2] + xs[1]) / 2 - 140, 700), "返回创建成功", FONT_28)

    draw_arrow(draw, (xs[1] - 14, 855), (xs[0] + 14, 855), "显示成功提示", font=FONT_28)

    draw_dashed_line(draw, (xs[2] - 14, 900), (xs[1] + 14, 900), dash=14, gap=8, width=3)
    draw.line([(xs[1] + 14, 900), (xs[1] + 32, 890)], fill="black", width=3)
    draw.line([(xs[1] + 14, 900), (xs[1] + 32, 910)], fill="black", width=3)
    centered_text(draw, ((xs[2] + xs[1]) / 2, 875), "返回库存不足提示", FONT_28)

    save(img, "图4-5_销售出库顺序图_参考风格.png")


def fig4_6():
    img, draw = make_canvas((1600, 980))
    center = (800, 490)
    draw_box(draw, (680, 420, 920, 560), "检测历史", FONT_36)

    ellipses = [
        ((610, 110, 990, 220), "检测时间"),
        ((1050, 210, 1450, 320), "结果图片"),
        ((1060, 420, 1480, 530), "缺陷类型"),
        ((1040, 660, 1420, 770), "判定等级"),
        ((560, 780, 1040, 890), "最高置信度"),
        ((130, 660, 510, 770), "用户编号"),
        ((120, 420, 510, 530), "原始文件名"),
        ((170, 210, 500, 320), "检测编号"),
    ]
    anchors = [
        (800, 430), (910, 450), (910, 490), (860, 550),
        (760, 550), (690, 510), (690, 470), (720, 430)
    ]
    ellipse_points = [
        (800, 220), (1040, 270), (1060, 470), (1020, 710),
        (800, 760), (520, 710), (510, 470), (500, 270)
    ]
    for (box, text), a, b in zip(ellipses, anchors, ellipse_points):
        draw_ellipse(draw, box, text, FONT_32)
        draw.line((a, b), fill="black", width=3)

    save(img, "图4-6_主要实体属性图_参考风格.png")


def fig4_7():
    img, draw = make_canvas((1700, 1100))
    # entities
    user = (150, 310, 340, 410)
    history = (730, 290, 960, 390)
    worklog = (710, 650, 940, 750)
    inventory = (1270, 280, 1470, 380)
    sales = (1250, 660, 1490, 760)
    for box, text in [
        (user, "用户"), (history, "检测历史"), (worklog, "工作日志"),
        (inventory, "库存"), (sales, "销售记录")
    ]:
        draw_box(draw, box, text, FONT_32)

    # relationships
    draw_diamond(draw, (520, 350), (160, 100), "产生", font=FONT_28)
    draw_diamond(draw, (500, 700), (160, 100), "记录", font=FONT_28)
    draw_diamond(draw, (1110, 330), (160, 100), "更新", font=FONT_28)
    draw_diamond(draw, (1090, 710), (160, 100), "出库", font=FONT_28)

    # lines and cardinality
    draw.line((340, 360, 440, 360), fill="black", width=3)
    draw.line((600, 350, 730, 340), fill="black", width=3)
    centered_text(draw, (390, 325), "1", FONT_28)
    centered_text(draw, (665, 320), "N", FONT_28)

    draw.line((315, 390, 430, 700), fill="black", width=3)
    draw.line((580, 700, 710, 700), fill="black", width=3)
    centered_text(draw, (355, 525), "1", FONT_28)
    centered_text(draw, (640, 665), "N", FONT_28)

    draw.line((960, 340, 1030, 330), fill="black", width=3)
    draw.line((1190, 330, 1270, 330), fill="black", width=3)
    centered_text(draw, (995, 300), "1", FONT_28)
    centered_text(draw, (1225, 300), "1", FONT_28)

    draw.line((940, 700, 1010, 710), fill="black", width=3)
    draw.line((1170, 710, 1250, 710), fill="black", width=3)
    centered_text(draw, (975, 665), "1", FONT_28)
    centered_text(draw, (1215, 675), "N", FONT_28)

    draw.line((960, 390, 1110, 660), fill="black", width=3)
    draw.line((1470, 380, 1110, 660), fill="black", width=3)
    centered_text(draw, (1075, 510), "1", FONT_28)
    centered_text(draw, (1355, 500), "N", FONT_28)

    save(img, "图4-7_系统E-R图_参考风格.png")


def main():
    fig4_1()
    fig4_2()
    fig4_3()
    fig4_4()
    fig4_5()
    fig4_6()
    fig4_7()
    print(OUT_DIR)


if __name__ == "__main__":
    main()
