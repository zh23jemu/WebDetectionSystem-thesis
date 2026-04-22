from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


OUT_DIR = Path(r"C:\Coding\WebDetectionSystem-thesis\paper_diagrams\usecase_generated")
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


FONT_24 = load_font(24)
FONT_28 = load_font(28)
FONT_32 = load_font(32)


def centered_text(draw, xy, text, font, fill="black"):
    bbox = draw.textbbox((0, 0), text, font=font)
    x = xy[0] - (bbox[2] - bbox[0]) / 2
    y = xy[1] - (bbox[3] - bbox[1]) / 2
    draw.text((x, y), text, font=font, fill=fill)


def draw_actor(draw, x, y, label):
    draw.ellipse((x - 18, y, x + 18, y + 36), outline="black", width=3)
    draw.line((x, y + 36, x, y + 96), fill="black", width=3)
    draw.line((x - 26, y + 58, x + 26, y + 58), fill="black", width=3)
    draw.line((x, y + 96, x - 24, y + 124), fill="black", width=3)
    draw.line((x, y + 96, x + 24, y + 124), fill="black", width=3)
    centered_text(draw, (x, y + 155), label, FONT_28)


def draw_usecase(draw, box, text):
    draw.ellipse(box, outline="black", width=3, fill="white")
    centered_text(draw, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), text, FONT_28)


def draw_link(draw, start, end):
    draw.line((start, end), fill="black", width=2)


def user_usecase():
    img = Image.new("RGB", (1500, 900), "white")
    draw = ImageDraw.Draw(img)

    draw_actor(draw, 210, 320, "用户")

    cases = [
        (980, 120, 1260, 190, "注册登录"),
        (980, 220, 1260, 290, "开始检测"),
        (980, 320, 1260, 390, "缺陷历史"),
        (980, 420, 1260, 490, "日志报表"),
        (980, 520, 1260, 590, "销售管理"),
        (980, 620, 1260, 690, "帮助中心"),
        (980, 720, 1260, 790, "个人中心"),
    ]

    anchor = (240, 375)
    for x1, y1, x2, y2, text in cases:
        draw_usecase(draw, (x1, y1, x2, y2), text)
        draw_link(draw, anchor, (x1, (y1 + y2) // 2))

    out = OUT_DIR / "图3-1_普通用户用例图.png"
    img.save(out)
    return out


def admin_usecase():
    img = Image.new("RGB", (1500, 820), "white")
    draw = ImageDraw.Draw(img)

    draw_actor(draw, 210, 270, "管理员")

    cases = [
        (980, 120, 1260, 190, "管理员登录"),
        (980, 250, 1260, 320, "用户管理"),
        (980, 380, 1260, 450, "系统设置"),
        (980, 510, 1260, 580, "模型切换"),
        (980, 640, 1260, 710, "日志查看"),
    ]

    anchor = (240, 325)
    for x1, y1, x2, y2, text in cases:
        draw_usecase(draw, (x1, y1, x2, y2), text)
        draw_link(draw, anchor, (x1, (y1 + y2) // 2))

    out = OUT_DIR / "图3-2_管理员用例图.png"
    img.save(out)
    return out


def main():
    p1 = user_usecase()
    p2 = admin_usecase()
    print(p1)
    print(p2)


if __name__ == "__main__":
    main()
