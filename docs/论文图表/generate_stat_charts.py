from __future__ import annotations

import math
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"C:\Users\m1889\Desktop\毕设")
OUT_DIR = ROOT / "论文图表优化" / "统计图表"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        r"C:\Windows\Fonts\msyhbd.ttc" if bold else r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


F_TITLE = font(54, True)
F_SUBTITLE = font(31, True)
F_AXIS = font(26)
F_LABEL = font(29)
F_SMALL = font(23)
F_BOLD = font(30, True)

NAVY = "#17456f"
BLUE = "#2f75c8"
CYAN = "#4aa3df"
GREEN = "#55a868"
ORANGE = "#f28e2b"
RED = "#d55e5e"
PURPLE = "#8172b2"
GRAY = "#4f4f4f"
LIGHT_BLUE = "#eaf2fb"
GRID = "#d8dee8"
BLACK = "#111111"


def canvas(w: int = 2400, h: int = 1500) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    img = Image.new("RGB", (w, h), "white")
    return img, ImageDraw.Draw(img)


def text_size(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=fnt)
    return box[2] - box[0], box[3] - box[1]


def center_text(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, fnt: ImageFont.ImageFont, fill=BLACK) -> None:
    x, y = xy
    w, h = text_size(draw, text, fnt)
    draw.text((x - w / 2, y - h / 2), text, fill=fill, font=fnt)


def title(draw: ImageDraw.ImageDraw, text: str, w: int) -> None:
    center_text(draw, (w // 2, 78), text, F_TITLE, NAVY)


def draw_panel(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], label: str) -> None:
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=24, fill="#fbfdff", outline="#c8d3e1", width=3)
    draw.text((x1 + 28, y1 + 20), label, fill=NAVY, font=F_SUBTITLE)


def draw_axes(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    y_max: float,
    y_ticks: Iterable[float],
    y_label: str = "",
) -> None:
    x1, y1, x2, y2 = box
    draw.line((x1, y2, x2, y2), fill=BLACK, width=3)
    draw.line((x1, y1, x1, y2), fill=BLACK, width=3)
    for tick in y_ticks:
        yy = y2 - (tick / y_max) * (y2 - y1)
        draw.line((x1, yy, x2, yy), fill=GRID, width=2)
        draw.text((x1 - 72, yy - 16), f"{tick:g}", fill=GRAY, font=F_SMALL)
    if y_label:
        draw.text((x1, y1 - 42), y_label, fill=GRAY, font=F_SMALL)


def line_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    x_labels: list[str],
    series: list[tuple[str, list[float], str, float, list[float]]],
) -> None:
    x1, y1, x2, y2 = box
    panel_w = (x2 - x1 - 80) // len(series)
    for idx, (name, values, color, y_max, ticks) in enumerate(series):
        px1 = x1 + idx * panel_w + 72
        px2 = px1 + panel_w - 78
        py1 = y1 + 74
        py2 = y2 - 92
        draw.text((px1, y1 + 16), name, fill=color, font=F_BOLD)
        draw_axes(draw, (px1, py1, px2, py2), y_max, ticks)
        plot_x1 = px1 + 62
        plot_x2 = px2 - 62
        xs = [plot_x1 + i * (plot_x2 - plot_x1) / (len(values) - 1) for i in range(len(values))]
        pts = [(xs[i], py2 - values[i] / y_max * (py2 - py1)) for i in range(len(values))]
        draw.line(pts, fill=color, width=6, joint="curve")
        for (x, y), val in zip(pts, values):
            draw.ellipse((x - 10, y - 10, x + 10, y + 10), fill=color, outline="white", width=3)
            center_text(draw, (int(x), int(y - 34)), f"{val:g}", F_SMALL, color)
        for x, lab in zip(xs, x_labels):
            center_text(draw, (int(x), py2 + 36), lab, F_SMALL, BLACK)


def chart_adaptive_params() -> None:
    img, draw = canvas(2400, 1500)
    title(draw, "图5-3 干扰等级与自适应传输参数变化", 2400)
    x_labels = ["无干扰", "轻度", "中度", "严重"]
    series = [
        ("JPEG质量 Q", [85, 55, 35, 15], BLUE, 100, [0, 25, 50, 75, 100]),
        ("FEC冗余倍数", [1, 2, 2, 3], ORANGE, 3.5, [0, 1, 2, 3]),
        ("帧率 FPS", [25, 15, 8, 5], GREEN, 30, [0, 10, 20, 30]),
    ]
    draw_panel(draw, (92, 142, 2308, 1088), "传输参数随干扰严重程度动态收缩")
    line_panel(draw, (120, 226, 2280, 1048), x_labels, series)

    blocks = [
        ("无干扰", "Clean 或置信度 < 30%", "高质量传输", BLUE),
        ("轻度", "NAM / NFM / SIN", "降质保传", GREEN),
        ("中度", "LFM / MTJ", "降帧保传", ORANGE),
        ("严重", "STJ", "最低质保", RED),
    ]
    x = 120
    for level, trigger, strategy, color in blocks:
        draw.rounded_rectangle((x, 1132, x + 510, 1352), radius=22, fill="#f7f9fc", outline=color, width=4)
        center_text(draw, (x + 255, 1185), level, F_BOLD, color)
        center_text(draw, (x + 255, 1246), trigger, F_AXIS, BLACK)
        center_text(draw, (x + 255, 1305), strategy, F_AXIS, GRAY)
        x += 550
    img.save(OUT_DIR / "图5-3_干扰等级与自适应传输参数变化.png", quality=95)


def grouped_bars(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    labels: list[str],
    values_a: list[float],
    values_b: list[float],
    name_a: str,
    name_b: str,
    y_max: float,
    unit: str,
) -> None:
    x1, y1, x2, y2 = box
    draw_axes(draw, box, y_max, [0, y_max / 4, y_max / 2, y_max * 3 / 4, y_max], unit)
    group_w = (x2 - x1) / len(labels)
    bar_w = group_w * 0.28
    for i, lab in enumerate(labels):
        cx = x1 + group_w * (i + 0.5)
        for j, (val, color) in enumerate([(values_a[i], BLUE), (values_b[i], ORANGE)]):
            bx1 = cx - bar_w - 6 if j == 0 else cx + 6
            bx2 = bx1 + bar_w
            by = y2 - val / y_max * (y2 - y1)
            draw.rounded_rectangle((bx1, by, bx2, y2), radius=8, fill=color)
            center_text(draw, (int((bx1 + bx2) / 2), int(by - 26)), f"{int(val):,}", F_SMALL, color)
        center_text(draw, (int(cx), y2 + 40), lab, F_SMALL, BLACK)
    draw.rounded_rectangle((x2 - 430, y1 - 70, x2 - 50, y1 - 14), radius=16, fill="white", outline="#cbd5e1", width=2)
    draw.rectangle((x2 - 400, y1 - 54, x2 - 372, y1 - 26), fill=BLUE)
    draw.text((x2 - 360, y1 - 58), name_a, fill=BLACK, font=F_SMALL)
    draw.rectangle((x2 - 210, y1 - 54, x2 - 182, y1 - 26), fill=ORANGE)
    draw.text((x2 - 170, y1 - 58), name_b, fill=BLACK, font=F_SMALL)


def chart_dataset_distribution() -> None:
    img, draw = canvas(2400, 1450)
    title(draw, "图4-5 训练集与测试集样本分布", 2400)
    draw_panel(draw, (100, 150, 2300, 1280), "七类干扰识别数据集构成")
    labels = ["LFM", "MTJ", "NAM", "NFM", "STJ", "SIN", "Clean"]
    train = [1515, 1529, 1488, 1533, 1528, 1543, 1504]
    test = [385, 371, 412, 367, 372, 357, 396]
    grouped_bars(draw, (230, 330, 2220, 1120), labels, train, test, "训练样本", "测试样本", 1800, "样本数")
    center_text(draw, (1200, 1220), "合计：训练样本 10,640，测试样本 2,660；Clean类作为无干扰基准样本", F_AXIS, GRAY)
    img.save(OUT_DIR / "图4-5_训练集与测试集样本分布.png", quality=95)


def chart_jsr_accuracy() -> None:
    img, draw = canvas(2400, 1500)
    title(draw, "图4-6 JSR条件下的干扰识别准确率变化", 2400)
    draw_panel(draw, (100, 150, 2300, 1320), "总体准确率与各类别低JSR鲁棒性")
    jsr = [-6, -4, -2, 0, 4, 8, 12, 16]
    overall = [65.3, 96.0, 99.4, 99.9, 100.0, 100.0, 100.0, 99.9]
    x1, y1, x2, y2 = (250, 320, 1440, 1060)
    draw_axes(draw, (x1, y1, x2, y2), 100, [0, 25, 50, 75, 100], "准确率（%）")
    xs = [x1 + i * (x2 - x1) / (len(jsr) - 1) for i in range(len(jsr))]
    pts = [(xs[i], y2 - overall[i] / 100 * (y2 - y1)) for i in range(len(jsr))]
    draw.line(pts, fill=BLUE, width=7, joint="curve")
    for i, ((x, y), val) in enumerate(zip(pts, overall)):
        draw.ellipse((x - 11, y - 11, x + 11, y + 11), fill=BLUE, outline="white", width=3)
        center_text(draw, (int(x), int(y - 38)), f"{val:g}", F_SMALL, BLUE)
        center_text(draw, (int(x), y2 + 42), str(jsr[i]), F_SMALL, BLACK)
    center_text(draw, ((x1 + x2) // 2, y2 + 95), "JSR（dB）", F_AXIS, BLACK)

    # Compact heatmap for low JSR rows.
    classes = ["LFM", "MTJ", "NAM", "NFM", "STJ", "SIN", "Clean"]
    low_rows = [
        ("-6dB", [66, 90, 26, 71, 55, 49, 100]),
        ("-4dB", [100, 97, 87, 99, 89, 100, 100]),
        ("-2dB", [100, 98, 99, 100, 99, 100, 100]),
    ]
    hx1, hy1 = 1540, 360
    cw, ch = 104, 92
    draw.text((hx1, hy1 - 74), "低JSR各类别准确率（%）", fill=NAVY, font=F_BOLD)
    for j, c in enumerate(classes):
        center_text(draw, (hx1 + 120 + j * cw + cw // 2, hy1 - 24), c, F_SMALL, BLACK)
    for i, (row, vals) in enumerate(low_rows):
        center_text(draw, (hx1 + 54, hy1 + i * ch + ch // 2), row, F_SMALL, BLACK)
        for j, v in enumerate(vals):
            intensity = int(255 - (v / 100) * 90)
            fill = (intensity, 235, 255) if v >= 90 else (255, 222, 205)
            x = hx1 + 120 + j * cw
            y = hy1 + i * ch
            draw.rectangle((x, y, x + cw, y + ch), fill=fill, outline="#667085", width=2)
            center_text(draw, (x + cw // 2, y + ch // 2), f"{v:g}", F_SMALL, BLACK)
    center_text(draw, (1840, 1168), "结论：JSR≥-2dB时总体准确率保持在99%以上；-6dB时主要短板为NAM、STJ和SIN。", F_AXIS, GRAY)
    img.save(OUT_DIR / "图4-6_JSR条件下的干扰识别准确率变化.png", quality=95)


def chart_model_params() -> None:
    img, draw = canvas(2200, 1400)
    title(draw, "图6-4 CCNN模型参数量构成", 2200)
    modules = [
        ("CNN特征提取", 38016, BLUE),
        ("残差网络（含SE）", 19714, GREEN),
        ("TCC注意力", 35984, ORANGE),
        ("分类器", 3207, PURPLE),
    ]
    total = sum(v for _, v, _ in modules)
    cx, cy, r = 610, 720, 330
    start = -90
    for name, val, color in modules:
        angle = 360 * val / total
        draw.pieslice((cx - r, cy - r, cx + r, cy + r), start, start + angle, fill=color, outline="white", width=4)
        start += angle
    draw.ellipse((cx - 175, cy - 175, cx + 175, cy + 175), fill="white", outline="white")
    center_text(draw, (cx, cy - 30), "总参数量", F_BOLD, NAVY)
    center_text(draw, (cx, cy + 28), f"{total:,}", F_TITLE, NAVY)
    x0, y0 = 1120, 355
    for i, (name, val, color) in enumerate(modules):
        y = y0 + i * 170
        pct = val / total * 100
        draw.rounded_rectangle((x0, y, x0 + 780, y + 118), radius=20, fill="#f8fafc", outline="#cbd5e1", width=2)
        draw.rectangle((x0 + 28, y + 34, x0 + 72, y + 78), fill=color)
        draw.text((x0 + 95, y + 22), name, fill=BLACK, font=F_BOLD)
        draw.text((x0 + 95, y + 66), f"{val:,} 参数，占比 {pct:.1f}%", fill=GRAY, font=F_AXIS)
    img.save(OUT_DIR / "图6-4_CCNN模型参数量构成.png", quality=95)


def chart_method_comparison() -> None:
    img, draw = canvas(2400, 1450)
    title(draw, "图6-5 不同识别方法性能对比", 2400)
    labels = ["PSD+SVM", "纯CNN", "CNN+ResNet", "CCNN\n本文"]
    acc = [92.0, 98.7, 99.5, 99.91]
    jsr = [75.0, 94.5, 97.2, 99.4]
    grouped_bars(draw, (280, 300, 2200, 1120), labels, acc, jsr, "测试准确率", "JSR=-2dB准确率", 105, "准确率（%）")
    center_text(draw, (1200, 1235), "CCNN在总体准确率和低JSR鲁棒性上均优于对比方法，体现注意力与残差结构的增益。", F_AXIS, GRAY)
    img.save(OUT_DIR / "图6-5_不同识别方法性能对比.png", quality=95)


def chart_demo_timeline() -> None:
    img, draw = canvas(2400, 1300)
    title(draw, "图6-6 自动演示测试干扰时序", 2400)
    stages = [
        (0, 5, "Clean", "#8fd694"),
        (5, 10, "LFM\n4dB", "#ffd166"),
        (10, 15, "MTJ\n0dB", "#f4a261"),
        (15, 20, "STJ\n10dB", "#e76f51"),
        (20, 25, "NAM\n2dB", "#9d7fe8"),
        (25, 30, "NFM\n6dB", "#6ec6ff"),
        (30, 35, "SIN\n8dB", "#b8c0ff"),
        (35, 40, "Clean\n恢复", "#8fd694"),
    ]
    x1, x2 = 210, 2200
    y = 560
    total = 40
    draw.line((x1, y, x2, y), fill=BLACK, width=5)
    for start, end, lab, color in stages:
        sx = x1 + (start / total) * (x2 - x1)
        ex = x1 + (end / total) * (x2 - x1)
        draw.rounded_rectangle((sx, y - 150, ex, y + 150), radius=18, fill=color, outline="white", width=4)
        center_text(draw, (int((sx + ex) / 2), y), lab, F_BOLD, BLACK)
        center_text(draw, (int(sx), y + 210), f"{start}s", F_SMALL, GRAY)
    center_text(draw, (x2, y + 210), "40s", F_SMALL, GRAY)
    draw.rounded_rectangle((260, 930, 2140, 1088), radius=24, fill="#f8fafc", outline="#cbd5e1", width=3)
    center_text(draw, (1200, 985), "演示序列覆盖无干扰、扫频、多音、单音、窄带调幅、窄带调频和正弦波干扰，用于验证识别与策略切换闭环。", F_AXIS, GRAY)
    img.save(OUT_DIR / "图6-6_自动演示测试干扰时序.png", quality=95)


def chart_usrp_cards() -> None:
    img, draw = canvas(2400, 1400)
    title(draw, "图6-7 USRP实机模式测试结果摘要", 2400)
    cards = [
        ("120秒实机模式", [("发送帧数", "365"), ("接收帧数", "732"), ("干扰事件", "23"), ("策略切换", "23次")], BLUE),
        ("实时可视化长测", [("发送帧数", "63"), ("接收帧数", "83"), ("干扰事件", "2"), ("策略切换", "2次")], GREEN),
    ]
    x = 170
    for title_text, items, color in cards:
        draw.rounded_rectangle((x, 250, x + 970, 1030), radius=28, fill="#fbfdff", outline=color, width=5)
        center_text(draw, (x + 485, 325), title_text, F_TITLE, color)
        for i, (k, v) in enumerate(items):
            yy = 445 + i * 130
            draw.rounded_rectangle((x + 80, yy, x + 890, yy + 90), radius=18, fill="#f8fafc", outline="#d0d7e2", width=2)
            draw.text((x + 130, yy + 22), k, fill=GRAY, font=F_AXIS)
            tw, _ = text_size(draw, v, F_BOLD)
            draw.text((x + 830 - tw, yy + 18), v, fill=color, font=F_BOLD)
        x += 1090
    center_text(draw, (1200, 1180), "实机测试表明平台可完成干扰检测、类型识别、参数切换和实时可视化闭环。", F_AXIS, GRAY)
    img.save(OUT_DIR / "图6-7_USRP实机模式测试结果摘要.png", quality=95)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    chart_adaptive_params()
    chart_dataset_distribution()
    chart_jsr_accuracy()
    chart_model_params()
    chart_method_comparison()
    chart_demo_timeline()
    chart_usrp_cards()
    print(f"saved={OUT_DIR}")


if __name__ == "__main__":
    main()
