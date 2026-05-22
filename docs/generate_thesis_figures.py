from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"C:\Users\m1889\Desktop\毕设")
OUT = ROOT / "提交材料_李智杰_封面修正版" / "论文插图"

W = 1800
H = 1050

COLORS = {
    "bg": "#FFFFFF",
    "text": "#111827",
    "muted": "#475569",
    "stroke": "#334155",
    "tx": "#DBEAFE",
    "rx": "#DCFCE7",
    "model": "#F3E8FF",
    "control": "#FEF3C7",
    "channel": "#FEE2E2",
    "neutral": "#F8FAFC",
    "data": "#2563EB",
    "feedback": "#D97706",
}


def font_path() -> str:
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\Deng.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    return ""


FONT_PATH = font_path()


def font(size: int, bold: bool = False):
    if bold:
        candidates = [
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\Dengb.ttf",
            FONT_PATH,
        ]
    else:
        candidates = [FONT_PATH]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


F_TITLE = font(34, True)
F_LABEL = font(26, True)
F_BODY = font(23)
F_SMALL = font(20)


@dataclass
class Box:
    x: int
    y: int
    w: int
    h: int
    text: str
    fill: str = COLORS["neutral"]
    stroke: str = COLORS["stroke"]
    font_size: int = 23
    bold: bool = False
    radius: int = 18


@dataclass
class Arrow:
    points: list[tuple[int, int]]
    color: str = COLORS["data"]
    width: int = 4
    dashed: bool = False
    label: str | None = None


@dataclass
class Diagram:
    name: str
    boxes: list[Box]
    arrows: list[Arrow]
    labels: list[tuple[int, int, str, int, bool, str]]


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def text_lines(text: str) -> list[str]:
    return text.split("\n")


def draw_centered_text(draw: ImageDraw.ImageDraw, box: Box) -> None:
    lines = text_lines(box.text)
    line_font = font(box.font_size, box.bold)
    line_h = box.font_size + 8
    total_h = len(lines) * line_h
    y = box.y + (box.h - total_h) / 2
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=line_font)
        x = box.x + (box.w - (bbox[2] - bbox[0])) / 2
        draw.text((x, y), line, fill=hex_to_rgb(COLORS["text"]), font=line_font)
        y += line_h


def arrow_head(p1: tuple[int, int], p2: tuple[int, int], size: int = 14) -> list[tuple[float, float]]:
    angle = math.atan2(p2[1] - p1[1], p2[0] - p1[0])
    return [
        p2,
        (p2[0] - size * math.cos(angle - math.pi / 6), p2[1] - size * math.sin(angle - math.pi / 6)),
        (p2[0] - size * math.cos(angle + math.pi / 6), p2[1] - size * math.sin(angle + math.pi / 6)),
    ]


def draw_arrow(draw: ImageDraw.ImageDraw, arrow: Arrow) -> None:
    color = hex_to_rgb(arrow.color)
    if arrow.dashed:
        for p1, p2 in zip(arrow.points, arrow.points[1:]):
            draw_dashed_line(draw, p1, p2, color, arrow.width)
    else:
        draw.line(arrow.points, fill=color, width=arrow.width, joint="curve")
    if len(arrow.points) >= 2:
        head = arrow_head(arrow.points[-2], arrow.points[-1])
        draw.polygon(head, fill=color)
    if arrow.label:
        mid = arrow.points[len(arrow.points) // 2]
        draw.text((mid[0] + 8, mid[1] - 24), arrow.label, fill=color, font=F_SMALL)


def draw_dashed_line(
    draw: ImageDraw.ImageDraw,
    p1: tuple[int, int],
    p2: tuple[int, int],
    color: tuple[int, int, int],
    width: int,
    dash: int = 18,
    gap: int = 12,
) -> None:
    x1, y1 = p1
    x2, y2 = p2
    length = math.hypot(x2 - x1, y2 - y1)
    if not length:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    pos = 0
    while pos < length:
        end = min(pos + dash, length)
        draw.line(
            [(x1 + ux * pos, y1 + uy * pos), (x1 + ux * end, y1 + uy * end)],
            fill=color,
            width=width,
        )
        pos += dash + gap


def render_png(diagram: Diagram, path: Path) -> None:
    img = Image.new("RGB", (W, H), hex_to_rgb(COLORS["bg"]))
    draw = ImageDraw.Draw(img)
    for arrow in diagram.arrows:
        draw_arrow(draw, arrow)
    for box in diagram.boxes:
        draw.rounded_rectangle(
            [box.x, box.y, box.x + box.w, box.y + box.h],
            radius=box.radius,
            fill=hex_to_rgb(box.fill),
            outline=hex_to_rgb(box.stroke),
            width=3,
        )
        draw_centered_text(draw, box)
    for x, y, text, size, bold, color in diagram.labels:
        draw.text((x, y), text, fill=hex_to_rgb(color), font=font(size, bold))
    img.save(path)


def svg_text(x: int, y: int, text: str, size: int, bold: bool, anchor: str = "middle") -> str:
    weight = "700" if bold else "400"
    lines = text_lines(text)
    tspans = []
    for i, line in enumerate(lines):
        dy = 0 if i == 0 else size + 8
        tspans.append(f'<tspan x="{x}" dy="{dy}">{escape(line)}</tspan>')
    return (
        f'<text x="{x}" y="{y}" text-anchor="{anchor}" font-family="Microsoft YaHei, SimSun, Times New Roman" '
        f'font-size="{size}" font-weight="{weight}" fill="{COLORS["text"]}">'
        + "".join(tspans)
        + "</text>"
    )


def svg_box(box: Box) -> str:
    line_count = len(text_lines(box.text))
    line_h = box.font_size + 8
    y = box.y + (box.h - line_count * line_h) / 2 + box.font_size
    return (
        f'<rect x="{box.x}" y="{box.y}" width="{box.w}" height="{box.h}" rx="{box.radius}" '
        f'fill="{box.fill}" stroke="{box.stroke}" stroke-width="3"/>'
        + svg_text(box.x + box.w // 2, int(y), box.text, box.font_size, box.bold)
    )


def svg_arrow(arrow: Arrow) -> str:
    marker_id = "arrow-feedback" if arrow.color == COLORS["feedback"] else "arrow-data"
    points = " ".join(f"{x},{y}" for x, y in arrow.points)
    dash = ' stroke-dasharray="18 12"' if arrow.dashed else ""
    label = ""
    if arrow.label:
        x, y = arrow.points[len(arrow.points) // 2]
        label = svg_text(x + 36, y - 8, arrow.label, 20, False, "start")
    return (
        f'<polyline points="{points}" fill="none" stroke="{arrow.color}" stroke-width="{arrow.width}" '
        f'stroke-linejoin="round" stroke-linecap="round"{dash} marker-end="url(#{marker_id})"/>'
        + label
    )


def render_svg(diagram: Diagram, path: Path) -> None:
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">',
        "<defs>",
        f'<marker id="arrow-data" markerWidth="12" markerHeight="12" refX="10" refY="6" orient="auto"><path d="M2,2 L10,6 L2,10 Z" fill="{COLORS["data"]}"/></marker>',
        f'<marker id="arrow-feedback" markerWidth="12" markerHeight="12" refX="10" refY="6" orient="auto"><path d="M2,2 L10,6 L2,10 Z" fill="{COLORS["feedback"]}"/></marker>',
        "</defs>",
        f'<rect width="{W}" height="{H}" fill="{COLORS["bg"]}"/>',
    ]
    parts.extend(svg_arrow(a) for a in diagram.arrows)
    parts.extend(svg_box(b) for b in diagram.boxes)
    for x, y, text, size, bold, color in diagram.labels:
        weight = "700" if bold else "400"
        parts.append(
            f'<text x="{x}" y="{y}" font-family="Microsoft YaHei, SimSun, Times New Roman" '
            f'font-size="{size}" font-weight="{weight}" fill="{color}">{escape(text)}</text>'
        )
    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")


def b(x: int, y: int, w: int, h: int, text: str, fill: str, size: int = 23, bold: bool = False) -> Box:
    return Box(x, y, w, h, text, fill=fill, font_size=size, bold=bold)


def diagrams() -> list[Diagram]:
    return [
        Diagram(
            "图3-1_系统总体架构图",
            [
                b(70, 190, 210, 120, "视频源\n内置测试图案\n或视频文件", COLORS["neutral"], 22, True),
                b(350, 130, 300, 240, "TX发送端\nJPEG编码\n帧分片/CRC\nFEC编码\nQPSK调制", COLORS["tx"], 22, True),
                b(720, 150, 300, 200, "信道/干扰环境\nSimulatedChannel\n或 USRP B210\nLFM/MTJ/NAM/NFM/STJ/SIN", COLORS["channel"], 21, True),
                b(1090, 130, 300, 240, "RX接收端\nAGC/CFO校正\n同步解调\nFEC恢复\nJPEG解码/显示", COLORS["rx"], 22, True),
                b(1120, 500, 260, 150, "CCNN识别\nIQ片段\n7类输出\n置信度", COLORS["model"], 22, True),
                b(660, 680, 360, 150, "自适应控制器\n滞后计数器\n调整JPEG质量/FEC/FPS", COLORS["control"], 22, True),
                b(1440, 190, 250, 120, "接收视频\n状态监测\n实时可视化", COLORS["neutral"], 22, True),
            ],
            [
                Arrow([(280, 250), (350, 250)]),
                Arrow([(650, 250), (720, 250)]),
                Arrow([(1020, 250), (1090, 250)]),
                Arrow([(1390, 250), (1440, 250)]),
                Arrow([(1240, 370), (1240, 500)]),
                Arrow([(1120, 575), (1020, 755)]),
                Arrow([(660, 755), (500, 755), (500, 370)], color=COLORS["feedback"], dashed=True, label="策略反馈"),
            ],
            [(72, 70, "数据流：视频业务链路与干扰感知闭环", 30, True, COLORS["text"])],
        ),
        Diagram(
            "图3-2_仿真模式与USRP实机模式对比图",
            [
                b(120, 110, 420, 130, "上层业务逻辑\n视频TX/RX、CCNN推理、自适应策略", COLORS["neutral"], 24, True),
                b(660, 110, 420, 130, "统一软件接口\nsend()/recv()、IQ缓存、状态回调", COLORS["control"], 24, True),
                b(1200, 110, 420, 130, "可视化与实验记录\nPSD、I/Q波形、类别概率、JSR曲线", COLORS["neutral"], 23, True),
                b(180, 390, 560, 210, "仿真模式 sim\nSimulatedChannel\n软件注入干扰\n便于参数扫描和重复实验", COLORS["tx"], 25, True),
                b(1040, 390, 560, 210, "实机模式 full\nUSRP B210 + UHD\n真实射频收发\n验证频偏、增益和缓冲限制", COLORS["rx"], 25, True),
                b(180, 720, 560, 150, "适用：模型验证、链路调试、策略回归测试\n特点：可控、可复现、调试成本低", COLORS["neutral"], 22, False),
                b(1040, 720, 560, 150, "适用：硬件在环验证、实机演示、工程问题定位\n特点：更接近真实链路，但不确定性更高", COLORS["neutral"], 22, False),
            ],
            [
                Arrow([(540, 175), (660, 175)]),
                Arrow([(1080, 175), (1200, 175)]),
                Arrow([(870, 240), (480, 390)]),
                Arrow([(870, 240), (1320, 390)]),
                Arrow([(460, 600), (460, 720)]),
                Arrow([(1320, 600), (1320, 720)]),
            ],
            [(80, 70, "同一套上层算法，两种底层信道实现", 30, True, COLORS["text"])],
        ),
        Diagram(
            "图4-1_CCNN网络结构图",
            [
                b(80, 300, 210, 120, "输入\nIQ序列\n2 × 5000", COLORS["neutral"], 24, True),
                b(360, 260, 260, 200, "CNN特征提取\n局部幅相变化\n频谱纹理\n短时扰动", COLORS["tx"], 22, True),
                b(690, 260, 260, 200, "残差模块\n梯度稳定\n深层特征\n参数轻量", COLORS["rx"], 22, True),
                b(1020, 260, 260, 200, "SE通道注意力\n通道重标定\n突出关键特征", COLORS["control"], 22, True),
                b(1350, 260, 300, 200, "多头因果注意力\n长距离依赖\n在线推理友好", COLORS["model"], 22, True),
                b(690, 620, 300, 150, "全局池化 + 分类器\nSoftmax输出", COLORS["neutral"], 24, True),
                b(1100, 620, 390, 150, "7类结果\nClean / LFM / MTJ / NAM\nNFM / STJ / SIN", COLORS["channel"], 22, True),
            ],
            [
                Arrow([(290, 360), (360, 360)]),
                Arrow([(620, 360), (690, 360)]),
                Arrow([(950, 360), (1020, 360)]),
                Arrow([(1280, 360), (1350, 360)]),
                Arrow([(1500, 460), (1500, 545), (840, 545), (840, 620)]),
                Arrow([(990, 695), (1100, 695)]),
            ],
            [(80, 80, "从IQ原始序列到干扰类别输出的复合网络结构", 30, True, COLORS["text"])],
        ),
        Diagram(
            "图5-1_视频传输收发流程图",
            [
                b(90, 130, 210, 90, "视频帧", COLORS["neutral"], 24, True),
                b(360, 110, 230, 130, "JPEG编码\n质量Q可调", COLORS["tx"], 23, True),
                b(650, 110, 230, 130, "分片封装\n帧头/序号/CRC", COLORS["tx"], 22, True),
                b(940, 110, 230, 130, "FEC编码\n冗余倍数可调", COLORS["tx"], 22, True),
                b(1230, 110, 230, 130, "QPSK调制\nRRC成形", COLORS["tx"], 22, True),
                b(1520, 130, 210, 90, "基带IQ发送", COLORS["neutral"], 23, True),
                b(760, 385, 310, 130, "信道与干扰\nAWGN + 典型干扰\n或USRP射频链路", COLORS["channel"], 22, True),
                b(1520, 700, 210, 90, "IQ接收", COLORS["neutral"], 23, True),
                b(1230, 680, 230, 130, "同步/AGC\nCFO校正", COLORS["rx"], 22, True),
                b(940, 680, 230, 130, "QPSK解调\n符号判决", COLORS["rx"], 22, True),
                b(650, 680, 230, 130, "FEC恢复\nCRC检查", COLORS["rx"], 22, True),
                b(360, 680, 230, 130, "JPEG解码\n帧重组", COLORS["rx"], 22, True),
                b(90, 700, 210, 90, "视频显示", COLORS["neutral"], 24, True),
            ],
            [
                Arrow([(300, 175), (360, 175)]),
                Arrow([(590, 175), (650, 175)]),
                Arrow([(880, 175), (940, 175)]),
                Arrow([(1170, 175), (1230, 175)]),
                Arrow([(1460, 175), (1520, 175)]),
                Arrow([(1625, 220), (1625, 385), (1070, 450)]),
                Arrow([(1070, 450), (1625, 680)]),
                Arrow([(1520, 745), (1460, 745)]),
                Arrow([(1230, 745), (1170, 745)]),
                Arrow([(940, 745), (880, 745)]),
                Arrow([(650, 745), (590, 745)]),
                Arrow([(360, 745), (300, 745)]),
            ],
            [
                (92, 80, "发送端 TX", 28, True, COLORS["data"]),
                (92, 650, "接收端 RX", 28, True, COLORS["data"]),
            ],
        ),
        Diagram(
            "图5-2_自适应传输策略闭环图",
            [
                b(130, 170, 270, 140, "接收侧观测\nIQ片段\nPSD/IQ波形\n链路状态", COLORS["neutral"], 22, True),
                b(500, 160, 280, 160, "CCNN识别\n干扰类别\n置信度\nClean/6类干扰", COLORS["model"], 22, True),
                b(900, 160, 300, 160, "干扰等级判断\nnone / mild\nmoderate / severe", COLORS["control"], 22, True),
                b(1320, 170, 300, 140, "滞后计数器\n8帧确认干扰\n3帧确认等级\n12帧恢复", COLORS["control"], 21, True),
                b(1110, 520, 350, 180, "参数调整\nJPEG质量 85/55/35/15\nFEC 1/2/2/3\nFPS 25/15/8/5", COLORS["tx"], 21, True),
                b(560, 550, 330, 140, "视频传输链路\n编码、分片、调制\n信道传输、解码", COLORS["rx"], 22, True),
                b(130, 560, 280, 120, "视频质量反馈\n画面连续性\n误码/丢帧现象", COLORS["neutral"], 22, True),
            ],
            [
                Arrow([(400, 240), (500, 240)]),
                Arrow([(780, 240), (900, 240)]),
                Arrow([(1200, 240), (1320, 240)]),
                Arrow([(1470, 310), (1470, 610), (1460, 610)]),
                Arrow([(1110, 610), (890, 610)]),
                Arrow([(560, 620), (410, 620)]),
                Arrow([(270, 560), (270, 310)], color=COLORS["feedback"], dashed=True, label="状态回流"),
                Arrow([(1285, 520), (1130, 320)], color=COLORS["feedback"], dashed=True),
            ],
            [(92, 90, "识别结果驱动传输参数动态调整", 30, True, COLORS["text"])],
        ),
    ]


def write_guide(items: list[Diagram]) -> None:
    captions = [
        ("图3-1 系统总体架构图", "建议插入位置：第三章 3.1“系统架构”中“系统整体结构如图3-1所示”之后。"),
        ("图3-2 仿真模式与USRP实机模式对比图", "建议插入位置：第三章 3.2“系统运行模式”开头或模式说明表之前。"),
        ("图4-1 CCNN网络结构图", "建议插入位置：第四章 4.2“CCNN网络结构设计”开头。"),
        ("图5-1 视频传输收发流程图", "建议插入位置：第五章 5.1“TX端设计”和 5.2“RX端设计”之间，作为收发链路总览。"),
        ("图5-2 自适应传输策略闭环图", "建议插入位置：第五章 5.4“自适应传输策略”开头。"),
    ]
    lines = [
        "# 论文插图使用说明",
        "",
        "这些图片按论文插图风格生成，建议优先插入 PNG；如需要后期无损编辑，可使用 SVG。",
        "",
    ]
    for diagram, (caption, position) in zip(items, captions):
        lines.extend(
            [
                f"## {caption}",
                f"- PNG：`{diagram.name}.png`",
                f"- SVG：`{diagram.name}.svg`",
                f"- {position}",
                "",
            ]
        )
    (OUT / "插图使用说明.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    items = diagrams()
    for diagram in items:
        render_png(diagram, OUT / f"{diagram.name}.png")
        render_svg(diagram, OUT / f"{diagram.name}.svg")
    write_guide(items)
    for diagram in items:
        print(OUT / f"{diagram.name}.png")
        print(OUT / f"{diagram.name}.svg")
    print(OUT / "插图使用说明.md")


if __name__ == "__main__":
    main()
