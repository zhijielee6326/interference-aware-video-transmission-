from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(r"C:\Users\m1889\Desktop\毕设")
SOURCE_MD = ROOT / "毕业论文.md"
OUT_DIR = ROOT / "论文图表优化" / "表格图片"
OUT_MD = ROOT / "论文图表优化" / "毕业论文_图表图片版.md"

TABLE_TITLES = [
    "表2-1 USRP B210主要参数",
    "表3-1 系统运行模式对比",
    "表3-2 系统参数配置",
    "表4-1 训练数据集构成",
    "表4-2 训练超参数配置",
    "表4-3 不同JSR下的识别准确率（%）",
    "表4-4 MATLAB与Python数据生成对比",
    "表5-1 数据包字段说明",
    "表5-2 干扰等级划分与传输参数调整",
    "表6-1 实验环境",
    "表6-2 各类别分类性能",
    "表6-3 CCNN模型各模块参数量分析",
    "表6-4 不同方法在7类干扰识别任务上的性能对比",
    "表6-5 仿真模式无干扰测试",
    "表6-6 自动演示测试序列",
    "表6-7 USRP实机模式120秒测试结果",
    "表6-8 USRP实机可视化长测结果",
    "表6-9 自适应策略有效性验证",
]


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


TITLE_FONT = font(44, True)
HEADER_FONT = font(31, True)
BODY_FONT = font(29, False)
BODY_BOLD_FONT = font(29, True)


def clean_cell(value: str) -> tuple[str, bool]:
    value = value.strip()
    bold = value.startswith("**") and value.endswith("**")
    value = value.replace("**", "")
    value = value.replace("<br>", "\n")
    value = re.sub(r"\s+", " ", value)
    return value.strip(), bold


def split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [cell.strip() for cell in line.split("|")]


def is_sep(line: str) -> bool:
    cells = split_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", c.strip()) for c in cells)


def find_tables(lines: list[str]) -> list[tuple[int, int, list[list[str]]]]:
    tables: list[tuple[int, int, list[list[str]]]] = []
    i = 0
    while i < len(lines) - 1:
        if lines[i].lstrip().startswith("|") and is_sep(lines[i + 1]):
            start = i
            raw = [split_row(lines[i])]
            i += 2
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                raw.append(split_row(lines[i]))
                i += 1
            tables.append((start, i, raw))
        else:
            i += 1
    return tables


def text_width(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont) -> int:
    if not text:
        return 0
    box = draw.textbbox((0, 0), text, font=fnt)
    return box[2] - box[0]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, fnt: ImageFont.ImageFont, max_width: int) -> list[str]:
    if "\n" in text:
        lines: list[str] = []
        for part in text.split("\n"):
            lines.extend(wrap_text(draw, part, fnt, max_width))
        return lines

    tokens: list[str] = []
    buf = ""
    for ch in text:
        if ord(ch) < 128 and not ch.isspace():
            buf += ch
        else:
            if buf:
                tokens.append(buf)
                buf = ""
            if not ch.isspace():
                tokens.append(ch)
    if buf:
        tokens.append(buf)

    lines: list[str] = []
    cur = ""
    for token in tokens:
        candidate = cur + token
        if cur and text_width(draw, candidate, fnt) > max_width:
            lines.append(cur)
            cur = token
        else:
            cur = candidate
    if cur:
        lines.append(cur)
    return lines or [""]


def weighted_widths(rows: list[list[str]], draw: ImageDraw.ImageDraw, usable: int) -> list[int]:
    col_count = max(len(row) for row in rows)
    weights = [1.0] * col_count
    for c in range(col_count):
        values = [row[c] if c < len(row) else "" for row in rows]
        max_px = max(text_width(draw, clean_cell(v)[0], BODY_FONT) for v in values)
        avg_len = sum(len(clean_cell(v)[0]) for v in values) / max(1, len(values))
        weights[c] = max(0.9, min(2.8, max_px / 280 + avg_len / 22))
    total = sum(weights)
    widths = [int(usable * w / total) for w in weights]
    delta = usable - sum(widths)
    widths[-1] += delta
    return widths


def draw_table(rows: list[list[str]], title: str, out_path: Path) -> None:
    tmp = Image.new("RGB", (2400, 1200), "white")
    draw = ImageDraw.Draw(tmp)

    margin_x = 110
    top = 72
    title_h = 76
    usable = 2400 - margin_x * 2
    col_widths = weighted_widths(rows, draw, usable)
    pad_x = 24
    pad_y = 18
    line_gap = 8

    row_heights: list[int] = []
    wrapped_rows: list[list[tuple[list[str], bool]]] = []
    for r, row in enumerate(rows):
        wrapped_row: list[tuple[list[str], bool]] = []
        max_lines = 1
        for c, cell in enumerate(row):
            value, bold = clean_cell(cell)
            fnt = HEADER_FONT if r == 0 else (BODY_BOLD_FONT if bold else BODY_FONT)
            lines = wrap_text(draw, value, fnt, col_widths[c] - pad_x * 2)
            max_lines = max(max_lines, len(lines))
            wrapped_row.append((lines, bold))
        wrapped_rows.append(wrapped_row)
        base_font_size = 31 if r == 0 else 29
        row_heights.append(max(66, pad_y * 2 + max_lines * base_font_size + (max_lines - 1) * line_gap))

    height = top + title_h + sum(row_heights) + 42
    img = Image.new("RGB", (2400, height), "white")
    draw = ImageDraw.Draw(img)

    title_box = draw.textbbox((0, 0), title, font=TITLE_FONT)
    draw.text(((2400 - (title_box[2] - title_box[0])) / 2, top), title, fill="#1f4e79", font=TITLE_FONT)

    x0 = margin_x
    y = top + title_h
    border = "#2f2f2f"
    header_fill = "#eaf2fb"
    alt_fill = "#fafafa"
    total_width = sum(col_widths)

    for r, row in enumerate(rows):
        x = x0
        fill = header_fill if r == 0 else (alt_fill if r % 2 == 0 else "white")
        for c in range(len(col_widths)):
            w = col_widths[c]
            draw.rectangle([x, y, x + w, y + row_heights[r]], fill=fill, outline=border, width=2)
            if c < len(wrapped_rows[r]):
                lines, bold = wrapped_rows[r][c]
                fnt = HEADER_FONT if r == 0 else (BODY_BOLD_FONT if bold else BODY_FONT)
                text_h = len(lines) * fnt.size + (len(lines) - 1) * line_gap
                ty = y + (row_heights[r] - text_h) / 2 - 2
                for line in lines:
                    tw = text_width(draw, line, fnt)
                    tx = x + (w - tw) / 2
                    draw.text((tx, ty), line, fill="#111111", font=fnt)
                    ty += fnt.size + line_gap
            x += w
        y += row_heights[r]

    draw.rectangle([x0, top + title_h, x0 + total_width, y], outline="#111111", width=4)
    img.save(out_path, quality=95)


def markdown_image_path(path: Path) -> str:
    return path.relative_to(OUT_MD.parent).as_posix()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    lines = SOURCE_MD.read_text(encoding="utf-8").splitlines()
    tables = find_tables(lines)
    if len(tables) > len(TABLE_TITLES):
        raise RuntimeError(f"Need more table titles: found {len(tables)} tables")

    replacements: dict[tuple[int, int], list[str]] = {}
    for idx, (start, end, rows) in enumerate(tables, start=1):
        title = TABLE_TITLES[idx - 1]
        safe = re.sub(r"[\\/:*?\"<>|（）()\\s]+", "_", title).strip("_")
        out_path = OUT_DIR / f"{idx:02d}_{safe}.png"
        draw_table(rows, title, out_path)
        replacements[(start, end)] = [
            f"![{title}]({markdown_image_path(out_path)})",
            "",
            f"*{title}*",
        ]

    out_lines: list[str] = []
    i = 0
    ranges = sorted(replacements)
    range_idx = 0
    while i < len(lines):
        if range_idx < len(ranges) and i == ranges[range_idx][0]:
            start, end = ranges[range_idx]
            out_lines.extend(replacements[(start, end)])
            i = end
            range_idx += 1
        else:
            out_lines.append(lines[i])
            i += 1

    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text("\n".join(out_lines) + "\n", encoding="utf-8")
    print(f"rendered_tables={len(tables)}")
    print(f"out_dir={OUT_DIR}")
    print(f"out_md={OUT_MD}")


if __name__ == "__main__":
    main()
