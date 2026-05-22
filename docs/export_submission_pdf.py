from __future__ import annotations

import html
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(r"C:\Users\m1889\Desktop\毕设")
OUT = ROOT / "提交材料_李智杰"
DOCX = OUT / "毕业论文_李智杰_终稿.docx"
PDF = OUT / "毕业论文_李智杰_终稿.pdf"
TITLE = "面向信号干扰检测与识别的视频传输平台开发"


def register_fonts() -> tuple[str, str]:
    regular = Path(r"C:\Windows\Fonts\Deng.ttf")
    bold = Path(r"C:\Windows\Fonts\Dengb.ttf")
    kai = Path(r"C:\Windows\Fonts\simkai.ttf")
    if not regular.exists():
        regular = Path(r"C:\Windows\Fonts\simkai.ttf")
    if not bold.exists():
        bold = regular
    if not kai.exists():
        kai = regular
    pdfmetrics.registerFont(TTFont("CJK", str(regular)))
    pdfmetrics.registerFont(TTFont("CJK-Bold", str(bold)))
    pdfmetrics.registerFont(TTFont("CJK-Kai", str(kai)))
    return "CJK", "CJK-Bold"


FONT, FONT_BOLD = register_fonts()


def make_styles():
    base = getSampleStyleSheet()
    styles = {
        "normal": ParagraphStyle(
            "normal-cjk",
            parent=base["Normal"],
            fontName=FONT,
            fontSize=10.5,
            leading=18,
            firstLineIndent=18,
            alignment=TA_JUSTIFY,
            wordWrap="CJK",
            spaceAfter=3,
        ),
        "h1": ParagraphStyle(
            "h1-cjk",
            parent=base["Heading1"],
            fontName=FONT_BOLD,
            fontSize=16,
            leading=22,
            alignment=TA_CENTER,
            wordWrap="CJK",
            spaceBefore=10,
            spaceAfter=10,
        ),
        "h2": ParagraphStyle(
            "h2-cjk",
            parent=base["Heading2"],
            fontName=FONT_BOLD,
            fontSize=12,
            leading=18,
            alignment=TA_LEFT,
            wordWrap="CJK",
            spaceBefore=8,
            spaceAfter=5,
        ),
        "h3": ParagraphStyle(
            "h3-cjk",
            parent=base["Heading3"],
            fontName=FONT_BOLD,
            fontSize=10.5,
            leading=18,
            alignment=TA_LEFT,
            wordWrap="CJK",
            spaceBefore=6,
            spaceAfter=4,
        ),
        "caption": ParagraphStyle(
            "caption-cjk",
            parent=base["Normal"],
            fontName=FONT,
            fontSize=9,
            leading=14,
            alignment=TA_CENTER,
            wordWrap="CJK",
            spaceAfter=5,
        ),
        "toc": ParagraphStyle(
            "toc-cjk",
            parent=base["Normal"],
            fontName=FONT,
            fontSize=10,
            leading=16,
            leftIndent=18,
            wordWrap="CJK",
        ),
        "cell": ParagraphStyle(
            "cell-cjk",
            parent=base["Normal"],
            fontName=FONT,
            fontSize=7.5,
            leading=10,
            wordWrap="CJK",
        ),
        "cover_school": ParagraphStyle(
            "cover-school",
            parent=base["Normal"],
            fontName="CJK-Kai",
            fontSize=26,
            leading=32,
            alignment=TA_CENTER,
            wordWrap="CJK",
            spaceBefore=22,
            spaceAfter=20,
        ),
        "cover_title": ParagraphStyle(
            "cover-title",
            parent=base["Normal"],
            fontName="CJK-Kai",
            fontSize=42,
            leading=50,
            alignment=TA_CENTER,
            wordWrap="CJK",
            spaceBefore=8,
            spaceAfter=54,
        ),
        "cover_field": ParagraphStyle(
            "cover-field",
            parent=base["Normal"],
            fontName="CJK-Kai",
            fontSize=14,
            leading=22,
            leftIndent=28,
            firstLineIndent=0,
            alignment=TA_LEFT,
            wordWrap="CJK",
            spaceAfter=16,
        ),
        "declaration_title": ParagraphStyle(
            "declaration-title",
            parent=base["Heading1"],
            fontName=FONT_BOLD,
            fontSize=16,
            leading=22,
            alignment=TA_CENTER,
            wordWrap="CJK",
            spaceBefore=10,
            spaceAfter=12,
        ),
    }
    return styles


STYLES = make_styles()


def iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield DocxParagraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield DocxTable(child, doc)


def para_style(paragraph: DocxParagraph):
    style = paragraph.style.name if paragraph.style else ""
    text = paragraph.text.strip()
    compact = " ".join(text.split())
    if compact == "北京信息科技大学":
        return STYLES["cover_school"]
    if compact == "毕业设计（论文）":
        return STYLES["cover_title"]
    if compact.startswith(("题", "学", "专", "学生姓名", "指导老师", "指导教师", "起止时间")) and "：" in compact:
        return STYLES["cover_field"]
    if compact in {"毕业设计（论文）原创性声明", "毕业设计（论文）版权授权声明"}:
        return STYLES["declaration_title"]
    if style == "Heading 1":
        return STYLES["h1"]
    if style == "Heading 2":
        return STYLES["h2"]
    if style == "Heading 3":
        return STYLES["h3"]
    if text.startswith("图") or text.startswith("表"):
        return STYLES["caption"]
    return STYLES["normal"]


def clean_text(text: str) -> str:
    return html.escape(" ".join((text or "").split()))


def heading_items(doc: Document) -> list[tuple[int, str]]:
    items: list[tuple[int, str]] = []
    for paragraph in doc.paragraphs:
        style = paragraph.style.name if paragraph.style else ""
        text = " ".join(paragraph.text.split())
        if not text:
            continue
        if style == "Heading 1":
            items.append((1, text))
        elif style == "Heading 2":
            items.append((2, text))
        elif style == "Heading 3":
            items.append((3, text))
    return items


def paragraph_images(doc: Document, paragraph: DocxParagraph) -> list[Image]:
    images: list[Image] = []
    for run in paragraph.runs:
        for blip in run._r.xpath(".//*[local-name()='blip']"):
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rid or rid not in doc.part.related_parts:
                continue
            blob = doc.part.related_parts[rid].blob
            image = Image(BytesIO(blob))
            max_w = 14.5 * cm
            max_h = 9 * cm
            scale = min(max_w / image.imageWidth, max_h / image.imageHeight, 1)
            image.drawWidth = image.imageWidth * scale
            image.drawHeight = image.imageHeight * scale
            image.hAlign = "CENTER"
            images.append(image)
    return images


def table_flowable(table: DocxTable):
    rows = []
    for row in table.rows:
        values = []
        for cell in row.cells:
            text = clean_text(cell.text)
            values.append(Paragraph(text if text else " ", STYLES["cell"]))
        rows.append(values)
    if not rows:
        return Spacer(1, 1)
    cols = max(len(row) for row in rows)
    width = 16 * cm
    col_widths = [width / cols] * cols
    flow = Table(rows, colWidths=col_widths, repeatRows=1)
    flow.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return flow


def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT, 8)
    canvas.drawCentredString(A4[0] / 2, A4[1] - 1.35 * cm, TITLE)
    canvas.drawCentredString(A4[0] / 2, 1.05 * cm, str(canvas.getPageNumber()))
    canvas.restoreState()


def on_first_page(canvas, doc):
    return


def build_pdf() -> None:
    doc = Document(str(DOCX))
    toc_items = heading_items(doc)
    story = []
    skip_next_toc_placeholder = False
    for block in iter_blocks(doc):
        if isinstance(block, DocxParagraph):
            text = clean_text(block.text)
            images = paragraph_images(doc, block)
            has_page_break = bool(block.runs and any("w:br" in run._r.xml and 'type="page"' in run._r.xml for run in block.runs))
            if not text and not images:
                if has_page_break:
                    story.append(PageBreak())
                continue
            if "目录域：请在 Word 中更新目录" in text:
                continue
            if skip_next_toc_placeholder:
                skip_next_toc_placeholder = False
            if text == "目录":
                story.append(Paragraph(text, STYLES["h1"]))
                for level, title in toc_items:
                    if title == "目录":
                        continue
                    indent = "&nbsp;" * (level - 1) * 4
                    story.append(Paragraph(f"{indent}{clean_text(title)}", STYLES["toc"]))
                story.append(PageBreak())
                skip_next_toc_placeholder = True
                continue
            story.append(Paragraph(text, para_style(block)))
            for image in images:
                story.append(image)
                story.append(Spacer(1, 5))
            if has_page_break:
                story.append(PageBreak())
        else:
            story.append(table_flowable(block))
            story.append(Spacer(1, 6))
    pdf = SimpleDocTemplate(
        str(PDF),
        pagesize=A4,
        leftMargin=3 * cm,
        rightMargin=2 * cm,
        topMargin=2.2 * cm,
        bottomMargin=2.0 * cm,
        title=TITLE,
        author="李智杰",
    )
    pdf.build(story, onFirstPage=on_first_page, onLaterPages=on_page)
    print(PDF)


if __name__ == "__main__":
    build_pdf()
