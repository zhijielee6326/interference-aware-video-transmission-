#!/usr/bin/env python3
"""
生成毕业设计答辩PPT。

输出文件：答辩PPT_李智杰.pptx
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "答辩PPT_李智杰.pptx"
FONT = "微软雅黑"

BLUE = RGBColor(18, 73, 132)
BLUE_DARK = RGBColor(10, 47, 92)
BLUE_LIGHT = RGBColor(225, 239, 252)
BLUE_SOFT = RGBColor(240, 247, 253)
WHITE = RGBColor(255, 255, 255)
TEXT = RGBColor(38, 48, 63)
MUTED = RGBColor(93, 110, 128)
ACCENT = RGBColor(29, 142, 178)
GREEN = RGBColor(38, 145, 102)
ORANGE = RGBColor(222, 132, 45)


def img(path: str) -> Path:
    return ROOT / path


def set_text(frame, text, size=18, color=TEXT, bold=False, align=None):
    frame.clear()
    p = frame.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, x, y, w, h, text, size=18, color=TEXT, bold=False, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    set_text(box.text_frame, text, size=size, color=color, bold=bold, align=align)
    return box


def add_title(slide, title: str, subtitle: str | None = None):
    bar = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.68)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = BLUE_DARK
    bar.line.color.rgb = BLUE_DARK
    add_textbox(slide, 0.42, 0.16, 10.8, 0.42, title, size=24, color=WHITE, bold=True)
    if subtitle:
        add_textbox(slide, 11.0, 0.22, 1.9, 0.32, subtitle, size=11, color=BLUE_LIGHT, align=PP_ALIGN.RIGHT)


def add_footer(slide, page: int):
    add_textbox(slide, 11.95, 7.12, 0.9, 0.2, f"{page:02d}", size=9, color=MUTED, align=PP_ALIGN.RIGHT)


def bullet_list(slide, x, y, w, h, items, size=17, color=TEXT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.name = FONT
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(8)
    return box


def add_band(slide, x, y, w, h, color=BLUE_SOFT, line=RGBColor(197, 216, 235)):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.color.rgb = line
    return shape


def add_image(slide, path: Path, x, y, w, h):
    if not path.exists():
        add_band(slide, x, y, w, h, color=RGBColor(252, 238, 238), line=RGBColor(220, 120, 120))
        add_textbox(slide, x + 0.2, y + h / 2 - 0.2, w - 0.4, 0.4, f"图片缺失：{path.relative_to(ROOT)}", size=14, color=RGBColor(160, 50, 50), align=PP_ALIGN.CENTER)
        return None
    return slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def add_table(slide, x, y, w, h, data, col_widths=None, font_size=13):
    rows = len(data)
    cols = len(data[0])
    table = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    if col_widths:
        for i, width in enumerate(col_widths):
            table.columns[i].width = Inches(width)
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(value)
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = BLUE if r == 0 else WHITE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = FONT
                    run.font.size = Pt(font_size)
                    run.font.bold = r == 0
                    run.font.color.rgb = WHITE if r == 0 else TEXT
    return table


def add_flow(slide, labels, y=3.0):
    count = len(labels)
    box_w = 1.55
    gap = 0.38
    start = (13.333 - (count * box_w + (count - 1) * gap)) / 2
    for i, label in enumerate(labels):
        x = start + i * (box_w + gap)
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(box_w), Inches(0.62))
        shape.fill.solid()
        shape.fill.fore_color.rgb = BLUE if i in (0, count - 1) else BLUE_LIGHT
        shape.line.color.rgb = BLUE
        set_text(shape.text_frame, label, size=14, color=WHITE if i in (0, count - 1) else TEXT, bold=True, align=PP_ALIGN.CENTER)
        if i < count - 1:
            add_textbox(slide, x + box_w, y + 0.17, gap, 0.2, "→", size=18, color=BLUE, bold=True, align=PP_ALIGN.CENTER)


def new_slide(prs, title, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE
    add_title(slide, title, subtitle)
    return slide


def build_ppt():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides = []

    # 1
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BLUE_SOFT
    add_textbox(s, 0.8, 0.75, 11.7, 0.8, "面向信号干扰检测与识别的视频传输平台开发", size=30, color=BLUE_DARK, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(s, 1.2, 1.7, 10.8, 0.35, "Interference-Aware Adaptive Video Transmission Platform Based on CCNN", size=15, color=MUTED, align=PP_ALIGN.CENTER)
    add_flow(s, ["视频源", "编码", "QPSK", "USRP", "CCNN识别", "自适应"], y=3.0)
    meta = "李智杰  2022010843\n通信工程（卓越）\n指导教师：刘磊\n北京信息科技大学 信息与通信工程学院"
    add_textbox(s, 4.0, 5.0, 5.3, 1.1, meta, size=16, color=TEXT, align=PP_ALIGN.CENTER)
    slides.append(s)

    # 2
    s = new_slide(prs, "汇报提纲")
    agenda = ["研究背景与意义", "系统总体设计", "CCNN干扰识别模型", "视频传输系统实现", "实验结果与分析", "总结与展望"]
    for i, item in enumerate(agenda):
        y = 1.25 + i * 0.78
        add_band(s, 2.0, y, 9.3, 0.5)
        add_textbox(s, 2.35, y + 0.1, 0.6, 0.25, f"{i+1}", size=15, color=BLUE, bold=True)
        add_textbox(s, 3.0, y + 0.08, 7.5, 0.28, item, size=18, color=TEXT, bold=True)
    slides.append(s)

    # 3
    s = new_slide(prs, "研究背景")
    bullet_list(s, 0.8, 1.15, 5.7, 4.9, [
        "无线通信场景中电磁环境复杂，干扰会导致视频链路质量下降甚至中断",
        "传统功率谱阈值方法可解释性强，但低JSR下区分能力有限",
        "深度学习可直接从IQ信号中学习时域和频域特征",
        "需要从算法识别走向可演示、可验证的软件无线电系统"
    ], size=18)
    add_band(s, 7.0, 1.25, 4.9, 3.6)
    add_textbox(s, 7.4, 1.65, 4.1, 0.4, "本文关注的问题", size=20, color=BLUE_DARK, bold=True)
    bullet_list(s, 7.45, 2.35, 3.9, 1.8, ["识别：干扰类型是什么", "决策：链路参数如何调整", "验证：硬件平台能否闭环运行"], size=16)
    slides.append(s)

    # 4
    s = new_slide(prs, "国内外研究现状")
    add_table(s, 0.8, 1.2, 11.7, 3.7, [
        ["方向", "典型方法", "不足"],
        ["干扰检测", "PSD、谱熵、统计阈值", "低JSR下灵敏度不足，类型区分能力弱"],
        ["深度识别", "CNN、Transformer、时频图分类", "需要与真实链路结合验证"],
        ["软件无线电", "USRP/GNU Radio硬件在环", "实时处理和缓冲调度要求高"],
        ["自适应传输", "AMC、FEC、码率控制", "需要稳定的干扰感知输入"]
    ], col_widths=[2.0, 4.2, 5.5], font_size=12)
    add_textbox(s, 1.2, 5.4, 10.6, 0.5, "本文将 CCNN 干扰识别、视频传输链路和 USRP 实机验证整合为完整平台。", size=19, color=BLUE_DARK, bold=True, align=PP_ALIGN.CENTER)
    slides.append(s)

    # 5
    s = new_slide(prs, "本文主要工作")
    bullet_list(s, 1.0, 1.2, 11.0, 4.7, [
        "构建7类干扰识别模型：LFM、MTJ、NAM、NFM、STJ、SIN、Clean",
        "实现视频源、JPEG、分包、FEC、QPSK调制解调和RF收发链路",
        "设计多级滞后自适应策略，动态调整JPEG质量、FEC和帧率",
        "完成仿真模式、USRP full模式和实时可视化面板验证",
        "形成论文、Word文档和答辩PPT等交付材料"
    ], size=19)
    slides.append(s)

    # 6
    s = new_slide(prs, "系统总体架构")
    add_flow(s, ["视频源", "JPEG", "分包/FEC", "QPSK", "信道/USRP", "AGC/CFO", "CCNN", "自适应"], y=2.1)
    add_table(s, 1.2, 4.0, 10.8, 1.5, [
        ["模块", "功能"],
        ["TX链路", "视频采集、压缩、编码、调制、发射"],
        ["RX链路", "接收、同步、解调、解码、显示"],
        ["控制链路", "干扰识别、等级判定、参数反馈"]
    ], col_widths=[2.4, 8.4], font_size=13)
    slides.append(s)

    # 7
    s = new_slide(prs, "运行模式与关键参数")
    add_table(s, 0.9, 1.2, 11.5, 4.7, [
        ["参数", "仿真模式", "USRP实机模式"],
        ["视频分辨率", "320×240", "160×120"],
        ["基准JPEG质量", "70", "30"],
        ["基准帧率", "25 fps", "5 fps"],
        ["调制方式", "QPSK", "QPSK"],
        ["采样率", "2 MHz", "2 MHz"],
        ["中心频率", "2.45 GHz", "2.45 GHz"]
    ], col_widths=[3.0, 4.2, 4.3], font_size=13)
    slides.append(s)

    # 8
    s = new_slide(prs, "7类干扰信号模型")
    add_image(s, img("gnuradio/figure_4_2_interference_waveforms.png"), 0.8, 1.05, 11.7, 5.45)
    slides.append(s)

    # 9
    s = new_slide(prs, "CCNN网络结构")
    add_flow(s, ["IQ输入", "CNN", "ResNet+SE", "TCC注意力", "AvgPool", "Softmax"], y=2.0)
    add_table(s, 1.2, 3.6, 10.8, 2.1, [
        ["模块", "作用", "参数量占比"],
        ["CNN特征提取", "提取局部时域特征", "39.2%"],
        ["残差网络+SE", "增强重要通道并稳定训练", "20.3%"],
        ["多头因果注意力", "建模序列依赖关系", "37.1%"],
        ["分类器", "输出7类概率", "3.3%"]
    ], col_widths=[2.5, 5.9, 2.4], font_size=12)
    slides.append(s)

    # 10
    s = new_slide(prs, "模型训练结果")
    add_image(s, img("gnuradio/figure_4_4_training_curve.png"), 0.8, 1.0, 5.8, 4.9)
    add_table(s, 7.0, 1.2, 5.2, 3.2, [
        ["指标", "结果"],
        ["最佳模型", "epoch 26"],
        ["总体准确率", "99.91%"],
        ["参数量", "96,921"],
        ["输入长度", "5000点IQ信号"]
    ], col_widths=[2.0, 3.2], font_size=13)
    add_textbox(s, 7.0, 4.85, 5.1, 0.7, "模型规模小，适合实时推理；在中高JSR下分类稳定。", size=17, color=BLUE_DARK, bold=True)
    slides.append(s)

    # 11
    s = new_slide(prs, "混淆矩阵")
    add_image(s, img("gnuradio/figure_4_5_confusion_matrix.png"), 1.1, 1.0, 7.0, 5.7)
    bullet_list(s, 8.5, 1.5, 3.6, 3.7, [
        "7类分类任务",
        "Clean类误报率低",
        "主要混淆集中在低JSR窄带类",
        "为自适应策略提供稳定输入"
    ], size=17)
    slides.append(s)

    # 12
    s = new_slide(prs, "JSR鲁棒性分析")
    add_image(s, img("gnuradio/jsr_sweep_results.png"), 0.8, 1.0, 6.5, 5.5)
    add_table(s, 7.65, 1.4, 4.8, 2.6, [
        ["JSR区间", "识别表现"],
        ["≥ 0 dB", "约99.9%"],
        ["-2 dB", "约99.4%"],
        ["-4 dB", "约96.0%"],
        ["-6 dB", "约65.3%"]
    ], col_widths=[1.9, 2.9], font_size=13)
    add_textbox(s, 7.75, 4.45, 4.5, 0.8, "低JSR下NAM/STJ更难区分，后续可增加低JSR样本密度。", size=16, color=TEXT)
    slides.append(s)

    # 13
    s = new_slide(prs, "视频传输TX/RX链路")
    add_flow(s, ["视频帧", "JPEG", "分包", "FEC", "QPSK", "USRP发送"], y=1.55)
    add_flow(s, ["USRP接收", "AGC", "同步/CFO", "解调", "FEC解码", "视频显示"], y=3.1)
    add_table(s, 1.0, 4.75, 11.2, 1.2, [
        ["链路设计", "关键点"],
        ["帧结构", "14字节协议头 + 480字节载荷 + CRC16"],
        ["物理层", "QPSK，RRC滚降系数0.35，2 MHz采样率"]
    ], col_widths=[2.2, 9.0], font_size=12)
    slides.append(s)

    # 14
    s = new_slide(prs, "功率谱监测与实时可视化")
    add_image(s, img("gnuradio/fig_psd_comparison.png"), 0.8, 1.0, 6.1, 5.4)
    bullet_list(s, 7.3, 1.25, 4.7, 4.2, [
        "PSD实时显示当前RF频谱",
        "CCNN概率柱状图展示7类置信度",
        "I/Q时域波形辅助判断接收状态",
        "JSR曲线用于展示模型鲁棒性"
    ], size=17)
    slides.append(s)

    # 15
    s = new_slide(prs, "自适应传输策略")
    add_table(s, 0.8, 1.05, 11.7, 2.6, [
        ["等级", "触发类型", "JPEG质量", "FEC", "FPS"],
        ["none", "Clean", "85", "1×", "25"],
        ["mild", "NAM/NFM/SIN", "55", "2×", "15"],
        ["moderate", "LFM/MTJ", "35", "2×", "8"],
        ["severe", "STJ", "15", "3×", "5"]
    ], col_widths=[2.0, 3.3, 2.0, 1.7, 1.7], font_size=12)
    add_image(s, img("gnuradio/fig_adaptive_transmission.png"), 1.2, 4.0, 5.1, 2.5)
    add_image(s, img("gnuradio/fig_adaptive_comparison.png"), 7.0, 4.0, 5.1, 2.5)
    slides.append(s)

    # 16
    s = new_slide(prs, "仿真模式测试结果")
    add_table(s, 1.0, 1.1, 11.2, 3.0, [
        ["测试项", "结果"],
        ["30秒可视化演示", "发送125帧，接收125帧"],
        ["干扰事件", "5次"],
        ["策略切换", "5次"],
        ["离线CCNN测试", "70/70，100.00%"],
        ["自适应逻辑", "8帧干扰确认、3帧等级确认、12帧恢复通过"]
    ], col_widths=[3.3, 7.9], font_size=13)
    add_textbox(s, 1.3, 4.75, 10.6, 0.7, "仿真模式验证了系统端到端闭环、模型推理和自适应控制逻辑。", size=19, color=BLUE_DARK, bold=True, align=PP_ALIGN.CENTER)
    slides.append(s)

    # 17
    s = new_slide(prs, "USRP实机与长时间可视化")
    add_image(s, img("gnuradio/usrp_test_report.png"), 0.8, 1.0, 5.6, 4.8)
    add_table(s, 6.8, 1.2, 5.6, 3.4, [
        ["项目", "结果"],
        ["设备", "USRP B210, serial 7MFTKFU"],
        ["连接", "USB 3.0，UHD probe通过"],
        ["120秒full长测", "发送482帧，接收处理块823"],
        ["实时面板", "PSD/概率/IQ/JSR四面板正常刷新"]
    ], col_widths=[2.0, 3.6], font_size=12)
    add_textbox(s, 6.95, 5.2, 5.2, 0.65, "实机模式证明系统具备硬件在环运行和演示能力。", size=17, color=BLUE_DARK, bold=True)
    slides.append(s)

    # 18
    s = new_slide(prs, "USRP JSR实测结果")
    add_image(s, img("USRP/jsr_accuracy.png"), 0.8, 1.0, 6.1, 5.4)
    add_table(s, 7.25, 1.15, 4.9, 3.3, [
        ["类型", "中高JSR表现"],
        ["LFM", "稳定"],
        ["MTJ", "稳定"],
        ["NFM", "10dB以上稳定"],
        ["STJ", "6dB以上较稳定"],
        ["NAM/SIN", "仍需优化"]
    ], col_widths=[1.8, 3.1], font_size=13)
    slides.append(s)

    # 19
    s = new_slide(prs, "主要工作总结")
    bullet_list(s, 1.0, 1.1, 11.0, 4.8, [
        "完成7类CCNN干扰识别模型，测试准确率99.91%",
        "实现视频传输完整链路：编码、分包、FEC、QPSK、接收解调",
        "设计稳定的自适应传输策略，能够根据干扰类型切换参数",
        "完成仿真、USRP实机和实时可视化多层验证",
        "形成可演示的视频传输平台和完整毕业设计材料"
    ], size=19)
    slides.append(s)

    # 20
    s = new_slide(prs, "不足与展望")
    bullet_list(s, 1.0, 1.1, 11.0, 4.9, [
        "当前主要在应用层调节JPEG质量、FEC和帧率，后续可加入物理层抗干扰",
        "低JSR窄带干扰仍有混淆，可增加低JSR样本和时频域特征",
        "实时USRP链路仍存在overflow，可拆分采集线程和推理线程",
        "FEC可升级为Reed-Solomon或LDPC以提升纠错能力",
        "未来可扩展到多节点协同检测和在线模型更新"
    ], size=18)
    slides.append(s)

    # 21
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BLUE_DARK
    add_textbox(s, 0.8, 2.55, 11.7, 0.7, "谢谢各位老师", size=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(s, 0.8, 3.45, 11.7, 0.35, "请批评指正", size=20, color=BLUE_LIGHT, align=PP_ALIGN.CENTER)
    slides.append(s)

    for idx, slide in enumerate(slides, start=1):
        add_footer(slide, idx)

    prs.save(OUTPUT)
    print(f"已生成：{OUTPUT}")
    print(f"幻灯片页数：{len(prs.slides)}")


if __name__ == "__main__":
    build_ppt()
