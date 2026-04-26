#!/usr/bin/env python3
"""
AI漫剧制作课程 PPT 批量生成脚本 v3
视觉丰富版 - 多种布局、数据可视化、装饰元素
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ── 配色方案 ──
COLORS = {
    'bg_dark':    RGBColor(0x0B, 0x11, 0x20),
    'bg_card':    RGBColor(0x14, 0x1E, 0x33),
    'bg_card2':   RGBColor(0x1A, 0x28, 0x42),
    'accent':     RGBColor(0x00, 0xD4, 0xFF),
    'accent2':    RGBColor(0x7C, 0x3A, 0xED),
    'accent3':    RGBColor(0xFF, 0x6B, 0x35),
    'accent4':    RGBColor(0x10, 0xB9, 0x81),
    'accent5':    RGBColor(0xF5, 0x9E, 0x0B),
    'accent6':    RGBColor(0xEF, 0x44, 0x44),
    'text_white': RGBColor(0xFF, 0xFF, 0xFF),
    'text_light': RGBColor(0xCB, 0xD5, 0xE8),
    'text_dim':   RGBColor(0x6B, 0x7B, 0x9E),
    'border':     RGBColor(0x2D, 0x3B, 0x55),
    'gradient1':  RGBColor(0x0E, 0x1A, 0x2D),
    'gradient2':  RGBColor(0x12, 0x24, 0x3E),
}

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
OUTPUT_DIR = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程/PPT课件'

# 模块主题色
MODULE_COLORS = [
    COLORS['accent'],    # 0-导论
    COLORS['accent3'],   # 1-行业
    COLORS['accent'],    # 2-LLM
    COLORS['accent2'],   # 3-分镜
    COLORS['accent3'],   # 4-即梦
    COLORS['accent4'],   # 5-其他工具
    COLORS['accent'],    # 6-可灵
    COLORS['accent2'],   # 7-Seedance
    COLORS['accent5'],   # 8-海螺
    COLORS['accent6'],   # 9-角色一致
    COLORS['accent4'],   # 10-音频
    COLORS['accent3'],   # 11-剪辑
    COLORS['accent'],    # 12-一站式
    COLORS['accent2'],   # 13-ComfyUI
    COLORS['accent5'],   # 14-工业化
    COLORS['accent4'],   # 15-变现
    COLORS['accent6'],   # 16-合规
    COLORS['accent3'],   # 17-实战
]


def get_module_color(num):
    if 0 <= num < len(MODULE_COLORS):
        return MODULE_COLORS[num]
    return COLORS['accent']


# ══════════════════════════════════════════════
# 基础图形工具
# ══════════════════════════════════════════════

def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, left, top, width, height, fill_color=None, border_color=None, border_width=Pt(0)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
    return shape


def add_rounded_rect(slide, left, top, width, height, fill_color=None, border_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
    return shape


def add_circle(slide, left, top, size, fill_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18, color=None, bold=False,
                 alignment=PP_ALIGN.LEFT, font_name='Microsoft YaHei'):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color or COLORS['text_white']
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_multi_text(slide, left, top, width, height, lines, font_size=16, color=None,
                   line_spacing=1.5, bullet_color=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if line.strip():
            p.text = f"▸ {line}"
        else:
            p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color or COLORS['text_light']
        p.font.name = 'Microsoft YaHei'
        p.space_after = Pt(font_size * 0.5)
        # 设置 bullet 颜色需要通过 XML
        if bullet_color and line.strip():
            run = p.runs[0] if p.runs else None
            if run:
                rPr = run._r.get_or_add_rPr()
                solidFill = rPr.makeelement(qn('a:solidFill'), {})
                srgb = solidFill.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (bullet_color[0], bullet_color[1], bullet_color[2]) if isinstance(bullet_color, tuple) else '00D4FF'})
                solidFill.append(srgb)
    return txBox


def add_accent_line(slide, left, top, width, color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color or COLORS['accent']
    shape.line.fill.background()
    return shape


def add_footer(slide, module_num, module_name, page_num):
    add_shape(slide, Inches(0), Inches(7.05), SLIDE_WIDTH, Pt(1.5), fill_color=COLORS['border'])
    add_text_box(slide, Inches(0.5), Inches(7.1), Inches(6), Inches(0.3),
                 f"模块{module_num:02d}：{module_name}", font_size=9, color=COLORS['text_dim'])
    add_text_box(slide, Inches(11.5), Inches(7.1), Inches(1.5), Inches(0.3),
                 str(page_num), font_size=9, color=COLORS['text_dim'], alignment=PP_ALIGN.RIGHT)


def add_corner_decoration(slide, color=None):
    """右上角装饰几何"""
    c = color or COLORS['accent']
    # 小三角装饰
    add_shape(slide, Inches(12.0), Inches(0), Inches(1.333), Inches(0.08), fill_color=c)
    add_shape(slide, Inches(13.18), Inches(0), Inches(0.15), Inches(1.0), fill_color=c)


def add_page_header(slide, title, accent_color=None):
    """统一页面标题区"""
    c = accent_color or COLORS['accent']
    # 标题背景条
    add_shape(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(1.15), fill_color=COLORS['bg_card'])
    # 左侧色条
    add_shape(slide, Inches(0), Inches(0), Inches(0.12), Inches(1.15), fill_color=c)
    # 标题文字
    add_text_box(slide, Inches(0.6), Inches(0.2), Inches(11), Inches(0.7),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    # 底部线
    add_shape(slide, Inches(0), Inches(1.15), SLIDE_WIDTH, Pt(2), fill_color=c)


# ══════════════════════════════════════════════
# 幻灯片类型 - 高级版
# ══════════════════════════════════════════════

def make_title_slide(prs, title, subtitle, module_num=0):
    """封面页 - 全新设计"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    c = get_module_color(module_num)

    # 左侧大色块
    add_shape(slide, Inches(0), Inches(0), Inches(5.5), SLIDE_HEIGHT, fill_color=c)

    # 左上角几何装饰
    add_shape(slide, Inches(0), Inches(0), Inches(5.5), Inches(0.12), fill_color=COLORS['text_white'])
    add_shape(slide, Inches(0), Inches(0), Inches(0.12), Inches(7.5), fill_color=COLORS['text_white'])

    # 模块编号（大号）
    if module_num > 0:
        add_text_box(slide, Inches(0.5), Inches(1.5), Inches(4.5), Inches(2.5),
                     f"{module_num:02d}", font_size=140, color=RGBColor(0xFF, 0xFF, 0xFF), bold=True)
        # MODULE 标签
        add_text_box(slide, Inches(0.6), Inches(4.0), Inches(4), Inches(0.5),
                     "M O D U L E", font_size=14, color=RGBColor(0xFF, 0xFF, 0xFF))
    else:
        # 导论页大图标
        add_text_box(slide, Inches(0.5), Inches(2.0), Inches(4.5), Inches(2.0),
                     "🎬", font_size=100, color=RGBColor(0xFF, 0xFF, 0xFF), alignment=PP_ALIGN.CENTER)

    # 右侧内容区
    # 课程标识
    add_rounded_rect(slide, Inches(6.2), Inches(1.5), Inches(6.5), Inches(0.5), fill_color=c)
    add_text_box(slide, Inches(6.4), Inches(1.5), Inches(6), Inches(0.5),
                 "AI漫剧制作全流程课程 · 2026版", font_size=12, color=COLORS['text_white'],
                 alignment=PP_ALIGN.CENTER)

    # 主标题
    add_text_box(slide, Inches(6.2), Inches(2.5), Inches(6.5), Inches(1.8),
                 title, font_size=36, color=COLORS['text_white'], bold=True)

    # 装饰线
    add_accent_line(slide, Inches(6.2), Inches(4.4), Inches(3), color=c)

    # 副标题
    add_text_box(slide, Inches(6.2), Inches(4.8), Inches(6.5), Inches(1.0),
                 subtitle, font_size=18, color=COLORS['text_light'])

    # 底部统计条
    add_shape(slide, Inches(6.2), Inches(6.2), Inches(6.5), Inches(0.7), fill_color=COLORS['bg_card'])
    stats = ["124课时", "17模块", "8-16周", "全流程"]
    for i, stat in enumerate(stats):
        x = Inches(6.4) + Inches(i * 1.6)
        add_text_box(slide, x, Inches(6.25), Inches(1.5), Inches(0.6),
                     stat, font_size=12, color=c, bold=True, alignment=PP_ALIGN.CENTER)


def make_toc_slide(prs, items, module_num, module_name, page_num):
    """目录页 - 网格卡片布局"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, "📋 本节内容", c)

    # 网格卡片（3列）
    cols = 3
    card_w = Inches(3.7)
    card_h = Inches(1.1)
    gap_x = Inches(0.35)
    gap_y = Inches(0.25)
    start_x = Inches(0.6)
    start_y = Inches(1.5)

    for i, item in enumerate(items):
        col = i % cols
        row = i // cols
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)

        if y > Inches(6.5):
            break

        # 卡片背景
        add_rounded_rect(slide, x, y, card_w, card_h, fill_color=COLORS['bg_card'], border_color=c)
        # 编号圆
        add_circle(slide, x + Inches(0.15), y + Inches(0.25), Inches(0.5), fill_color=c)
        add_text_box(slide, x + Inches(0.15), y + Inches(0.25), Inches(0.5), Inches(0.5),
                     str(i + 1), font_size=14, color=COLORS['text_white'], bold=True,
                     alignment=PP_ALIGN.CENTER)
        # 文字
        add_text_box(slide, x + Inches(0.8), y + Inches(0.2), Inches(2.7), Inches(0.7),
                     item, font_size=13, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_content_slide(prs, title, bullets, module_num, module_name, page_num, accent_color=None):
    """标准内容页 - 增强版"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = accent_color or get_module_color(module_num)

    add_page_header(slide, title, c)

    # 内容区 - 每个bullet一个卡片
    for i, bullet in enumerate(bullets):
        y = Inches(1.45) + Inches(i * 0.95)
        if y > Inches(6.5):
            break

        # 卡片
        add_rounded_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.8),
                         fill_color=COLORS['bg_card'])
        # 左侧色条
        add_shape(slide, Inches(0.5), y, Inches(0.08), Inches(0.8), fill_color=c)
        # 要点标记
        add_circle(slide, Inches(0.8), y + Inches(0.18), Inches(0.4), fill_color=c)
        add_text_box(slide, Inches(0.8), y + Inches(0.18), Inches(0.4), Inches(0.4),
                     str(i + 1), font_size=11, color=COLORS['text_white'], bold=True,
                     alignment=PP_ALIGN.CENTER)
        # 文字
        add_text_box(slide, Inches(1.4), y + Inches(0.1), Inches(11), Inches(0.6),
                     bullet, font_size=15, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_stat_cards_slide(prs, title, stats, module_num, module_name, page_num):
    """数据统计卡片页 - 大数字展示"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    cols = min(len(stats), 4)
    card_w = Inches(12.3 / cols - 0.3)
    start_x = Inches(0.5)
    card_y = Inches(1.6)
    card_h = Inches(5.2)

    stat_colors = [COLORS['accent'], COLORS['accent2'], COLORS['accent3'], COLORS['accent4'],
                   COLORS['accent5'], COLORS['accent6']]

    for i, stat in enumerate(stats[:cols]):
        x = start_x + i * (card_w + Inches(0.3))
        sc = stat_colors[i % len(stat_colors)]

        # 大卡片
        add_rounded_rect(slide, x, card_y, card_w, card_h, fill_color=COLORS['bg_card'])
        # 顶部色条
        add_shape(slide, x, card_y, card_w, Inches(0.08), fill_color=sc)

        # 大数字
        add_text_box(slide, x + Inches(0.2), card_y + Inches(0.5), card_w - Inches(0.4), Inches(1.5),
                     stat.get('value', ''), font_size=48, color=sc, bold=True,
                     alignment=PP_ALIGN.CENTER)

        # 单位
        if stat.get('unit'):
            add_text_box(slide, x + Inches(0.2), card_y + Inches(1.8), card_w - Inches(0.4), Inches(0.4),
                         stat['unit'], font_size=16, color=COLORS['text_dim'],
                         alignment=PP_ALIGN.CENTER)

        # 标签
        add_text_box(slide, x + Inches(0.2), card_y + Inches(2.3), card_w - Inches(0.4), Inches(0.5),
                     stat.get('label', ''), font_size=16, color=COLORS['text_white'], bold=True,
                     alignment=PP_ALIGN.CENTER)

        # 描述
        if stat.get('desc'):
            add_text_box(slide, x + Inches(0.3), card_y + Inches(3.0), card_w - Inches(0.6), Inches(1.5),
                         stat['desc'], font_size=12, color=COLORS['text_dim'],
                         alignment=PP_ALIGN.CENTER)

        # 趋势
        if stat.get('trend'):
            trend_color = COLORS['accent4'] if '↑' in stat['trend'] or '+' in stat['trend'] else COLORS['accent6']
            add_rounded_rect(slide, x + Inches(0.3), card_y + Inches(4.5), card_w - Inches(0.6), Inches(0.4),
                             fill_color=COLORS['bg_card2'])
            add_text_box(slide, x + Inches(0.3), card_y + Inches(4.5), card_w - Inches(0.6), Inches(0.4),
                         stat['trend'], font_size=13, color=trend_color, bold=True,
                         alignment=PP_ALIGN.CENTER)

    add_footer(slide, module_num, module_name, page_num)


def make_two_col_slide(prs, title, left_title, left_items, right_title, right_items,
                       module_num, module_name, page_num):
    """双栏对比页 - 增强版"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    # 左栏卡片
    add_rounded_rect(slide, Inches(0.4), Inches(1.5), Inches(6.0), Inches(5.3),
                     fill_color=COLORS['bg_card'])
    add_shape(slide, Inches(0.4), Inches(1.5), Inches(6.0), Inches(0.08), fill_color=c)
    add_text_box(slide, Inches(0.7), Inches(1.8), Inches(5.4), Inches(0.5),
                 left_title, font_size=20, color=c, bold=True)
    add_accent_line(slide, Inches(0.7), Inches(2.4), Inches(1.5), color=c)

    for i, item in enumerate(left_items):
        y = Inches(2.7) + Inches(i * 0.65)
        if y > Inches(6.4):
            break
        # 小圆点
        add_circle(slide, Inches(0.8), y + Inches(0.08), Inches(0.2), fill_color=c)
        add_text_box(slide, Inches(1.15), y, Inches(5), Inches(0.55),
                     item, font_size=14, color=COLORS['text_light'])

    # 右栏卡片
    add_rounded_rect(slide, Inches(6.8), Inches(1.5), Inches(6.0), Inches(5.3),
                     fill_color=COLORS['bg_card'])
    add_shape(slide, Inches(6.8), Inches(1.5), Inches(6.0), Inches(0.08), fill_color=COLORS['accent2'])
    add_text_box(slide, Inches(7.1), Inches(1.8), Inches(5.4), Inches(0.5),
                 right_title, font_size=20, color=COLORS['accent2'], bold=True)
    add_accent_line(slide, Inches(7.1), Inches(2.4), Inches(1.5), color=COLORS['accent2'])

    for i, item in enumerate(right_items):
        y = Inches(2.7) + Inches(i * 0.65)
        if y > Inches(6.4):
            break
        add_circle(slide, Inches(7.2), y + Inches(0.08), Inches(0.2), fill_color=COLORS['accent2'])
        add_text_box(slide, Inches(7.55), y, Inches(5), Inches(0.55),
                     item, font_size=14, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_table_slide(prs, title, headers, rows, module_num, module_name, page_num):
    """表格页 - 增强版"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    cols = len(headers)
    tbl_rows_count = len(rows) + 1
    table_left = Inches(0.6)
    table_top = Inches(1.6)
    table_w = Inches(12.1)
    row_h = Inches(0.65)
    table_h = row_h * tbl_rows_count

    table_shape = slide.shapes.add_table(tbl_rows_count, cols, table_left, table_top, table_w, table_h)
    table = table_shape.table

    col_w = table_w / cols
    for j in range(cols):
        table.columns[j].width = int(col_w)

    # 表头
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = c
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(14)
            p.font.color.rgb = COLORS['text_white']
            p.font.bold = True
            p.font.name = 'Microsoft YaHei'
            p.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 数据行
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLORS['bg_card'] if i % 2 == 0 else COLORS['bg_card2']
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(13)
                p.font.color.rgb = COLORS['text_light']
                p.font.name = 'Microsoft YaHei'
                p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    add_footer(slide, module_num, module_name, page_num)


def make_key_point_slide(prs, title, key_point, explanation, module_num, module_name, page_num):
    """重点强调页 - 视觉焦点"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    # 大号重点卡片 - 居中
    card_x = Inches(0.8)
    card_y = Inches(1.8)
    card_w = Inches(11.7)
    card_h = Inches(3.0)

    add_rounded_rect(slide, card_x, card_y, card_w, card_h, fill_color=COLORS['bg_card'])
    # 顶部渐变色条
    add_shape(slide, card_x, card_y, card_w, Inches(0.1), fill_color=c)
    # 左侧大色块
    add_shape(slide, card_x, card_y, Inches(0.12), card_h, fill_color=c)

    # 💡 图标
    add_text_box(slide, card_x + Inches(0.4), card_y + Inches(0.3), Inches(1), Inches(0.6),
                 "💡", font_size=32, color=c)
    add_text_box(slide, card_x + Inches(1.0), card_y + Inches(0.35), Inches(3), Inches(0.5),
                 "核心要点", font_size=16, color=c, bold=True)

    # 重点文字
    add_text_box(slide, card_x + Inches(0.5), card_y + Inches(1.0), card_w - Inches(1.0), Inches(1.5),
                 key_point, font_size=24, color=COLORS['text_white'], bold=True)

    # 说明区
    add_rounded_rect(slide, Inches(0.8), Inches(5.2), Inches(11.7), Inches(1.5),
                     fill_color=COLORS['bg_card2'])
    add_shape(slide, Inches(0.8), Inches(5.2), Inches(0.08), Inches(1.5), fill_color=COLORS['text_dim'])
    add_text_box(slide, Inches(1.2), Inches(5.35), Inches(11), Inches(1.2),
                 explanation, font_size=14, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_step_slide(prs, title, steps, module_num, module_name, page_num):
    """步骤流程页 - 带连接线"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    step_colors = [COLORS['accent'], COLORS['accent2'], COLORS['accent3'],
                   COLORS['accent4'], COLORS['accent5'], COLORS['accent6']]

    n = len(steps)
    if n <= 4:
        # 横排卡片
        card_w = Inches(12.3 / n - 0.25)
        card_h = Inches(5.0)
        start_x = Inches(0.5)
        card_y = Inches(1.5)

        for i, step in enumerate(steps):
            x = start_x + i * (card_w + Inches(0.3))
            sc = step_colors[i % len(step_colors)]

            add_rounded_rect(slide, x, card_y, card_w, card_h, fill_color=COLORS['bg_card'])
            add_shape(slide, x, card_y, card_w, Inches(0.08), fill_color=sc)

            # 编号圆
            add_circle(slide, x + card_w / 2 - Inches(0.3), card_y + Inches(0.4), Inches(0.6), fill_color=sc)
            add_text_box(slide, x + card_w / 2 - Inches(0.3), card_y + Inches(0.4), Inches(0.6), Inches(0.6),
                         str(i + 1), font_size=18, color=COLORS['text_white'], bold=True,
                         alignment=PP_ALIGN.CENTER)

            # 步骤文字
            add_text_box(slide, x + Inches(0.2), card_y + Inches(1.3), card_w - Inches(0.4), Inches(3.5),
                         step, font_size=14, color=COLORS['text_light'])

            # 连接箭头（非最后一步）
            if i < n - 1:
                arrow_x = x + card_w + Inches(0.02)
                add_text_box(slide, arrow_x, card_y + card_h / 2 - Inches(0.2), Inches(0.25), Inches(0.4),
                             "→", font_size=18, color=sc, bold=True)
    else:
        # 竖排步骤
        for i, step in enumerate(steps):
            y = Inches(1.5) + Inches(i * 0.9)
            if y > Inches(6.5):
                break
            sc = step_colors[i % len(step_colors)]

            # 卡片
            add_rounded_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.75),
                             fill_color=COLORS['bg_card'])
            add_shape(slide, Inches(0.5), y, Inches(0.08), Inches(0.75), fill_color=sc)

            # 编号
            add_circle(slide, Inches(0.8), y + Inches(0.12), Inches(0.45), fill_color=sc)
            add_text_box(slide, Inches(0.8), y + Inches(0.12), Inches(0.45), Inches(0.45),
                         str(i + 1), font_size=13, color=COLORS['text_white'], bold=True,
                         alignment=PP_ALIGN.CENTER)

            # 步骤文字
            add_text_box(slide, Inches(1.5), y + Inches(0.1), Inches(11), Inches(0.55),
                         step, font_size=15, color=COLORS['text_light'])

            # 连接线
            if i < len(steps) - 1 and y + Inches(0.9) <= Inches(6.5):
                add_shape(slide, Inches(1.0), y + Inches(0.75), Pt(2), Inches(0.15),
                          fill_color=sc)

    add_footer(slide, module_num, module_name, page_num)


def make_practice_slide(prs, title, tasks, module_num, module_name, page_num):
    """实操练习页 - 清单卡片"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = COLORS['accent4']  # 绿色

    add_page_header(slide, "🛠️ " + title, c)

    for i, task in enumerate(tasks):
        y = Inches(1.5) + Inches(i * 0.9)
        if y > Inches(6.5):
            break

        # 任务卡片
        add_rounded_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.75),
                         fill_color=COLORS['bg_card'])
        # checkbox
        add_rounded_rect(slide, Inches(0.7), y + Inches(0.15), Inches(0.4), Inches(0.4),
                         border_color=c)
        # 编号
        add_text_box(slide, Inches(1.3), y + Inches(0.1), Inches(0.5), Inches(0.5),
                     f"#{i+1}", font_size=12, color=c, bold=True)
        # 文字
        add_text_box(slide, Inches(1.8), y + Inches(0.1), Inches(10.5), Inches(0.55),
                     task, font_size=15, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_summary_slide(prs, takeaways, module_num, module_name, page_num):
    """总结页 - 渐进编号"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, "📝 本节小结", c)

    for i, item in enumerate(takeaways):
        y = Inches(1.5) + Inches(i * 0.85)
        if y > Inches(6.5):
            break

        # 渐进色编号
        intensity = max(0.5, 1.0 - i * 0.1)
        rc = RGBColor(
            min(255, int(c.red * intensity)),
            min(255, int(c.green * intensity)),
            min(255, int(c.blue * intensity))
        )

        # 卡片
        add_rounded_rect(slide, Inches(0.5), y, Inches(12.3), Inches(0.7),
                         fill_color=COLORS['bg_card'])
        # 编号
        add_circle(slide, Inches(0.7), y + Inches(0.1), Inches(0.45), fill_color=rc)
        add_text_box(slide, Inches(0.7), y + Inches(0.1), Inches(0.45), Inches(0.45),
                     str(i + 1), font_size=13, color=COLORS['text_white'], bold=True,
                     alignment=PP_ALIGN.CENTER)
        # 文字
        add_text_box(slide, Inches(1.4), y + Inches(0.08), Inches(11), Inches(0.55),
                     item, font_size=15, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_section_slide(prs, section_title, subtitle, module_num, module_name, page_num):
    """章节分隔页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    # 左侧色块
    add_shape(slide, Inches(0), Inches(0), Inches(0.15), SLIDE_HEIGHT, fill_color=c)
    # 装饰线
    add_accent_line(slide, Inches(1), Inches(3.0), Inches(4), color=c)
    # 标题
    add_text_box(slide, Inches(1), Inches(3.3), Inches(10), Inches(1.0),
                 section_title, font_size=36, color=COLORS['text_white'], bold=True)
    # 副标题
    add_text_box(slide, Inches(1), Inches(4.5), Inches(10), Inches(0.8),
                 subtitle, font_size=18, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_end_slide(prs, next_module_title, module_num, module_name):
    """结束页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    # 中央大圆
    add_circle(slide, Inches(5.5), Inches(1.5), Inches(2.3), fill_color=c)
    add_text_box(slide, Inches(5.5), Inches(1.8), Inches(2.3), Inches(1.8),
                 "✅", font_size=60, color=COLORS['text_white'], alignment=PP_ALIGN.CENTER)

    # 完成文字
    add_text_box(slide, Inches(0), Inches(4.2), SLIDE_WIDTH, Inches(0.8),
                 "本模块结束", font_size=32, color=COLORS['text_white'], bold=True,
                 alignment=PP_ALIGN.CENTER)

    # 分割线
    add_accent_line(slide, Inches(5.5), Inches(5.2), Inches(2.3), color=c)

    # 下一模块
    if next_module_title:
        add_text_box(slide, Inches(0), Inches(5.5), SLIDE_WIDTH, Inches(0.5),
                     f"下一模块：{next_module_title}", font_size=18, color=COLORS['text_light'],
                     alignment=PP_ALIGN.CENTER)

    add_text_box(slide, Inches(0), Inches(6.5), SLIDE_WIDTH, Inches(0.4),
                 "AI漫剧制作全流程课程 · 2026版", font_size=12, color=COLORS['text_dim'],
                 alignment=PP_ALIGN.CENTER)


def make_icon_grid_slide(prs, title, items, module_num, module_name, page_num):
    """图标网格页 - 适合展示工具/特性列表"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    cols = 3
    card_w = Inches(3.7)
    card_h = Inches(2.2)
    gap_x = Inches(0.35)
    gap_y = Inches(0.3)
    start_x = Inches(0.6)
    start_y = Inches(1.5)

    for i, item in enumerate(items[:9]):
        col = i % cols
        row = i // cols
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)

        sc = [COLORS['accent'], COLORS['accent2'], COLORS['accent3'],
              COLORS['accent4'], COLORS['accent5'], COLORS['accent6']][i % 6]

        # 卡片
        add_rounded_rect(slide, x, y, card_w, card_h, fill_color=COLORS['bg_card'])
        add_shape(slide, x, y, card_w, Inches(0.06), fill_color=sc)

        # 图标/emoji
        icon = item.get('icon', '📌') if isinstance(item, dict) else '📌'
        add_text_box(slide, x + Inches(0.2), y + Inches(0.2), Inches(0.6), Inches(0.6),
                     icon, font_size=28, color=sc)

        # 标题
        item_title = item.get('title', '') if isinstance(item, dict) else item
        add_text_box(slide, x + Inches(0.2), y + Inches(0.8), card_w - Inches(0.4), Inches(0.4),
                     item_title, font_size=14, color=COLORS['text_white'], bold=True)

        # 描述
        if isinstance(item, dict) and item.get('desc'):
            add_text_box(slide, x + Inches(0.2), y + Inches(1.3), card_w - Inches(0.4), Inches(0.8),
                         item['desc'], font_size=11, color=COLORS['text_dim'])

    add_footer(slide, module_num, module_name, page_num)


def make_timeline_slide(prs, title, events, module_num, module_name, page_num):
    """时间线页 - 横向时间轴"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    n = len(events)
    if n == 0:
        add_footer(slide, module_num, module_name, page_num)
        return

    # 横轴
    line_y = Inches(3.5)
    line_x = Inches(0.8)
    line_w = Inches(11.7)
    add_shape(slide, line_x, line_y, line_w, Pt(3), fill_color=COLORS['border'])

    step_colors = [COLORS['accent'], COLORS['accent2'], COLORS['accent3'],
                   COLORS['accent4'], COLORS['accent5'], COLORS['accent6']]

    for i, ev in enumerate(events):
        x = line_x + i * (line_w / n)
        sc = step_colors[i % len(step_colors)]

        # 节点圆
        add_circle(slide, x + Inches(0.3), line_y - Inches(0.15), Inches(0.3), fill_color=sc)

        # 年份/标签
        label = ev.get('label', '') if isinstance(ev, dict) else str(ev)
        add_text_box(slide, x, line_y - Inches(0.8), Inches(1.5), Inches(0.5),
                     label, font_size=14, color=sc, bold=True, alignment=PP_ALIGN.CENTER)

        # 描述（交替上下）
        desc = ev.get('desc', '') if isinstance(ev, dict) else ''
        if desc:
            if i % 2 == 0:
                # 上方
                add_rounded_rect(slide, x - Inches(0.2), Inches(1.6), Inches(1.9), Inches(1.2),
                                 fill_color=COLORS['bg_card'])
                add_text_box(slide, x - Inches(0.1), Inches(1.7), Inches(1.7), Inches(1.0),
                             desc, font_size=11, color=COLORS['text_light'])
            else:
                # 下方
                add_rounded_rect(slide, x - Inches(0.2), Inches(4.0), Inches(1.9), Inches(1.2),
                                 fill_color=COLORS['bg_card'])
                add_text_box(slide, x - Inches(0.1), Inches(4.1), Inches(1.7), Inches(1.0),
                             desc, font_size=11, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_comparison_matrix_slide(prs, title, headers, rows, module_num, module_name, page_num):
    """对比矩阵页 - 带评分视觉"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])
    c = get_module_color(module_num)

    add_page_header(slide, title, c)

    cols = len(headers)
    n_rows = len(rows) + 1
    table_left = Inches(0.5)
    table_top = Inches(1.5)
    table_w = Inches(12.3)
    row_h = Inches(0.7)

    table_shape = slide.shapes.add_table(n_rows, cols, table_left, table_top, table_w, row_h * n_rows)
    table = table_shape.table

    col_w = table_w / cols
    for j in range(cols):
        table.columns[j].width = int(col_w)

    # 表头
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLORS['bg_card']
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(13)
            p.font.color.rgb = c
            p.font.bold = True
            p.font.name = 'Microsoft YaHei'
            p.alignment = PP_ALIGN.CENTER

    # 数据行
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLORS['bg_card'] if i % 2 == 0 else COLORS['bg_card2']
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.color.rgb = COLORS['text_light']
                p.font.name = 'Microsoft YaHei'
                p.alignment = PP_ALIGN.CENTER

    add_footer(slide, module_num, module_name, page_num)


# ══════════════════════════════════════════════
# PR 设置
# ══════════════════════════════════════════════

def make_pr(prs):
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs


# ══════════════════════════════════════════════
# 模块内容定义
# ══════════════════════════════════════════════

MODULES = [
    {
        'num': 0,
        'name': '课程导论',
        'title': 'AI漫剧制作全流程课程',
        'subtitle': '124课时 · 17个模块 · 从零基础到独立制作完整AI漫剧',
        'next': '行业认知与趋势',
        'pages': [
            {'type': 'toc', 'items': [
                '行业认知与趋势', '大语言模型与剧本创作', '分镜设计与提示词工程',
                '即梦AI深度精通', '其他图像生成工具', '可灵AI深度精通',
                'Seedance 2.0深度精通', '海螺AI与Vidu', '角色一致性攻克',
                '配音配乐与音画同步', '剪辑与后期全流程', '一站式平台精讲',
                'ComfyUI工作流', '工业化生产与团队协作', '平台分发与商业变现',
                '合规要求与行业规范', '综合实战项目'
            ]},
            {'type': 'content', 'title': '课程目标', 'bullets': [
                '掌握AI漫剧从剧本到成片的完整制作流程',
                '精通即梦AI、可灵AI、Seedance 2.0等核心工具',
                '能独立制作一集完整的AI漫剧',
                '了解平台分发与商业变现的路径',
                '建立工业化生产思维，实现高效量产'
            ]},
            {'type': 'stat_cards', 'title': '📊 行业数据一览', 'stats': [
                {'value': '240', 'unit': '亿元', 'label': '市场规模', 'trend': '↑ +50% YoY'},
                {'value': '2.8', 'unit': '亿', 'label': '用户规模', 'trend': '↑ +40% YoY'},
                {'value': '15', 'unit': '亿次/天', 'label': '日均播放', 'trend': '↑ +50% YoY'},
                {'value': '120', 'unit': '万', 'label': '创作者', 'trend': '↑ +50% YoY'},
            ]},
            {'type': 'two_col', 'title': '学习路径选择',
             'left_title': '🅰️ 零基础速成（8周）', 'left_items': ['快速上手出作品', '重点模块：1,2,3,4,6,10,11,12,17', '适合：想快速入行变现'],
             'right_title': '🅱️ 专业进阶（16周）', 'right_items': ['全面掌握全链路', '全部17个模块', '适合：追求专业深度']},
            {'type': 'step', 'title': '课程学习路线图', 'steps': [
                'Step 1：行业认知 → 理解AI漫剧是什么、市场多大',
                'Step 2：创作技能 → 剧本、分镜、提示词工程',
                'Step 3：工具精通 → 即梦AI、可灵AI、Seedance等',
                'Step 4：后期制作 → 音频、剪辑、调色、成片',
                'Step 5：商业运营 → 平台分发、变现、合规',
                'Step 6：综合实战 → 从0到1独立完成一集AI漫剧'
            ]},
            {'type': 'key_point', 'title': '核心竞争力',
             'key_point': 'AI漫剧的核心竞争力 = 创意 × 工具熟练度 × 量产能力',
             'explanation': '工具会更新迭代，但创意能力和工业化思维是持久的竞争力。本课程不仅教工具操作，更注重培养你的创作思维和生产效率。'},
        ]
    },
    {
        'num': 1,
        'name': '行业认知与趋势',
        'title': '模块一：AI漫剧行业认知与趋势',
        'subtitle': '6课时 · 从行业全景到创业路径',
        'next': '大语言模型与剧本创作',
        'pages': [
            {'type': 'toc', 'items': ['行业全景：从手工作坊到智能流水线', '市场数据：240亿规模与2.8亿用户', '内容形态：沙雕漫→动态漫→AI原生漫剧', '五步工业化流程', '爆款拆解：《斩仙台》等', '创业路径：个人vs团队vs工作室']},
            {'type': 'content', 'title': '什么是AI漫剧？', 'bullets': [
                '利用AI技术辅助或主导制作的动态漫画短剧',
                '制作周期：传统1-3个月/集 → AI 1-3天/集',
                '制作成本：传统5-20万/集 → AI 500-5000元/集',
                '团队规模：传统10-30人 → AI 1-5人',
                '产能：传统2-4集/月 → AI 30+集/月'
            ]},
            {'type': 'timeline', 'title': '行业发展四个阶段', 'events': [
                {'label': '2023', 'desc': '沙雕漫：静态图片+配音+字幕'},
                {'label': '2024', 'desc': '动态漫：图生视频初步应用'},
                {'label': '2025', 'desc': 'AI原生漫剧：全链路AI辅助'},
                {'label': '2026', 'desc': 'AI仿真人剧：接近真人质量'},
            ]},
            {'type': 'stat_cards', 'title': '2026年核心市场数据', 'stats': [
                {'value': '240', 'unit': '亿元', 'label': '市场规模', 'trend': '↑ +50%', 'desc': '较2025年160亿'},
                {'value': '2.8', 'unit': '亿', 'label': '用户规模', 'trend': '↑ +40%', 'desc': '较2025年2.0亿'},
                {'value': '15', 'unit': '亿次', 'label': '日均播放', 'trend': '↑ +50%', 'desc': '较2025年10亿次'},
                {'value': '120', 'unit': '万', 'label': '创作者', 'trend': '↑ +50%', 'desc': '较2025年80万'},
            ]},
            {'type': 'step', 'title': '五步工业化流程', 'steps': [
                'Step 1 · 剧本创作：用AI生成分集剧本',
                'Step 2 · 分镜设计：将剧本转化为分镜表',
                'Step 3 · 文生图：用即梦AI生成关键帧画面',
                'Step 4 · 图生视频：用可灵AI/Seedance让画面动起来',
                'Step 5 · 剪辑配音：用剪映完成后期制作'
            ]},
            {'type': 'content', 'title': '爆款案例拆解', 'bullets': [
                '《斩仙台》：玄幻题材，角色一致性标杆，单集播放破千万',
                '《气运三角洲》：都市逆袭，节奏把控精准',
                '《霍去病》：历史题材，画面质感接近电影级别',
                '共同特点：前3秒强钩子、角色一致性好、节奏紧凑'
            ]},
            {'type': 'two_col', 'title': '创业路径选择',
             'left_title': '👤 个人创作者', 'left_items': ['月入1-10万', '日产1集', '月成本~5000元', '适合：副业/自由职业'],
             'right_title': '👥 小团队（3-5人）', 'right_items': ['月入10-50万', '日产4集', '月成本~7万', '适合：创业团队']},
            {'type': 'practice', 'title': '本模块实操任务', 'tasks': [
                '关注3个AI漫剧行业公众号/博主，建立信息源',
                '在抖音搜索10个AI漫剧账号，分析其内容特点',
                '画出完整的AI漫剧产业链地图',
                '确定自己的定位：个人/团队/工作室',
                '选择一个感兴趣的题材方向'
            ]},
        ]
    },
]


# ══════════════════════════════════════════════
# 生成函数
# ══════════════════════════════════════════════

def generate_module_ppt(module_data):
    """生成单个模块的PPT"""
    prs = Presentation()
    make_pr(prs)

    num = module_data['num']
    name = module_data['name']
    title = module_data['title']
    subtitle = module_data['subtitle']
    next_mod = module_data.get('next', '')
    pages = module_data['pages']

    # 封面页
    make_title_slide(prs, title, subtitle, num)

    # 内容页
    page_num = 1
    for page in pages:
        ptype = page['type']
        if ptype == 'toc':
            make_toc_slide(prs, page['items'], num, name, page_num)
        elif ptype == 'content':
            make_content_slide(prs, page['title'], page['bullets'], num, name, page_num)
        elif ptype == 'two_col':
            make_two_col_slide(prs, page['title'], page['left_title'], page['left_items'],
                               page['right_title'], page['right_items'], num, name, page_num)
        elif ptype == 'table':
            make_table_slide(prs, page['title'], page['headers'], page['rows'], num, name, page_num)
        elif ptype == 'stat_cards':
            make_stat_cards_slide(prs, page['title'], page['stats'], num, name, page_num)
        elif ptype == 'icon_grid':
            make_icon_grid_slide(prs, page['title'], page['items'], num, name, page_num)
        elif ptype == 'timeline':
            make_timeline_slide(prs, page['title'], page['events'], num, name, page_num)
        elif ptype == 'comparison':
            make_comparison_matrix_slide(prs, page['title'], page['headers'], page['rows'], num, name, page_num)
        elif ptype == 'key_point':
            make_key_point_slide(prs, page.get('title', '重点回顾'),
                                 page['key_point'], page['explanation'], num, name, page_num)
        elif ptype == 'step':
            make_step_slide(prs, page['title'], page['steps'], num, name, page_num)
        elif ptype == 'practice':
            make_practice_slide(prs, page['title'], page['tasks'], num, name, page_num)
        elif ptype == 'section':
            make_section_slide(prs, page['title'], page['subtitle'], num, name, page_num)
        elif ptype == 'summary':
            make_summary_slide(prs, page['items'], num, name, page_num)
        page_num += 1

    # 结束页
    make_end_slide(prs, next_mod, num, name)

    # 保存
    filename = f"模块{num:02d}-{name}.pptx" if num > 0 else "00-课程导论.pptx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    prs.save(filepath)
    print(f"✅ 已生成：{filename}（{len(prs.slides)}页）")
    return filepath


if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for m in MODULES:
        generate_module_ppt(m)
    print(f"\n🎉 已生成 {len(MODULES)} 个PPT文件")
