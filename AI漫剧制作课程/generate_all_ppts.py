#!/usr/bin/env python3
"""
AI漫剧制作课程 PPT 生成器 v4
基于 CRAP 设计原则 + 现代高级审美
- Contrast: 强对比层次感
- Repetition: 重复视觉元素统一风格
- Alignment: 网格对齐系统
- Proximity: 亲密性分组
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ══════════════════════════════════════════════
# 设计系统 Design System
# ══════════════════════════════════════════════

# 60-30-10 配色法则
# 60% 主色（深色背景）  30% 辅色（卡片/文字）  10% 强调色（点缀）
C = {
    # 60% - 背景层
    'bg':         RGBColor(0x0A, 0x0E, 0x1A),   # 极深蓝黑
    'bg_alt':     RGBColor(0x10, 0x17, 0x29),   # 微亮背景

    # 30% - 中间层
    'surface':    RGBColor(0x16, 0x1F, 0x35),   # 卡片/区块
    'surface2':   RGBColor(0x1C, 0x28, 0x42),   # 次级卡片
    'border':     RGBColor(0x25, 0x33, 0x50),   # 边框/分割线

    # 文字层级（对比原则：4级层次）
    'text_primary':   RGBColor(0xF0, 0xF4, 0xFC),  # 主文字 - 最亮白
    'text_secondary': RGBColor(0xA0, 0xAE, 0xC4),  # 次文字
    'text_tertiary':  RGBColor(0x60, 0x70, 0x90),  # 辅助文字
    'text_disabled':  RGBColor(0x3A, 0x48, 0x60),  # 装饰文字

    # 10% - 强调色（每个模块不同，这里定义调色板）
    'cyan':       RGBColor(0x00, 0xC8, 0xF0),   # 科技青
    'violet':     RGBColor(0x8B, 0x5C, 0xF6),   # 紫罗兰
    'orange':     RGBColor(0xF0, 0x7A, 0x35),   # 活力橙
    'emerald':    RGBColor(0x10, 0xB9, 0x81),   # 翡翠绿
    'amber':      RGBColor(0xF5, 0x9E, 0x0B),   # 琥珀黄
    'rose':       RGBColor(0xF4, 0x3F, 0x5E),   # 玫瑰红
}

# 模块强调色映射
MODULE_ACCENTS = [
    C['cyan'], C['orange'], C['cyan'], C['violet'],
    C['orange'], C['emerald'], C['cyan'], C['violet'],
    C['amber'], C['rose'], C['emerald'], C['orange'],
    C['cyan'], C['violet'], C['amber'], C['emerald'],
    C['rose'], C['orange'],
]

# 排版系统（对比原则：字号梯度）
FONT = {
    'display':  48,   # 展示级 - 封面大标题
    'h1':       32,   # 一级标题
    'h2':       24,   # 二级标题
    'h3':       18,   # 三级标题
    'body':     14,   # 正文
    'caption':  11,   # 辅助说明
    'micro':    9,    # 极小标注
}

# 网格系统（Alignment原则：统一边距）
MARGIN_L = Inches(0.8)
MARGIN_R = Inches(0.8)
MARGIN_T = Inches(0.6)
CONTENT_W = Inches(11.733)  # 13.333 - 0.8*2
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

OUTPUT_DIR = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程/PPT课件'

# ══════════════════════════════════════════════
# 基础图形工具
# ══════════════════════════════════════════════

def _acc(num):
    """获取模块强调色"""
    if 0 <= num < len(MODULE_ACCENTS):
        return MODULE_ACCENTS[num]
    return C['cyan']


def set_bg(slide, color=None):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color or C['bg']


def rect(slide, x, y, w, h, fill=None, border=None, border_w=Pt(1)):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.line.fill.background()
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if border:
        s.line.color.rgb = border
        s.line.width = border_w
    return s


def rrect(slide, x, y, w, h, fill=None, border=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.line.fill.background()
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(1)
    return s


def circle(slide, x, y, size, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    s.line.fill.background()
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    return s


def txt(slide, x, y, w, h, text, size=14, color=None, bold=False, align=PP_ALIGN.LEFT, font='Microsoft YaHei'):
    """单行文本框"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(size)
    p.font.color.rgb = color or C['text_primary']
    p.font.bold = bold
    p.font.name = font
    p.alignment = align
    return tb


def multiline(slide, x, y, w, h, lines, size=14, color=None, spacing=1.3):
    """多行文本（Proximity原则：紧凑分组）"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"▸ {line}" if line.strip() else line
        p.font.size = Pt(size)
        p.font.color.rgb = color or C['text_secondary']
        p.font.name = 'Microsoft YaHei'
        p.space_after = Pt(size * 0.4)
    return tb


def line(slide, x, y, w, color=None, thickness=Pt(2)):
    return rect(slide, x, y, w, thickness, fill=color or C['border'])


# ══════════════════════════════════════════════
# 页面结构组件
# ══════════════════════════════════════════════

def page_footer(slide, mod_num, mod_name, page_num, accent):
    """统一页脚（Repetition原则：每页重复）"""
    line(slide, Inches(0), Inches(7.05), SLIDE_W, C['border'], Pt(1))
    txt(slide, Inches(0.8), Inches(7.1), Inches(5), Inches(0.3),
        f"模块{mod_num:02d} · {mod_name}", FONT['micro'], C['text_tertiary'])
    txt(slide, Inches(11.5), Inches(7.1), Inches(1.5), Inches(0.3),
        str(page_num), FONT['micro'], C['text_tertiary'], align=PP_ALIGN.RIGHT)


def page_header(slide, title, accent, icon=""):
    """统一页面头部（Repetition原则）"""
    # 顶部强调线
    rect(slide, Inches(0), Inches(0), SLIDE_W, Pt(3), fill=accent)
    # 标题区背景
    rect(slide, Inches(0), Pt(3), SLIDE_W, Inches(1.0), fill=C['surface'])
    # 左侧强调块
    rect(slide, Inches(0), Pt(3), Inches(0.1), Inches(1.0), fill=accent)
    # 标题文字
    label = f"{icon} {title}" if icon else title
    txt(slide, MARGIN_L, Inches(0.15), CONTENT_W, Inches(0.7),
        label, FONT['h2'], C['text_primary'], bold=True)
    # 底部分割线
    line(slide, Inches(0), Inches(1.18), SLIDE_W, C['border'], Pt(1))


# ══════════════════════════════════════════════
# 幻灯片类型
# ══════════════════════════════════════════════

def make_cover(prs, title, subtitle, mod_num=0):
    """封面页 - 大留白 + 大对比 + 几何装饰"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    # ── 左侧色块（60-30-10 中的 10%）──
    rect(slide, Inches(0), Inches(0), Inches(5.2), SLIDE_H, fill=accent)

    # 左上角极细白线装饰（Alignment原则）
    rect(slide, Inches(0), Inches(0), Inches(5.2), Pt(2), fill=C['text_primary'])
    rect(slide, Inches(0), Inches(0), Pt(2), SLIDE_H, fill=C['text_primary'])

    # 大号模块编号（Contrast原则：超大字号形成视觉锚点）
    if mod_num > 0:
        txt(slide, Inches(0.6), Inches(1.2), Inches(4.0), Inches(2.8),
            f"{mod_num:02d}", 120, C['text_primary'], bold=True)
        # MODULE 标签（Repetition: 固定标签格式）
        txt(slide, Inches(0.7), Inches(3.8), Inches(3), Inches(0.4),
            "M O D U L E", 12, C['text_primary'])
        # 细装饰线
        rect(slide, Inches(0.7), Inches(4.3), Inches(1.5), Pt(1.5), fill=C['text_primary'])
    else:
        txt(slide, Inches(0.6), Inches(1.8), Inches(4.0), Inches(2.5),
            "🎬", 80, C['text_primary'], align=PP_ALIGN.CENTER)

    # ── 右侧内容区（大量留白）──
    # 课程标签
    rrect(slide, Inches(6.0), Inches(1.5), Inches(6.5), Inches(0.45), fill=accent)
    txt(slide, Inches(6.2), Inches(1.5), Inches(6.1), Inches(0.45),
        "AI漫剧制作全流程课程 · 2026版", 11, C['text_primary'], align=PP_ALIGN.CENTER)

    # 主标题（Contrast: 大标题 vs 小副标题）
    txt(slide, Inches(6.0), Inches(2.6), Inches(6.5), Inches(1.6),
        title, FONT['display'] - 4, C['text_primary'], bold=True)

    # 装饰线（Proximity: 紧贴标题下方）
    line(slide, Inches(6.0), Inches(4.4), Inches(2.5), accent, Pt(3))

    # 副标题
    txt(slide, Inches(6.0), Inches(4.8), Inches(6.5), Inches(0.8),
        subtitle, FONT['body'] + 2, C['text_secondary'])

    # 底部统计条（Repetition: 固定格式的信息条）
    rect(slide, Inches(6.0), Inches(6.0), Inches(6.5), Inches(0.65), fill=C['surface'])
    stats = ["124课时", "17模块", "8-16周", "全流程"]
    for i, s in enumerate(stats):
        x = Inches(6.2) + Inches(i * 1.6)
        txt(slide, x, Inches(6.05), Inches(1.4), Inches(0.55),
            s, FONT['caption'], accent, bold=True, align=PP_ALIGN.CENTER)


def make_toc(prs, items, mod_num, mod_name, page_num):
    """目录页 - 网格卡片（Alignment: 网格对齐）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, "本节内容", accent, "📋")

    # 3列网格（Proximity: 相关项分组）
    cols = 3
    card_w = Inches(3.7)
    card_h = Inches(1.05)
    gap = Inches(0.3)
    sx = Inches(0.65)
    sy = Inches(1.5)

    for i, item in enumerate(items):
        col, row = i % cols, i // cols
        x = sx + col * (card_w + gap)
        y = sy + row * (card_h + gap)
        if y > Inches(6.5):
            break

        # 卡片（Contrast: 卡片与背景对比）
        rrect(slide, x, y, card_w, card_h, fill=C['surface'], border=C['border'])
        # 编号
        circle(slide, x + Inches(0.12), y + Inches(0.22), Inches(0.45), fill=accent)
        txt(slide, x + Inches(0.12), y + Inches(0.22), Inches(0.45), Inches(0.45),
            str(i + 1), 12, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)
        # 文字
        txt(slide, x + Inches(0.7), y + Inches(0.18), Inches(2.8), Inches(0.65),
            item, FONT['caption'] + 1, C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_content(prs, title, bullets, mod_num, mod_name, page_num):
    """内容页 - 逐条卡片（Contrast + Proximity）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent)

    n = len(bullets)
    # 动态计算间距（Alignment: 均匀分布）
    max_y = Inches(6.6)
    available = max_y - Inches(1.4)
    card_h = min(Inches(0.75), available / max(n, 1) - Inches(0.1))
    gap = (available - card_h * n) / max(n - 1, 1) if n > 1 else Inches(0)

    for i, bullet in enumerate(bullets):
        y = Inches(1.4) + i * (card_h + gap)
        if y + card_h > max_y:
            break

        # 卡片
        rrect(slide, Inches(0.5), y, Inches(12.3), card_h, fill=C['surface'])
        # 左侧色条（Repetition: 每张卡片重复）
        rect(slide, Inches(0.5), y, Inches(0.06), card_h, fill=accent)
        # 编号
        circle(slide, Inches(0.75), y + (card_h - Inches(0.35)) / 2, Inches(0.35), fill=accent)
        txt(slide, Inches(0.75), y + (card_h - Inches(0.35)) / 2, Inches(0.35), Inches(0.35),
            str(i + 1), 10, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)
        # 文字（Proximity: 文字紧邻编号）
        txt(slide, Inches(1.3), y + (card_h - Inches(0.4)) / 2, Inches(11), Inches(0.4),
            bullet, FONT['body'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_stat_cards(prs, title, stats, mod_num, mod_name, page_num):
    """数据卡片页 - 大数字（Contrast原则的极致体现）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent, "📊")

    n = min(len(stats), 4)
    card_w = Inches(12.3 / n - 0.2)
    gap = Inches(0.25)
    sx = Inches(0.5)
    card_y = Inches(1.5)
    card_h = Inches(5.2)

    palette = [C['cyan'], C['violet'], C['orange'], C['emerald'], C['amber'], C['rose']]

    for i, stat in enumerate(stats[:n]):
        x = sx + i * (card_w + gap)
        sc = palette[i % len(palette)]

        # 卡片
        rrect(slide, x, card_y, card_w, card_h, fill=C['surface'])
        # 顶部色线（Repetition）
        rect(slide, x, card_y, card_w, Pt(3), fill=sc)

        # 大数字（Contrast: 48pt数字 vs 11pt标签）
        txt(slide, x + Inches(0.15), card_y + Inches(0.6), card_w - Inches(0.3), Inches(1.2),
            stat.get('value', ''), 52, sc, bold=True, align=PP_ALIGN.CENTER)

        # 单位
        if stat.get('unit'):
            txt(slide, x + Inches(0.15), card_y + Inches(1.7), card_w - Inches(0.3), Inches(0.35),
                stat['unit'], FONT['body'], C['text_tertiary'], align=PP_ALIGN.CENTER)

        # 分割线
        line(slide, x + Inches(0.4), card_y + Inches(2.2), card_w - Inches(0.8), C['border'], Pt(1))

        # 标签
        txt(slide, x + Inches(0.15), card_y + Inches(2.5), card_w - Inches(0.3), Inches(0.4),
            stat.get('label', ''), FONT['body'], C['text_primary'], bold=True, align=PP_ALIGN.CENTER)

        # 描述
        if stat.get('desc'):
            txt(slide, x + Inches(0.2), card_y + Inches(3.1), card_w - Inches(0.4), Inches(1.2),
                stat['desc'], FONT['caption'], C['text_tertiary'], align=PP_ALIGN.CENTER)

        # 趋势标签
        if stat.get('trend'):
            is_up = any(c in stat['trend'] for c in ['↑', '+'])
            tc = C['emerald'] if is_up else C['rose']
            rrect(slide, x + Inches(0.3), card_y + Inches(4.4), card_w - Inches(0.6), Inches(0.35),
                  fill=C['bg_alt'])
            txt(slide, x + Inches(0.3), card_y + Inches(4.4), card_w - Inches(0.6), Inches(0.35),
                stat['trend'], FONT['caption'], tc, bold=True, align=PP_ALIGN.CENTER)

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_two_col(prs, title, left_title, left_items, right_title, right_items,
                 mod_num, mod_name, page_num):
    """双栏对比页（Contrast: 左右对比）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent)

    col_w = Inches(5.85)
    col_h = Inches(5.2)
    col_y = Inches(1.5)

    # 左栏
    rrect(slide, Inches(0.5), col_y, col_w, col_h, fill=C['surface'])
    rect(slide, Inches(0.5), col_y, col_w, Pt(3), fill=accent)
    txt(slide, Inches(0.8), col_y + Inches(0.25), Inches(5.2), Inches(0.45),
        left_title, FONT['h3'], accent, bold=True)
    line(slide, Inches(0.8), col_y + Inches(0.8), Inches(1.2), accent, Pt(2))

    for i, item in enumerate(left_items):
        iy = col_y + Inches(1.1) + i * Inches(0.55)
        if iy > col_y + col_h - Inches(0.3):
            break
        circle(slide, Inches(0.8), iy + Inches(0.06), Inches(0.18), fill=accent)
        txt(slide, Inches(1.15), iy, Inches(5.0), Inches(0.45),
            item, FONT['body'], C['text_secondary'])

    # 右栏
    rrect(slide, Inches(6.85), col_y, col_w, col_h, fill=C['surface'])
    rect(slide, Inches(6.85), col_y, col_w, Pt(3), fill=C['violet'])
    txt(slide, Inches(7.15), col_y + Inches(0.25), Inches(5.2), Inches(0.45),
        right_title, FONT['h3'], C['violet'], bold=True)
    line(slide, Inches(7.15), col_y + Inches(0.8), Inches(1.2), C['violet'], Pt(2))

    for i, item in enumerate(right_items):
        iy = col_y + Inches(1.1) + i * Inches(0.55)
        if iy > col_y + col_h - Inches(0.3):
            break
        circle(slide, Inches(7.15), iy + Inches(0.06), Inches(0.18), fill=C['violet'])
        txt(slide, Inches(7.5), iy, Inches(5.0), Inches(0.45),
            item, FONT['body'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_table(prs, title, headers, rows, mod_num, mod_name, page_num):
    """表格页（Alignment: 网格对齐）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent)

    cols = len(headers)
    n_rows = len(rows) + 1
    tbl_left = Inches(0.6)
    tbl_top = Inches(1.5)
    tbl_w = Inches(12.1)
    row_h = Inches(0.6)

    ts = slide.shapes.add_table(n_rows, cols, tbl_left, tbl_top, tbl_w, row_h * n_rows)
    tbl = ts.table
    col_w = tbl_w / cols
    for j in range(cols):
        tbl.columns[j].width = int(col_w)

    # 表头
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = accent
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(13)
            p.font.color.rgb = C['text_primary']
            p.font.bold = True
            p.font.name = 'Microsoft YaHei'
            p.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i + 1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C['surface'] if i % 2 == 0 else C['surface2']
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.color.rgb = C['text_secondary']
                p.font.name = 'Microsoft YaHei'
                p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_key_point(prs, title, key_point, explanation, mod_num, mod_name, page_num):
    """重点强调页（Contrast: 视觉焦点）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title or "核心要点", accent, "💡")

    # 大号重点卡片
    cx, cy = Inches(0.8), Inches(1.6)
    cw, ch = Inches(11.7), Inches(3.2)

    rrect(slide, cx, cy, cw, ch, fill=C['surface'])
    # 顶部+左侧双色条（Contrast: 强调）
    rect(slide, cx, cy, cw, Pt(4), fill=accent)
    rect(slide, cx, cy, Pt(5), ch, fill=accent)

    # 核心要点文字（大字号，高对比）
    txt(slide, cx + Inches(0.5), cy + Inches(0.5), cw - Inches(1.0), Inches(2.2),
        key_point, FONT['h1'], C['text_primary'], bold=True)

    # 说明区（Proximity: 紧贴重点卡片下方）
    rrect(slide, Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.6), fill=C['surface2'])
    rect(slide, Inches(0.8), Inches(5.1), Pt(4), Inches(1.6), fill=C['text_tertiary'])
    txt(slide, Inches(1.3), Inches(5.25), Inches(10.8), Inches(1.3),
        explanation, FONT['body'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_steps(prs, title, steps, mod_num, mod_name, page_num):
    """步骤流程页（Alignment + Proximity）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent)

    palette = [C['cyan'], C['violet'], C['orange'], C['emerald'], C['amber'], C['rose']]
    n = len(steps)

    if n <= 4:
        # 横排卡片（Alignment: 等宽等距）
        card_w = Inches(12.3 / n - 0.2)
        card_h = Inches(5.0)
        sx = Inches(0.5)
        gap = Inches(0.25)

        for i, step in enumerate(steps):
            x = sx + i * (card_w + gap)
            sc = palette[i % len(palette)]

            rrect(slide, x, Inches(1.5), card_w, card_h, fill=C['surface'])
            rect(slide, x, Inches(1.5), card_w, Pt(3), fill=sc)

            # 编号圆
            circle_size = Inches(0.55)
            cx = x + (card_w - circle_size) / 2
            circle(slide, cx, Inches(2.0), circle_size, fill=sc)
            txt(slide, cx, Inches(2.0), circle_size, circle_size,
                str(i + 1), 16, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)

            # 步骤文字
            txt(slide, x + Inches(0.2), Inches(2.8), card_w - Inches(0.4), Inches(3.5),
                step, FONT['body'], C['text_secondary'])

            # 连接箭头
            if i < n - 1:
                ax = x + card_w + Inches(0.02)
                txt(slide, ax, Inches(3.8), Inches(0.2), Inches(0.4),
                    "›", 20, sc, bold=True, align=PP_ALIGN.CENTER)
    else:
        # 竖排（Proximity: 步骤紧凑排列）
        for i, step in enumerate(steps):
            y = Inches(1.4) + i * Inches(0.88)
            if y > Inches(6.4):
                break
            sc = palette[i % len(palette)]

            rrect(slide, Inches(0.5), y, Inches(12.3), Inches(0.72), fill=C['surface'])
            rect(slide, Inches(0.5), y, Inches(0.06), Inches(0.72), fill=sc)

            circle(slide, Inches(0.75), y + Inches(0.13), Inches(0.4), fill=sc)
            txt(slide, Inches(0.75), y + Inches(0.13), Inches(0.4), Inches(0.4),
                str(i + 1), 12, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)

            txt(slide, Inches(1.35), y + Inches(0.1), Inches(11), Inches(0.5),
                step, FONT['body'], C['text_secondary'])

            if i < n - 1 and y + Inches(0.88) <= Inches(6.4):
                rect(slide, Inches(0.93), y + Inches(0.72), Pt(1.5), Inches(0.16), fill=sc)

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_practice(prs, title, tasks, mod_num, mod_name, page_num):
    """实操练习页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = C['emerald']

    page_header(slide, title, accent, "🛠️")

    for i, task in enumerate(tasks):
        y = Inches(1.4) + i * Inches(0.88)
        if y > Inches(6.4):
            break

        rrect(slide, Inches(0.5), y, Inches(12.3), Inches(0.72), fill=C['surface'])
        # checkbox
        rrect(slide, Inches(0.7), y + Inches(0.16), Inches(0.35), Inches(0.35),
              border=accent)
        # 编号
        txt(slide, Inches(1.2), y + Inches(0.1), Inches(0.5), Inches(0.5),
            f"#{i+1}", FONT['caption'], accent, bold=True)
        # 文字
        txt(slide, Inches(1.7), y + Inches(0.1), Inches(10.5), Inches(0.5),
            task, FONT['body'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_summary(prs, takeaways, mod_num, mod_name, page_num):
    """总结页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, "本节小结", accent, "📝")

    for i, item in enumerate(takeaways):
        y = Inches(1.4) + i * Inches(0.85)
        if y > Inches(6.4):
            break

        # 渐变色编号（视觉节奏）
        factor = max(0.5, 1.0 - i * 0.08)
        fc = RGBColor(
            min(255, int(accent.red * factor)),
            min(255, int(accent.green * factor)),
            min(255, int(accent.blue * factor))
        )

        rrect(slide, Inches(0.5), y, Inches(12.3), Inches(0.68), fill=C['surface'])
        circle(slide, Inches(0.7), y + Inches(0.1), Inches(0.4), fill=fc)
        txt(slide, Inches(0.7), y + Inches(0.1), Inches(0.4), Inches(0.4),
            str(i + 1), 12, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)
        txt(slide, Inches(1.3), y + Inches(0.08), Inches(11), Inches(0.5),
            item, FONT['body'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_section(prs, section_title, subtitle, mod_num, mod_name, page_num):
    """章节分隔页（Contrast: 大留白 + 大标题）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    # 左侧极细色条
    rect(slide, Inches(0), Inches(0), Inches(0.08), SLIDE_H, fill=accent)

    # 大面积留白，只放核心信息
    line(slide, Inches(1.2), Inches(2.8), Inches(3.5), accent, Pt(3))
    txt(slide, Inches(1.2), Inches(3.2), Inches(10), Inches(1.0),
        section_title, FONT['h1'] + 4, C['text_primary'], bold=True)
    txt(slide, Inches(1.2), Inches(4.4), Inches(10), Inches(0.6),
        subtitle, FONT['body'] + 2, C['text_tertiary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_timeline(prs, title, events, mod_num, mod_name, page_num):
    """时间线页（Alignment: 横向对齐）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent, "⏱️")

    n = len(events)
    if n == 0:
        page_footer(slide, mod_num, mod_name, page_num, accent)
        return

    palette = [C['cyan'], C['violet'], C['orange'], C['emerald'], C['amber'], C['rose']]

    # 横轴
    line_y = Inches(3.5)
    line_x = Inches(1.0)
    line_w = Inches(11.3)
    line(slide, line_x, line_y, line_w, C['border'], Pt(2))

    step = line_w / n
    for i, ev in enumerate(events):
        x = line_x + i * step
        sc = palette[i % len(palette)]

        # 节点
        circle(slide, x + Inches(0.35), line_y - Inches(0.12), Inches(0.25), fill=sc)

        # 标签
        label = ev.get('label', '') if isinstance(ev, dict) else str(ev)
        txt(slide, x, line_y - Inches(0.7), Inches(1.2), Inches(0.4),
            label, FONT['body'], sc, bold=True, align=PP_ALIGN.CENTER)

        # 描述（交替上下）
        desc = ev.get('desc', '') if isinstance(ev, dict) else ''
        if desc:
            if i % 2 == 0:
                rrect(slide, x - Inches(0.1), Inches(1.5), Inches(1.6), Inches(1.3), fill=C['surface'])
                txt(slide, x, Inches(1.6), Inches(1.4), Inches(1.1),
                    desc, FONT['caption'], C['text_secondary'])
            else:
                rrect(slide, x - Inches(0.1), Inches(4.0), Inches(1.6), Inches(1.3), fill=C['surface'])
                txt(slide, x, Inches(4.1), Inches(1.4), Inches(1.1),
                    desc, FONT['caption'], C['text_secondary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_icon_grid(prs, title, items, mod_num, mod_name, page_num):
    """图标网格页（Proximity + Alignment）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    page_header(slide, title, accent)

    cols = 3
    card_w = Inches(3.7)
    card_h = Inches(2.1)
    gap = Inches(0.35)
    sx = Inches(0.65)
    sy = Inches(1.5)

    palette = [C['cyan'], C['violet'], C['orange'], C['emerald'], C['amber'], C['rose']]

    for i, item in enumerate(items[:9]):
        col, row = i % cols, i // cols
        x = sx + col * (card_w + gap)
        y = sy + row * (card_h + gap)
        sc = palette[i % len(palette)]

        rrect(slide, x, y, card_w, card_h, fill=C['surface'])
        rect(slide, x, y, card_w, Pt(3), fill=sc)

        icon = item.get('icon', '📌') if isinstance(item, dict) else '📌'
        txt(slide, x + Inches(0.2), y + Inches(0.2), Inches(0.5), Inches(0.5),
            icon, 26, sc)

        item_title = item.get('title', '') if isinstance(item, dict) else item
        txt(slide, x + Inches(0.2), y + Inches(0.75), card_w - Inches(0.4), Inches(0.35),
            item_title, FONT['body'], C['text_primary'], bold=True)

        if isinstance(item, dict) and item.get('desc'):
            txt(slide, x + Inches(0.2), y + Inches(1.2), card_w - Inches(0.4), Inches(0.7),
                item['desc'], FONT['caption'], C['text_tertiary'])

    page_footer(slide, mod_num, mod_name, page_num, accent)


def make_end(prs, next_title, mod_num, mod_name):
    """结束页（Contrast: 视觉焦点居中）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, C['bg'])
    accent = _acc(mod_num)

    # 中央圆（视觉锚点）
    circle_size = Inches(2.0)
    cx = (SLIDE_W - circle_size) / 2
    circle(slide, cx, Inches(1.8), circle_size, fill=accent)
    txt(slide, cx, Inches(2.1), circle_size, Inches(1.5),
        "✓", 56, C['text_primary'], bold=True, align=PP_ALIGN.CENTER)

    # 文字
    txt(slide, Inches(0), Inches(4.2), SLIDE_W, Inches(0.7),
        "本模块结束", FONT['h1'], C['text_primary'], bold=True, align=PP_ALIGN.CENTER)

    line(slide, Inches(5.5), Inches(5.1), Inches(2.3), accent, Pt(2))

    if next_title:
        txt(slide, Inches(0), Inches(5.4), SLIDE_W, Inches(0.5),
            f"下一模块：{next_title}", FONT['body'] + 2, C['text_secondary'], align=PP_ALIGN.CENTER)

    txt(slide, Inches(0), Inches(6.5), SLIDE_W, Inches(0.4),
        "AI漫剧制作全流程课程 · 2026版", FONT['caption'], C['text_tertiary'], align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════
# PR 设置
# ══════════════════════════════════════════════

def make_pr(prs):
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


# ══════════════════════════════════════════════
# 模块内容定义
# ══════════════════════════════════════════════

MODULES = [
    {
        'num': 0, 'name': '课程导论',
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
            {'type': 'stat_cards', 'title': '行业数据一览', 'stats': [
                {'value': '240', 'unit': '亿元', 'label': '市场规模', 'trend': '↑ +50% YoY'},
                {'value': '2.8', 'unit': '亿', 'label': '用户规模', 'trend': '↑ +40% YoY'},
                {'value': '15', 'unit': '亿次/天', 'label': '日均播放', 'trend': '↑ +50% YoY'},
                {'value': '120', 'unit': '万', 'label': '创作者', 'trend': '↑ +50% YoY'},
            ]},
            {'type': 'two_col', 'title': '学习路径选择',
             'left_title': '🅰️ 零基础速成（8周）', 'left_items': ['快速上手出作品', '重点模块：1,2,3,4,6,10,11,12,17', '适合：想快速入行变现'],
             'right_title': '🅱️ 专业进阶（16周）', 'right_items': ['全面掌握全链路', '全部17个模块', '适合：追求专业深度']},
            {'type': 'step', 'title': '课程学习路线图', 'steps': [
                '行业认知 → 理解AI漫剧是什么、市场多大',
                '创作技能 → 剧本、分镜、提示词工程',
                '工具精通 → 即梦AI、可灵AI、Seedance等',
                '后期制作 → 音频、剪辑、调色、成片',
                '商业运营 → 平台分发、变现、合规',
                '综合实战 → 从0到1独立完成一集AI漫剧'
            ]},
            {'type': 'key_point', 'title': '核心竞争力',
             'key_point': 'AI漫剧的核心竞争力 = 创意 × 工具熟练度 × 量产能力',
             'explanation': '工具会更新迭代，但创意能力和工业化思维是持久的竞争力。本课程不仅教工具操作，更注重培养你的创作思维和生产效率。'},
        ]
    },
    {
        'num': 1, 'name': '行业认知与趋势',
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
                '剧本创作：用AI生成分集剧本',
                '分镜设计：将剧本转化为分镜表',
                '文生图：用即梦AI生成关键帧画面',
                '图生视频：用可灵AI/Seedance让画面动起来',
                '剪辑配音：用剪映完成后期制作'
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
# 生成引擎
# ══════════════════════════════════════════════

def generate_module_ppt(module_data):
    prs = Presentation()
    make_pr(prs)

    num = module_data['num']
    name = module_data['name']
    title = module_data['title']
    subtitle = module_data['subtitle']
    next_mod = module_data.get('next', '')
    pages = module_data['pages']

    make_cover(prs, title, subtitle, num)

    page_num = 1
    for page in pages:
        t = page['type']
        if t == 'toc':
            make_toc(prs, page['items'], num, name, page_num)
        elif t == 'content':
            make_content(prs, page['title'], page['bullets'], num, name, page_num)
        elif t == 'two_col':
            make_two_col(prs, page['title'], page['left_title'], page['left_items'],
                         page['right_title'], page['right_items'], num, name, page_num)
        elif t == 'table':
            make_table(prs, page['title'], page['headers'], page['rows'], num, name, page_num)
        elif t == 'stat_cards':
            make_stat_cards(prs, page['title'], page['stats'], num, name, page_num)
        elif t == 'icon_grid':
            make_icon_grid(prs, page['title'], page['items'], num, name, page_num)
        elif t == 'timeline':
            make_timeline(prs, page['title'], page['events'], num, name, page_num)
        elif t == 'key_point':
            make_key_point(prs, page.get('title', '核心要点'),
                           page['key_point'], page['explanation'], num, name, page_num)
        elif t == 'step':
            make_steps(prs, page['title'], page['steps'], num, name, page_num)
        elif t == 'practice':
            make_practice(prs, page['title'], page['tasks'], num, name, page_num)
        elif t == 'section':
            make_section(prs, page['title'], page['subtitle'], num, name, page_num)
        elif t == 'summary':
            make_summary(prs, page['items'], num, name, page_num)
        page_num += 1

    make_end(prs, next_mod, num, name)

    filename = f"模块{num:02d}-{name}.pptx" if num > 0 else "00-课程导论.pptx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    prs.save(filepath)
    print(f"✅ {filename}（{len(prs.slides)}页）")
    return filepath


if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for m in MODULES:
        generate_module_ppt(m)
    print(f"\n🎉 已生成 {len(MODULES)} 个PPT文件")
