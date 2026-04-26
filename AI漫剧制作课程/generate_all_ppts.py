#!/usr/bin/env python3
"""
AI漫剧制作课程 PPT 生成器 v5
高级感 + 色彩感 + 视觉冲击力
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ══════════════════════════════════════════════
# 高级配色系统
# ══════════════════════════════════════════════

C = {
    # 背景
    'bg':       RGBColor(0x08, 0x0C, 0x18),
    'bg2':      RGBColor(0x0E, 0x14, 0x24),

    # 表面
    'surf':     RGBColor(0x13, 0x1B, 0x30),
    'surf2':    RGBColor(0x1A, 0x24, 0x3C),
    'border':   RGBColor(0x22, 0x30, 0x50),

    # 文字
    't1':       RGBColor(0xF5, 0xF7, 0xFC),
    't2':       RGBColor(0xA8, 0xB8, 0xD0),
    't3':       RGBColor(0x5C, 0x6E, 0x8A),
    't4':       RGBColor(0x34, 0x42, 0x5C),

    # 高饱和强调色（色彩感核心）
    'cyan':     RGBColor(0x00, 0xE5, 0xFF),   # 电光青
    'violet':   RGBColor(0xA8, 0x55, 0xF7),   # 亮紫
    'orange':   RGBColor(0xFF, 0x6D, 0x2E),   # 火焰橙
    'emerald':  RGBColor(0x00, 0xE6, 0x96),   # 翡翠绿
    'amber':    RGBColor(0xFF, 0xB8, 0x00),   # 金琥珀
    'rose':     RGBColor(0xFF, 0x3D, 0x71),   # 玫瑰红
    'blue':     RGBColor(0x38, 0x7A, 0xFF),   # 宝石蓝
    'pink':     RGBColor(0xFF, 0x6B, 0xB5),   # 亮粉

    # 半透明（玻璃态）
    'glass':    RGBColor(0x16, 0x20, 0x3A),
}

PALETTE = [C['cyan'], C['violet'], C['orange'], C['emerald'], C['amber'], C['rose'], C['blue'], C['pink']]

MODULE_ACCENTS = [
    C['cyan'], C['orange'], C['cyan'], C['violet'],
    C['orange'], C['emerald'], C['blue'], C['violet'],
    C['amber'], C['rose'], C['emerald'], C['orange'],
    C['cyan'], C['violet'], C['amber'], C['emerald'],
    C['rose'], C['blue'],
]

# 排版
F = {
    'display': 52, 'h1': 36, 'h2': 26, 'h3': 20,
    'body': 15, 'caption': 12, 'micro': 10,
}

SW, SH = Inches(13.333), Inches(7.5)
ML = Inches(0.8)
CW = Inches(11.733)
OUT = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程/PPT课件'

# ══════════════════════════════════════════════
# 基础工具
# ══════════════════════════════════════════════

def _a(n):
    return MODULE_ACCENTS[n] if 0 <= n < len(MODULE_ACCENTS) else C['cyan']

def _p(i):
    return PALETTE[i % len(PALETTE)]

def bg(slide, c=None):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = c or C['bg']

def rect(s, x, y, w, h, f=None, b=None):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.line.fill.background()
    if f: sh.fill.solid(); sh.fill.fore_color.rgb = f
    else: sh.fill.background()
    if b: sh.line.color.rgb = b; sh.line.width = Pt(1)
    return sh

def rrect(s, x, y, w, h, f=None, b=None):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.line.fill.background()
    if f: sh.fill.solid(); sh.fill.fore_color.rgb = f
    else: sh.fill.background()
    if b: sh.line.color.rgb = b; sh.line.width = Pt(1)
    return sh

def circ(s, x, y, sz, f=None):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, sz, sz)
    sh.line.fill.background()
    if f: sh.fill.solid(); sh.fill.fore_color.rgb = f
    return sh

def t(s, x, y, w, h, txt, sz=14, c=None, b=False, al=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = str(txt)
    p.font.size = Pt(sz); p.font.color.rgb = c or C['t1']
    p.font.bold = b; p.font.name = 'Microsoft YaHei'; p.alignment = al
    return tb

def ln(s, x, y, w, c=None, th=Pt(2)):
    return rect(s, x, y, w, th, f=c or C['border'])

def _foot(s, mn, mname, pn, ac):
    ln(s, Inches(0), Inches(7.05), SW, C['border'], Pt(1))
    t(s, ML, Inches(7.1), Inches(5), Inches(0.3), f"模块{mn:02d} · {mname}", F['micro'], C['t3'])
    t(s, Inches(11.5), Inches(7.1), Inches(1.5), Inches(0.3), str(pn), F['micro'], C['t3'], al=PP_ALIGN.RIGHT)

def _head(s, title, ac, icon=""):
    # 顶部渐变色条（3层叠加模拟渐变）
    rect(s, Inches(0), Inches(0), SW, Pt(6), f=ac)
    rect(s, Inches(0), Pt(6), SW, Pt(3), f=RGBColor(max(0,ac[0]-40), max(0,ac[1]-40), max(0,ac[2]-40)))
    # 标题区
    rect(s, Inches(0), Pt(9), SW, Inches(1.0), f=C['surf'])
    rect(s, Inches(0), Pt(9), Inches(0.12), Inches(1.0), f=ac)
    # 标题
    label = f"{icon} {title}" if icon else title
    t(s, ML, Inches(0.2), CW, Inches(0.7), label, F['h2'], C['t1'], b=True)
    ln(s, Inches(0), Inches(1.2), SW, C['border'], Pt(1))

# ══════════════════════════════════════════════
# 装饰系统（高级感 + 色彩感的核心）
# ══════════════════════════════════════════════

def deco_corner(s, ac):
    """右上角几何装饰"""
    rect(s, Inches(12.3), Inches(0), Inches(1.033), Pt(4), f=ac)
    rect(s, Inches(13.18), Inches(0), Pt(4), Inches(0.8), f=ac)

def deco_dots(s, x, y, rows, cols, color, spacing=Inches(0.25), size=Pt(4)):
    """点阵装饰"""
    for r in range(rows):
        for c in range(cols):
            circ(s, x + c * spacing, y + r * spacing, size, f=color)

def deco_circle_glow(s, x, y, size, color):
    """大圆光晕装饰（色彩感）"""
    circ(s, x, y, size, f=color)

def deco_accent_bar(s, x, y, w, h, ac):
    """渐变色条（3层模拟）"""
    rect(s, x, y, w, h, f=ac)
    c2 = RGBColor(min(255, ac[0] + 30), min(255, ac[1] + 30), min(255, ac[2] + 30))
    rect(s, x, y, w // 3, h, f=c2)

def deco_side_stripe(s, ac):
    """左侧渐变宽条（封面/章节用）"""
    rect(s, Inches(0), Inches(0), Inches(5.5), SH, f=ac)
    # 上层亮色叠加
    bright = RGBColor(min(255, ac[0] + 50), min(255, ac[1] + 50), min(255, ac[2] + 50))
    rect(s, Inches(0), Inches(0), Inches(5.5), Inches(0.1), f=C['t1'])
    rect(s, Inches(0), Inches(0), Inches(0.1), SH, f=C['t1'])
    # 点阵装饰
    deco_dots(s, Inches(0.5), Inches(5.5), 3, 8, RGBColor(255,255,255), Inches(0.35), Pt(3))

def deco_large_number(s, x, y, num, ac):
    """超大装饰数字（视觉冲击力）"""
    # 半透明背景数字
    t(s, x, y, Inches(4), Inches(3), f"{num:02d}", 140, C['surf2'], b=True)
    # 前景小数字
    t(s, x + Inches(0.1), y + Inches(0.15), Inches(3), Inches(2), f"{num:02d}", 100, ac, b=True)


# ══════════════════════════════════════════════
# 幻灯片类型
# ══════════════════════════════════════════════

def make_cover(prs, title, subtitle, mn=0):
    """封面页 - 大色块 + 大数字 + 装饰几何"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    # ── 左侧大色块（色彩感核心）──
    deco_side_stripe(s, ac)

    # 大号模块编号（双层阴影效果）
    if mn > 0:
        deco_large_number(s, Inches(0.6), Inches(1.2), mn, ac)
        # MODULE 标签
        rrect(s, Inches(0.7), Inches(4.0), Inches(1.8), Inches(0.35), f=C['surf'])
        t(s, Inches(0.7), Inches(4.0), Inches(1.8), Inches(0.35), "M O D U L E", 10, C['t2'], al=PP_ALIGN.CENTER)
        # 装饰线
        ln(s, Inches(0.7), Inches(4.5), Inches(2), ac, Pt(3))
    else:
        t(s, Inches(0.6), Inches(2.0), Inches(4.5), Inches(2), "🎬", 80, C['t1'], al=PP_ALIGN.CENTER)

    # ── 右侧内容区 ──
    # 课程标签（玻璃态卡片）
    rrect(s, Inches(6.0), Inches(1.5), Inches(6.5), Inches(0.5), f=ac)
    t(s, Inches(6.2), Inches(1.5), Inches(6.1), Inches(0.5),
      "AI漫剧制作全流程课程 · 2026版", 12, C['t1'], al=PP_ALIGN.CENTER)

    # 主标题（大字冲击力）
    t(s, Inches(6.0), Inches(2.6), Inches(6.5), Inches(1.8),
      title, F['display'] - 8, C['t1'], b=True)

    # 装饰线（渐变效果）
    deco_accent_bar(s, Inches(6.0), Inches(4.5), Inches(3), Pt(4), ac)

    # 副标题
    t(s, Inches(6.0), Inches(4.9), Inches(6.5), Inches(0.8),
      subtitle, F['body'] + 2, C['t2'])

    # 底部信息条
    rrect(s, Inches(6.0), Inches(6.0), Inches(6.5), Inches(0.7), f=C['surf'])
    stats = ["124课时", "17模块", "8-16周", "全流程"]
    for i, st in enumerate(stats):
        x = Inches(6.2) + Inches(i * 1.6)
        t(s, x, Inches(6.05), Inches(1.4), Inches(0.6), st, F['caption'], ac, b=True, al=PP_ALIGN.CENTER)

    # 右上角装饰
    deco_corner(s, ac)


def make_toc(prs, items, mn, mname, pn):
    """目录页 - 彩色卡片网格"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, "本节内容", ac, "📋")

    cols = 3
    cw, ch = Inches(3.7), Inches(1.05)
    gap = Inches(0.3)
    sx, sy = Inches(0.65), Inches(1.5)

    for i, item in enumerate(items):
        col, row = i % cols, i // cols
        x = sx + col * (cw + gap)
        y = sy + row * (ch + gap)
        if y > Inches(6.5): break

        sc = _p(i)

        # 卡片（左侧色条 + 编号圆）
        rrect(s, x, y, cw, ch, f=C['surf'], b=C['border'])
        rect(s, x, y, Inches(0.06), ch, f=sc)
        circ(s, x + Inches(0.15), y + Inches(0.22), Inches(0.45), f=sc)
        t(s, x + Inches(0.15), y + Inches(0.22), Inches(0.45), Inches(0.45),
          str(i + 1), 13, C['t1'], b=True, al=PP_ALIGN.CENTER)
        t(s, x + Inches(0.72), y + Inches(0.18), Inches(2.8), Inches(0.65),
          item, F['caption'] + 1, C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_content(prs, title, bullets, mn, mname, pn):
    """内容页 - 彩色编号卡片"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac)

    n = len(bullets)
    max_y = Inches(6.6)
    avail = max_y - Inches(1.4)
    ch = min(Inches(0.78), avail / max(n, 1) - Inches(0.08))
    gap = (avail - ch * n) / max(n - 1, 1) if n > 1 else Inches(0)

    for i, b in enumerate(bullets):
        y = Inches(1.4) + i * (ch + gap)
        if y + ch > max_y: break

        sc = _p(i)

        # 卡片
        rrect(s, Inches(0.5), y, Inches(12.3), ch, f=C['surf'])
        # 左侧渐变色条
        rect(s, Inches(0.5), y, Inches(0.08), ch, f=sc)
        # 编号圆
        circ(s, Inches(0.78), y + (ch - Inches(0.38)) / 2, Inches(0.38), f=sc)
        t(s, Inches(0.78), y + (ch - Inches(0.38)) / 2, Inches(0.38), Inches(0.38),
          str(i + 1), 11, C['t1'], b=True, al=PP_ALIGN.CENTER)
        # 文字
        t(s, Inches(1.35), y + (ch - Inches(0.4)) / 2, Inches(11), Inches(0.4),
          b, F['body'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_stat_cards(prs, title, stats, mn, mname, pn):
    """数据卡片页 - 大数字冲击力 + 色彩感"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac, "📊")

    n = min(len(stats), 4)
    cw = Inches(12.3 / n - 0.2)
    gap = Inches(0.25)
    sx = Inches(0.5)
    cy, ch = Inches(1.5), Inches(5.2)

    for i, st in enumerate(stats[:n]):
        x = sx + i * (cw + gap)
        sc = _p(i)

        # 卡片
        rrect(s, x, cy, cw, ch, f=C['surf'])
        # 顶部渐变色条（3层）
        rect(s, x, cy, cw, Pt(5), f=sc)
        bright = RGBColor(min(255, sc[0] + 40), min(255, sc[1] + 40), min(255, sc[2] + 40))
        rect(s, x, cy, cw // 3, Pt(5), f=bright)

        # 大数字（超大字号，视觉冲击力）
        t(s, x + Inches(0.1), cy + Inches(0.5), cw - Inches(0.2), Inches(1.3),
          st.get('value', ''), 56, sc, b=True, al=PP_ALIGN.CENTER)

        # 单位
        if st.get('unit'):
            t(s, x + Inches(0.1), cy + Inches(1.7), cw - Inches(0.2), Inches(0.35),
              st['unit'], F['body'], C['t3'], al=PP_ALIGN.CENTER)

        # 分割线
        ln(s, x + Inches(0.3), cy + Inches(2.2), cw - Inches(0.6), C['border'], Pt(1))

        # 标签
        t(s, x + Inches(0.1), cy + Inches(2.5), cw - Inches(0.2), Inches(0.4),
          st.get('label', ''), F['body'], C['t1'], b=True, al=PP_ALIGN.CENTER)

        # 描述
        if st.get('desc'):
            t(s, x + Inches(0.15), cy + Inches(3.1), cw - Inches(0.3), Inches(1.0),
              st['desc'], F['caption'], C['t3'], al=PP_ALIGN.CENTER)

        # 趋势标签
        if st.get('trend'):
            is_up = any(c in st['trend'] for c in ['↑', '+'])
            tc = C['emerald'] if is_up else C['rose']
            rrect(s, x + Inches(0.25), cy + Inches(4.4), cw - Inches(0.5), Inches(0.38), f=C['bg2'])
            t(s, x + Inches(0.25), cy + Inches(4.4), cw - Inches(0.5), Inches(0.38),
              st['trend'], F['caption'], tc, b=True, al=PP_ALIGN.CENTER)

    _foot(s, mn, mname, pn, ac)


def make_two_col(prs, title, lt, li, rt, ri, mn, mname, pn):
    """双栏对比页 - 双色对比"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac)

    cw, ch = Inches(5.85), Inches(5.2)
    cy = Inches(1.5)
    lc, rc = ac, C['violet']

    # 左栏
    rrect(s, Inches(0.5), cy, cw, ch, f=C['surf'])
    rect(s, Inches(0.5), cy, cw, Pt(4), f=lc)
    t(s, Inches(0.8), cy + Inches(0.3), Inches(5.2), Inches(0.45), lt, F['h3'], lc, b=True)
    ln(s, Inches(0.8), cy + Inches(0.85), Inches(1.5), lc, Pt(2))
    for i, item in enumerate(li):
        iy = cy + Inches(1.15) + i * Inches(0.55)
        if iy > cy + ch - Inches(0.3): break
        circ(s, Inches(0.8), iy + Inches(0.06), Inches(0.2), f=lc)
        t(s, Inches(1.15), iy, Inches(5.0), Inches(0.45), item, F['body'], C['t2'])

    # 右栏
    rrect(s, Inches(6.85), cy, cw, ch, f=C['surf'])
    rect(s, Inches(6.85), cy, cw, Pt(4), f=rc)
    t(s, Inches(7.15), cy + Inches(0.3), Inches(5.2), Inches(0.45), rt, F['h3'], rc, b=True)
    ln(s, Inches(7.15), cy + Inches(0.85), Inches(1.5), rc, Pt(2))
    for i, item in enumerate(ri):
        iy = cy + Inches(1.15) + i * Inches(0.55)
        if iy > cy + ch - Inches(0.3): break
        circ(s, Inches(7.15), iy + Inches(0.06), Inches(0.2), f=rc)
        t(s, Inches(7.5), iy, Inches(5.0), Inches(0.45), item, F['body'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_table(prs, title, headers, rows, mn, mname, pn):
    """表格页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac)

    cols = len(headers)
    nr = len(rows) + 1
    tl, tt = Inches(0.6), Inches(1.5)
    tw = Inches(12.1)
    rh = Inches(0.6)

    ts = s.shapes.add_table(nr, cols, tl, tt, tw, rh * nr)
    tbl = ts.table
    cw = tw / cols
    for j in range(cols): tbl.columns[j].width = int(cw)

    for j, h in enumerate(headers):
        c = tbl.cell(0, j); c.text = h; c.fill.solid(); c.fill.fore_color.rgb = ac
        for p in c.text_frame.paragraphs:
            p.font.size = Pt(13); p.font.color.rgb = C['t1']; p.font.bold = True
            p.font.name = 'Microsoft YaHei'; p.alignment = PP_ALIGN.CENTER
        c.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = tbl.cell(i+1, j); c.text = str(val); c.fill.solid()
            c.fill.fore_color.rgb = C['surf'] if i % 2 == 0 else C['surf2']
            for p in c.text_frame.paragraphs:
                p.font.size = Pt(12); p.font.color.rgb = C['t2']
                p.font.name = 'Microsoft YaHei'; p.alignment = PP_ALIGN.CENTER
            c.vertical_anchor = MSO_ANCHOR.MIDDLE

    _foot(s, mn, mname, pn, ac)


def make_key_point(prs, title, kp, exp, mn, mname, pn):
    """重点强调页 - 大视觉焦点"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title or "核心要点", ac, "💡")

    # 大号重点卡片
    cx, cy, cw, ch = Inches(0.8), Inches(1.6), Inches(11.7), Inches(3.2)
    rrect(s, cx, cy, cw, ch, f=C['surf'])
    # 双色边框（视觉冲击力）
    rect(s, cx, cy, cw, Pt(5), f=ac)
    rect(s, cx, cy, Pt(5), ch, f=ac)
    # 右上角装饰圆
    circ(s, cx + cw - Inches(1.2), cy + Inches(0.3), Inches(0.8), f=RGBColor(
        min(255, ac[0]), min(255, ac[1]), min(255, ac[2])))

    # 核心文字
    t(s, cx + Inches(0.5), cy + Inches(0.5), cw - Inches(1.5), Inches(2.2),
      kp, F['h1'], C['t1'], b=True)

    # 说明区
    rrect(s, Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.6), f=C['surf2'])
    rect(s, Inches(0.8), Inches(5.1), Pt(4), Inches(1.6), f=C['t3'])
    t(s, Inches(1.3), Inches(5.25), Inches(10.8), Inches(1.3), exp, F['body'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_steps(prs, title, steps, mn, mname, pn):
    """步骤流程页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac)

    n = len(steps)
    if n <= 4:
        cw = Inches(12.3 / n - 0.2)
        ch = Inches(5.0)
        sx = Inches(0.5)
        gap = Inches(0.25)

        for i, step in enumerate(steps):
            x = sx + i * (cw + gap)
            sc = _p(i)

            rrect(s, x, Inches(1.5), cw, ch, f=C['surf'])
            # 顶部渐变色条
            rect(s, x, Inches(1.5), cw, Pt(4), f=sc)
            bright = RGBColor(min(255, sc[0]+40), min(255, sc[1]+40), min(255, sc[2]+40))
            rect(s, x, Inches(1.5), cw//3, Pt(4), f=bright)

            # 编号
            csz = Inches(0.6)
            cx = x + (cw - csz) / 2
            circ(s, cx, Inches(2.0), csz, f=sc)
            t(s, cx, Inches(2.0), csz, csz, str(i+1), 18, C['t1'], b=True, al=PP_ALIGN.CENTER)

            # 文字
            t(s, x + Inches(0.2), Inches(2.9), cw - Inches(0.4), Inches(3.5),
              step, F['body'], C['t2'])

            # 连接箭头
            if i < n - 1:
                t(s, x + cw + Inches(0.02), Inches(3.8), Inches(0.2), Inches(0.4),
                  "›", 22, sc, b=True, al=PP_ALIGN.CENTER)
    else:
        for i, step in enumerate(steps):
            y = Inches(1.4) + i * Inches(0.88)
            if y > Inches(6.4): break
            sc = _p(i)

            rrect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), f=C['surf'])
            rect(s, Inches(0.5), y, Inches(0.08), Inches(0.72), f=sc)
            circ(s, Inches(0.78), y + Inches(0.14), Inches(0.4), f=sc)
            t(s, Inches(0.78), y + Inches(0.14), Inches(0.4), Inches(0.4),
              str(i+1), 12, C['t1'], b=True, al=PP_ALIGN.CENTER)
            t(s, Inches(1.38), y + Inches(0.1), Inches(11), Inches(0.5), step, F['body'], C['t2'])
            if i < n - 1 and y + Inches(0.88) <= Inches(6.4):
                rect(s, Inches(0.96), y + Inches(0.72), Pt(1.5), Inches(0.16), f=sc)

    _foot(s, mn, mname, pn, ac)


def make_practice(prs, title, tasks, mn, mname, pn):
    """实操练习页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = C['emerald']

    _head(s, title, ac, "🛠️")

    for i, task in enumerate(tasks):
        y = Inches(1.4) + i * Inches(0.88)
        if y > Inches(6.4): break

        rrect(s, Inches(0.5), y, Inches(12.3), Inches(0.72), f=C['surf'])
        rrect(s, Inches(0.7), y + Inches(0.16), Inches(0.36), Inches(0.36), b=ac)
        t(s, Inches(1.2), y + Inches(0.1), Inches(0.5), Inches(0.5), f"#{i+1}", F['caption'], ac, b=True)
        t(s, Inches(1.7), y + Inches(0.1), Inches(10.5), Inches(0.5), task, F['body'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_summary(prs, items, mn, mname, pn):
    """总结页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, "本节小结", ac, "📝")

    for i, item in enumerate(items):
        y = Inches(1.4) + i * Inches(0.85)
        if y > Inches(6.4): break

        sc = _p(i)
        rrect(s, Inches(0.5), y, Inches(12.3), Inches(0.68), f=C['surf'])
        circ(s, Inches(0.7), y + Inches(0.1), Inches(0.4), f=sc)
        t(s, Inches(0.7), y + Inches(0.1), Inches(0.4), Inches(0.4),
          str(i+1), 12, C['t1'], b=True, al=PP_ALIGN.CENTER)
        t(s, Inches(1.3), y + Inches(0.08), Inches(11), Inches(0.5), item, F['body'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_section(prs, stitle, sub, mn, mname, pn):
    """章节分隔页 - 大留白 + 大标题 + 装饰"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    # 左侧宽色条
    rect(s, Inches(0), Inches(0), Inches(0.12), SH, f=ac)
    # 装饰大圆（半透明效果）
    circ(s, Inches(9), Inches(-1), Inches(5), f=C['surf'])
    circ(s, Inches(10), Inches(4), Inches(3), f=C['surf'])

    # 标题
    ln(s, Inches(1.2), Inches(2.8), Inches(4), ac, Pt(3))
    t(s, Inches(1.2), Inches(3.2), Inches(10), Inches(1.0), stitle, F['h1'] + 4, C['t1'], b=True)
    t(s, Inches(1.2), Inches(4.4), Inches(10), Inches(0.6), sub, F['body'] + 2, C['t3'])

    _foot(s, mn, mname, pn, ac)


def make_timeline(prs, title, events, mn, mname, pn):
    """时间线页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac, "⏱️")

    n = len(events)
    if n == 0:
        _foot(s, mn, mname, pn, ac)
        return

    ly = Inches(3.5)
    lx = Inches(1.0)
    lw = Inches(11.3)
    ln(s, lx, ly, lw, C['border'], Pt(2))

    step = lw / n
    for i, ev in enumerate(events):
        x = lx + i * step
        sc = _p(i)

        circ(s, x + Inches(0.35), ly - Inches(0.12), Inches(0.28), f=sc)
        label = ev.get('label', '') if isinstance(ev, dict) else str(ev)
        t(s, x, ly - Inches(0.7), Inches(1.2), Inches(0.4), label, F['body'], sc, b=True, al=PP_ALIGN.CENTER)

        desc = ev.get('desc', '') if isinstance(ev, dict) else ''
        if desc:
            dy = Inches(1.5) if i % 2 == 0 else Inches(4.0)
            rrect(s, x - Inches(0.1), dy, Inches(1.6), Inches(1.3), f=C['surf'])
            rect(s, x - Inches(0.1), dy, Inches(1.6), Pt(3), f=sc)
            t(s, x, dy + Inches(0.15), Inches(1.4), Inches(1.1), desc, F['caption'], C['t2'])

    _foot(s, mn, mname, pn, ac)


def make_icon_grid(prs, title, items, mn, mname, pn):
    """图标网格页"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    _head(s, title, ac)

    cols = 3
    cw, ch = Inches(3.7), Inches(2.1)
    gap = Inches(0.35)
    sx, sy = Inches(0.65), Inches(1.5)

    for i, item in enumerate(items[:9]):
        col, row = i % cols, i // cols
        x = sx + col * (cw + gap)
        y = sy + row * (ch + gap)
        sc = _p(i)

        rrect(s, x, y, cw, ch, f=C['surf'])
        rect(s, x, y, cw, Pt(4), f=sc)

        icon = item.get('icon', '📌') if isinstance(item, dict) else '📌'
        t(s, x + Inches(0.2), y + Inches(0.2), Inches(0.5), Inches(0.5), icon, 28, sc)

        it = item.get('title', '') if isinstance(item, dict) else item
        t(s, x + Inches(0.2), y + Inches(0.78), cw - Inches(0.4), Inches(0.35),
          it, F['body'], C['t1'], b=True)

        if isinstance(item, dict) and item.get('desc'):
            t(s, x + Inches(0.2), y + Inches(1.2), cw - Inches(0.4), Inches(0.7),
              item['desc'], F['caption'], C['t3'])

    _foot(s, mn, mname, pn, ac)


def make_end(prs, next_title, mn, mname):
    """结束页 - 视觉焦点居中"""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, C['bg'])
    ac = _a(mn)

    # 中央大圆（渐变效果：多层叠加）
    csz = Inches(2.2)
    cx = (SW - csz) / 2
    # 外圈光晕
    circ(s, cx - Inches(0.15), Inches(1.65), csz + Inches(0.3), f=RGBColor(
        max(0, ac[0] - 60), max(0, ac[1] - 60), max(0, ac[2] - 60)))
    circ(s, cx, Inches(1.8), csz, f=ac)
    t(s, cx, Inches(2.1), csz, Inches(1.5), "✓", 56, C['t1'], b=True, al=PP_ALIGN.CENTER)

    t(s, Inches(0), Inches(4.3), SW, Inches(0.7), "本模块结束", F['h1'], C['t1'], b=True, al=PP_ALIGN.CENTER)

    deco_accent_bar(s, Inches(5.5), Inches(5.2), Inches(2.3), Pt(3), ac)

    if next_title:
        t(s, Inches(0), Inches(5.5), SW, Inches(0.5),
          f"下一模块：{next_title}", F['body'] + 2, C['t2'], al=PP_ALIGN.CENTER)

    t(s, Inches(0), Inches(6.5), SW, Inches(0.4),
      "AI漫剧制作全流程课程 · 2026版", F['caption'], C['t3'], al=PP_ALIGN.CENTER)


def make_pr(prs):
    prs.slide_width = SW
    prs.slide_height = SH
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


def generate_module_ppt(md):
    prs = Presentation(); make_pr(prs)
    num, name = md['num'], md['name']
    make_cover(prs, md['title'], md['subtitle'], num)
    pn = 1
    for page in md['pages']:
        t = page['type']
        if t == 'toc': make_toc(prs, page['items'], num, name, pn)
        elif t == 'content': make_content(prs, page['title'], page['bullets'], num, name, pn)
        elif t == 'two_col': make_two_col(prs, page['title'], page['left_title'], page['left_items'], page['right_title'], page['right_items'], num, name, pn)
        elif t == 'table': make_table(prs, page['title'], page['headers'], page['rows'], num, name, pn)
        elif t == 'stat_cards': make_stat_cards(prs, page['title'], page['stats'], num, name, pn)
        elif t == 'icon_grid': make_icon_grid(prs, page['title'], page['items'], num, name, pn)
        elif t == 'timeline': make_timeline(prs, page['title'], page['events'], num, name, pn)
        elif t == 'key_point': make_key_point(prs, page.get('title', '核心要点'), page['key_point'], page['explanation'], num, name, pn)
        elif t == 'step': make_steps(prs, page['title'], page['steps'], num, name, pn)
        elif t == 'practice': make_practice(prs, page['title'], page['tasks'], num, name, pn)
        elif t == 'section': make_section(prs, page['title'], page['subtitle'], num, name, pn)
        elif t == 'summary': make_summary(prs, page['items'], num, name, pn)
        pn += 1
    make_end(prs, md.get('next', ''), num, name)
    fn = f"模块{num:02d}-{name}.pptx" if num > 0 else "00-课程导论.pptx"
    fp = os.path.join(OUT, fn)
    prs.save(fp)
    print(f"✅ {fn}（{len(prs.slides)}页）")
    return fp


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    for m in MODULES:
        generate_module_ppt(m)
    print(f"\n🎉 已生成 {len(MODULES)} 个PPT文件")
