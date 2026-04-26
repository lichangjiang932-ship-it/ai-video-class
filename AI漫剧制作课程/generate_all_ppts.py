#!/usr/bin/env python3
"""
AI漫剧制作课程 PPT 生成器 v6
高级感 = 极致克制 + 巨型文字 + 大面积留白 + 单色冲击
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── 配色：极简 ──
BG   = RGBColor(0x08, 0x08, 0x0C)  # 近黑
SURF = RGBColor(0x11, 0x11, 0x18)  # 卡片
LINE = RGBColor(0x1E, 0x1E, 0x2A)  # 分割线
T1   = RGBColor(0xF0, 0xF0, 0xF5)  # 主文字
T2   = RGBColor(0x90, 0x90, 0xA0)  # 次文字
T3   = RGBColor(0x50, 0x50, 0x60)  # 辅助

# 每个模块只用一个强调色（克制）
ACCENTS = [
    RGBColor(0x00, 0xBF, 0xFF), # 青
    RGBColor(0xFF, 0x5F, 0x1F), # 橙
    RGBColor(0x00, 0xBF, 0xFF), # 青
    RGBColor(0x8B, 0x50, 0xF0), # 紫
    RGBColor(0xFF, 0x5F, 0x1F), # 橙
    RGBColor(0x00, 0xD0, 0x80), # 绿
    RGBColor(0x30, 0x70, 0xFF), # 蓝
    RGBColor(0x8B, 0x50, 0xF0), # 紫
    RGBColor(0xFF, 0xA0, 0x00), # 琥珀
    RGBColor(0xFF, 0x30, 0x60), # 玫瑰
    RGBColor(0x00, 0xD0, 0x80), # 绿
    RGBColor(0xFF, 0x5F, 0x1F), # 橙
    RGBColor(0x00, 0xBF, 0xFF), # 青
    RGBColor(0x8B, 0x50, 0xF0), # 紫
    RGBColor(0xFF, 0xA0, 0x00), # 琥珀
    RGBColor(0x00, 0xD0, 0x80), # 绿
    RGBColor(0xFF, 0x30, 0x60), # 玫瑰
    RGBColor(0x30, 0x70, 0xFF), # 蓝
]

SW, SH = Inches(13.333), Inches(7.5)
OUT = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程/PPT课件'

# ── 基础 ──

def _a(n):
    return ACCENTS[n] if 0 <= n < len(ACCENTS) else ACCENTS[0]

def _bg(s):
    f = s.background.fill; f.solid(); f.fore_color.rgb = BG

def _rect(s, x, y, w, h, c):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = c; sh.line.fill.background()

def _rrect(s, x, y, w, h, c):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = c; sh.line.fill.background()

def _circ(s, x, y, sz, c):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, sz, sz)
    sh.fill.solid(); sh.fill.fore_color.rgb = c; sh.line.fill.background()

def _t(s, x, y, w, h, txt, sz, c=T1, bold=False, al=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = str(txt)
    p.font.size = Pt(sz); p.font.color.rgb = c
    p.font.bold = bold; p.font.name = 'Microsoft YaHei'; p.alignment = al

def _line(s, x, y, w, c=LINE):
    _rect(s, x, y, w, Pt(1.5), c)


# ══════════════════════════════════════════════
# 幻灯片类型 - 极致克制
# ══════════════════════════════════════════════

def make_cover(prs, title, subtitle, mn=0):
    """封面：左色块 + 超大编号 + 极简信息"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # ── 左侧：纯色块 ──
    _rect(s, Inches(0), Inches(0), Inches(5.5), SH, ac)

    # 编号：超大 + 半透明叠加（唯一的装饰）
    if mn > 0:
        # 暗层
        _t(s, Inches(0.4), Inches(1.0), Inches(4.5), Inches(3),
           f"{mn:02d}", 150, RGBColor(255,255,255), bold=True)
        # 亮层偏移（微妙的阴影感）
        _t(s, Inches(0.45), Inches(1.05), Inches(4.5), Inches(3),
           f"{mn:02d}", 150, RGBColor(min(255,ac[0]+60), min(255,ac[1]+60), min(255,ac[2]+60)), bold=True)
        # MODULE
        _t(s, Inches(0.6), Inches(4.2), Inches(3), Inches(0.4),
           "M O D U L E", 11, T1)
        _line(s, Inches(0.6), Inches(4.65), Inches(2), T1)
    else:
        _t(s, Inches(1.2), Inches(2.5), Inches(3), Inches(2),
           "🎬", 90, T1, al=PP_ALIGN.CENTER)

    # ── 右侧：大量留白 + 关键信息 ──
    # 标签
    _rrect(s, Inches(6.2), Inches(1.5), Inches(6.3), Inches(0.45), ac)
    _t(s, Inches(6.2), Inches(1.5), Inches(6.3), Inches(0.45),
       "AI漫剧制作全流程课程 · 2026版", 11, T1, al=PP_ALIGN.CENTER)

    # 主标题（大字）
    _t(s, Inches(6.2), Inches(2.6), Inches(6.3), Inches(1.8),
       title, 42, T1, bold=True)

    # 线
    _line(s, Inches(6.2), Inches(4.5), Inches(2.5), ac)

    # 副标题
    _t(s, Inches(6.2), Inches(4.9), Inches(6.3), Inches(0.8),
       subtitle, 16, T2)

    # 底部条
    _rect(s, Inches(6.2), Inches(6.1), Inches(6.3), Inches(0.6), SURF)
    for i, st in enumerate(["124课时", "17模块", "8-16周", "全流程"]):
        _t(s, Inches(6.3) + Inches(i * 1.575), Inches(6.1), Inches(1.5), Inches(0.6),
           st, 11, ac, bold=True, al=PP_ALIGN.CENTER)


def make_toc(prs, items, mn, mname, pn):
    """目录：极简编号列表（两列）"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # 标题
    _t(s, Inches(0.8), Inches(0.4), Inches(10), Inches(0.6),
       "本节内容", 28, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    # 两列编号列表（克制：没有卡片，没有圆圈，只有数字+文字）
    col1 = items[:len(items)//2 + len(items)%2]
    col2 = items[len(items)//2 + len(items)%2:]

    for i, item in enumerate(col1):
        y = Inches(1.4) + i * Inches(0.65)
        if y > Inches(6.4): break
        _t(s, Inches(0.8), y, Inches(0.5), Inches(0.5), f"{i+1:02d}", 13, ac, bold=True)
        _t(s, Inches(1.5), y, Inches(5), Inches(0.5), item, 14, T2)

    for i, item in enumerate(col2):
        y = Inches(1.4) + i * Inches(0.65)
        if y > Inches(6.4): break
        idx = len(col1) + i + 1
        _t(s, Inches(7.0), y, Inches(0.5), Inches(0.5), f"{idx:02d}", 13, ac, bold=True)
        _t(s, Inches(7.7), y, Inches(5), Inches(0.5), item, 14, T2)

    # 页脚
    _line(s, Inches(0), Inches(7.05), SW)
    _t(s, Inches(0.8), Inches(7.1), Inches(5), Inches(0.3),
       f"模块{mn:02d} · {mname}", 9, T3)

    _foot(s, mn, mname, pn, ac)


def make_content(prs, title, bullets, mn, mname, pn):
    """内容：标题 + 要点（极简行间距）"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # 标题
    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    # 要点（无卡片、无编号圆，只有色块标记 + 文字）
    n = len(bullets)
    for i, b in enumerate(bullets):
        y = Inches(1.4) + i * Inches(0.85)
        if y > Inches(6.4): break
        # 色块标记（4px宽）
        _rect(s, Inches(0.8), y + Inches(0.05), Inches(0.05), Inches(0.5), ac)
        _t(s, Inches(1.1), y, Inches(11.5), Inches(0.65), b, 15, T2)

    _foot(s, mn, mname, pn, ac)


def make_stat_cards(prs, title, stats, mn, mname, pn):
    """数据页：大数字 + 极少文字（视觉冲击力核心）"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    n = min(len(stats), 4)
    card_w = Inches(12.3 / n - 0.2)
    gap = Inches(0.25)
    sx = Inches(0.5)

    for i, st in enumerate(stats[:n]):
        x = sx + i * (card_w + gap)

        # 大数字（56pt，页面上最大的元素）
        _t(s, x + Inches(0.1), Inches(1.6), card_w - Inches(0.2), Inches(1.5),
           st.get('value', ''), 56, ac, bold=True, al=PP_ALIGN.CENTER)

        # 单位
        if st.get('unit'):
            _t(s, x + Inches(0.1), Inches(2.9), card_w - Inches(0.2), Inches(0.3),
               st['unit'], 14, T3, al=PP_ALIGN.CENTER)

        # 标签
        _t(s, x + Inches(0.1), Inches(3.4), card_w - Inches(0.2), Inches(0.4),
           st.get('label', ''), 16, T1, bold=True, al=PP_ALIGN.CENTER)

        # 趋势
        if st.get('trend'):
            is_up = any(c in st['trend'] for c in ['↑', '+'])
            tc = RGBColor(0x00, 0xD0, 0x80) if is_up else RGBColor(0xFF, 0x30, 0x60)
            _t(s, x + Inches(0.1), Inches(4.0), card_w - Inches(0.2), Inches(0.35),
               st['trend'], 13, tc, bold=True, al=PP_ALIGN.CENTER)

        # 描述（小字）
        if st.get('desc'):
            _t(s, x + Inches(0.1), Inches(4.5), card_w - Inches(0.2), Inches(0.8),
               st['desc'], 11, T3, al=PP_ALIGN.CENTER)

    _foot(s, mn, mname, pn, ac)


def make_two_col(prs, title, lt, li, rt, ri, mn, mname, pn):
    """双栏：左右对比，极简"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    # 左栏
    _t(s, Inches(0.8), Inches(1.4), Inches(5.5), Inches(0.5), lt, 18, ac, bold=True)
    _line(s, Inches(0.8), Inches(1.95), Inches(1.5), ac)
    for i, item in enumerate(li):
        y = Inches(2.2) + i * Inches(0.6)
        if y > Inches(6.4): break
        _t(s, Inches(0.8), y, Inches(5.5), Inches(0.5), f"▸  {item}", 14, T2)

    # 右栏
    _t(s, Inches(7.0), Inches(1.4), Inches(5.5), Inches(0.5), rt, 18, RGBColor(0x8B,0x50,0xF0), bold=True)
    _line(s, Inches(7.0), Inches(1.95), Inches(1.5), RGBColor(0x8B,0x50,0xF0))
    for i, item in enumerate(ri):
        y = Inches(2.2) + i * Inches(0.6)
        if y > Inches(6.4): break
        _t(s, Inches(7.0), y, Inches(5.5), Inches(0.5), f"▸  {item}", 14, T2)

    _foot(s, mn, mname, pn, ac)


def make_table(prs, title, headers, rows, mn, mname, pn):
    """表格页"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    cols = len(headers)
    nr = len(rows) + 1
    tw = Inches(12.1)
    rh = Inches(0.55)

    ts = s.shapes.add_table(nr, cols, Inches(0.6), Inches(1.4), tw, rh * nr)
    tbl = ts.table
    cw = tw / cols
    for j in range(cols): tbl.columns[j].width = int(cw)

    for j, h in enumerate(headers):
        c = tbl.cell(0, j); c.text = h; c.fill.solid(); c.fill.fore_color.rgb = ac
        for p in c.text_frame.paragraphs:
            p.font.size = Pt(12); p.font.color.rgb = T1; p.font.bold = True
            p.font.name = 'Microsoft YaHei'; p.alignment = PP_ALIGN.CENTER
        c.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = tbl.cell(i+1, j); c.text = str(val); c.fill.solid()
            c.fill.fore_color.rgb = SURF if i % 2 == 0 else BG
            for p in c.text_frame.paragraphs:
                p.font.size = Pt(11); p.font.color.rgb = T2
                p.font.name = 'Microsoft YaHei'; p.alignment = PP_ALIGN.CENTER
            c.vertical_anchor = MSO_ANCHOR.MIDDLE

    _foot(s, mn, mname, pn, ac)


def make_key_point(prs, title, kp, exp, mn, mname, pn):
    """重点页：大字居中（视觉焦点）"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # 核心要点：超大字居中（Apple Keynote 风格）
    _t(s, Inches(1.5), Inches(1.5), Inches(10.3), Inches(2.5),
       kp, 32, T1, bold=True, al=PP_ALIGN.CENTER)

    # 装饰线（唯一装饰）
    _line(s, Inches(5.5), Inches(4.2), Inches(2.3), ac)

    # 说明
    _t(s, Inches(1.5), Inches(4.6), Inches(10.3), Inches(1.8),
       exp, 15, T2, al=PP_ALIGN.CENTER)

    _foot(s, mn, mname, pn, ac)


def make_steps(prs, title, steps, mn, mname, pn):
    """步骤页"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    n = len(steps)
    if n <= 4:
        cw = Inches(12.3 / n - 0.2)
        gap = Inches(0.25)
        for i, step in enumerate(steps):
            x = Inches(0.5) + i * (cw + gap)
            # 编号（大字）
            _t(s, x + Inches(0.2), Inches(1.5), cw - Inches(0.4), Inches(0.8),
               f"{i+1}", 48, ac, bold=True, al=PP_ALIGN.CENTER)
            # 分割线
            _line(s, x + Inches(0.3), Inches(2.4), cw - Inches(0.6), ac)
            # 文字
            _t(s, x + Inches(0.2), Inches(2.7), cw - Inches(0.4), Inches(3.5),
               step, 14, T2)
            # 连接符
            if i < n - 1:
                _t(s, x + cw + Inches(0.02), Inches(2.0), Inches(0.2), Inches(0.3),
                   "›", 20, ac, bold=True, al=PP_ALIGN.CENTER)
    else:
        for i, step in enumerate(steps):
            y = Inches(1.4) + i * Inches(0.88)
            if y > Inches(6.4): break
            _rect(s, Inches(0.8), y + Inches(0.08), Inches(0.05), Inches(0.55), ac)
            _t(s, Inches(1.1), y, Inches(0.4), Inches(0.7), f"{i+1}", 13, ac, bold=True)
            _t(s, Inches(1.6), y, Inches(11), Inches(0.7), step, 15, T2)

    _foot(s, mn, mname, pn, ac)


def make_practice(prs, title, tasks, mn, mname, pn):
    """练习页"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = RGBColor(0x00, 0xD0, 0x80)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    for i, task in enumerate(tasks):
        y = Inches(1.4) + i * Inches(0.85)
        if y > Inches(6.4): break
        _rect(s, Inches(0.8), y + Inches(0.08), Inches(0.05), Inches(0.55), ac)
        _t(s, Inches(1.1), y, Inches(0.5), Inches(0.5), f"#{i+1}", 12, ac, bold=True)
        _t(s, Inches(1.6), y, Inches(11), Inches(0.65), task, 14, T2)

    _foot(s, mn, mname, pn, ac)


def make_summary(prs, items, mn, mname, pn):
    """总结页"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), "本节小结", 28, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    for i, item in enumerate(items):
        y = Inches(1.4) + i * Inches(0.8)
        if y > Inches(6.4): break
        _t(s, Inches(0.8), y, Inches(0.4), Inches(0.5), f"{i+1:02d}", 13, ac, bold=True)
        _t(s, Inches(1.4), y, Inches(11.5), Inches(0.6), item, 14, T2)

    _foot(s, mn, mname, pn, ac)


def make_section(prs, stitle, sub, mn, mname, pn):
    """章节页：大面积留白 + 大标题"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # 唯一装饰：左侧细线
    _rect(s, Inches(0), Inches(0), Inches(0.08), SH, ac)

    # 大标题（居中偏左，利用留白）
    _t(s, Inches(1.2), Inches(2.5), Inches(10), Inches(1.2),
       stitle, 36, T1, bold=True)
    _line(s, Inches(1.2), Inches(3.8), Inches(3.5), ac)
    _t(s, Inches(1.2), Inches(4.2), Inches(10), Inches(0.6),
       sub, 16, T3)

    _foot(s, mn, mname, pn, ac)


def make_timeline(prs, title, events, mn, mname, pn):
    """时间线"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    n = len(events)
    if n == 0:
        _foot(s, mn, mname, pn, ac)
        return

    ly = Inches(3.5)
    _line(s, Inches(1.0), ly, Inches(11.3), LINE)

    step = Inches(11.3) / n
    for i, ev in enumerate(events):
        x = Inches(1.0) + i * step
        _circ(s, x + Inches(0.35), ly - Inches(0.1), Inches(0.22), ac)
        label = ev.get('label', '') if isinstance(ev, dict) else str(ev)
        _t(s, x, ly - Inches(0.65), Inches(1.2), Inches(0.4), label, 14, ac, bold=True, al=PP_ALIGN.CENTER)
        desc = ev.get('desc', '') if isinstance(ev, dict) else ''
        if desc:
            dy = Inches(1.8) if i % 2 == 0 else Inches(4.0)
            _t(s, x, dy, Inches(1.4), Inches(1.0), desc, 11, T2, al=PP_ALIGN.CENTER)

    _foot(s, mn, mname, pn, ac)


def make_icon_grid(prs, title, items, mn, mname, pn):
    """图标网格"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    _t(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6), title, 24, T1, bold=True)
    _line(s, Inches(0.8), Inches(1.05), Inches(2), ac)

    cols = 3
    cw, ch = Inches(3.7), Inches(2.0)
    gap = Inches(0.35)
    sx, sy = Inches(0.65), Inches(1.4)

    for i, item in enumerate(items[:9]):
        col, row = i % cols, i // cols
        x = sx + col * (cw + gap)
        y = sy + row * (ch + gap)

        _rrect(s, x, y, cw, ch, SURF)
        icon = item.get('icon', '📌') if isinstance(item, dict) else '📌'
        _t(s, x + Inches(0.2), y + Inches(0.15), Inches(0.5), Inches(0.5), icon, 24, ac)
        it = item.get('title', '') if isinstance(item, dict) else item
        _t(s, x + Inches(0.2), y + Inches(0.7), cw - Inches(0.4), Inches(0.35), it, 14, T1, bold=True)
        if isinstance(item, dict) and item.get('desc'):
            _t(s, x + Inches(0.2), y + Inches(1.15), cw - Inches(0.4), Inches(0.7), item['desc'], 11, T3)

    _foot(s, mn, mname, pn, ac)


def make_end(prs, nt, mn, mname):
    """结束页"""
    s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s)
    ac = _a(mn)

    # 中央对勾（极简）
    _t(s, Inches(0), Inches(2.0), SW, Inches(1.0),
       "✓", 72, ac, bold=True, al=PP_ALIGN.CENTER)
    _t(s, Inches(0), Inches(3.3), SW, Inches(0.8),
       "本模块结束", 32, T1, bold=True, al=PP_ALIGN.CENTER)
    _line(s, Inches(5.5), Inches(4.3), Inches(2.3), ac)
    if nt:
        _t(s, Inches(0), Inches(4.6), SW, Inches(0.5),
           f"下一模块：{nt}", 16, T2, al=PP_ALIGN.CENTER)
    _t(s, Inches(0), Inches(6.5), SW, Inches(0.4),
       "AI漫剧制作全流程课程 · 2026版", 10, T3, al=PP_ALIGN.CENTER)


def _foot(s, mn, mname, pn, ac):
    _line(s, Inches(0), Inches(7.05), SW)
    _t(s, Inches(0.8), Inches(7.1), Inches(5), Inches(0.3),
       f"模块{mn:02d} · {mname}", 9, T3)


def make_pr(prs):
    prs.slide_width = SW; prs.slide_height = SH
    return prs


# ══════════════════════════════════════════════
# 模块定义
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
             'key_point': 'AI漫剧的核心竞争力\n= 创意 × 工具熟练度 × 量产能力',
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
                {'label': '2023', 'desc': '沙雕漫\n静态图+配音'},
                {'label': '2024', 'desc': '动态漫\n图生视频应用'},
                {'label': '2025', 'desc': 'AI原生\n全链路AI辅助'},
                {'label': '2026', 'desc': 'AI仿真人\n接近真人质量'},
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
