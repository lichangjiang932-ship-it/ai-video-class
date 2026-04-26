#!/usr/bin/env python3
"""
AI漫剧制作课程 PPT 批量生成脚本
生成18个专业PPT课件（00-课程导论 + 17个模块）
"""

import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── 配色方案 ──
COLORS = {
    'bg_dark':    RGBColor(0x0F, 0x17, 0x2A),  # 深蓝黑背景
    'bg_card':    RGBColor(0x1A, 0x25, 0x3C),  # 卡片背景
    'accent':     RGBColor(0x00, 0xD4, 0xFF),  # 亮青色强调
    'accent2':    RGBColor(0x7C, 0x3A, 0xED),  # 紫色
    'accent3':    RGBColor(0xFF, 0x6B, 0x35),  # 橙色
    'text_white': RGBColor(0xFF, 0xFF, 0xFF),
    'text_light': RGBColor(0xB0, 0xBC, 0xD4),
    'text_dim':   RGBColor(0x6B, 0x7B, 0x9E),
    'success':    RGBColor(0x10, 0xB9, 0x81),
    'warning':    RGBColor(0xF5, 0x9E, 0x0B),
    'danger':     RGBColor(0xEF, 0x44, 0x44),
    'border':     RGBColor(0x2D, 0x3B, 0x55),
}

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

OUTPUT_DIR = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程/PPT课件'


def set_slide_bg(slide, color):
    """设置幻灯片背景色"""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, left, top, width, height, fill_color=None, border_color=None, border_width=Pt(0)):
    """添加矩形形状"""
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


def add_rounded_rect(slide, left, top, width, height, fill_color=None):
    """添加圆角矩形"""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18, color=None, bold=False, alignment=PP_ALIGN.LEFT, font_name='Microsoft YaHei'):
    """添加文本框"""
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


def add_multi_text(slide, left, top, width, height, lines, font_size=16, color=None, line_spacing=1.5, bullet=False):
    """添加多行文本"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if bullet and line.strip():
            p.text = f"• {line}"
        else:
            p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color or COLORS['text_light']
        p.font.name = 'Microsoft YaHei'
        p.space_after = Pt(font_size * 0.6)
    return txBox


def add_accent_line(slide, left, top, width, color=None):
    """添加装饰线"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color or COLORS['accent']
    shape.line.fill.background()
    return shape


def add_number_badge(slide, left, top, number, color=None):
    """添加数字标记"""
    c = color or COLORS['accent']
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, Inches(0.5), Inches(0.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = c
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = str(number)
    p.font.size = Pt(14)
    p.font.color.rgb = COLORS['text_white']
    p.font.bold = True
    p.font.name = 'Microsoft YaHei'
    p.alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].space_before = Pt(0)
    tf.paragraphs[0].space_after = Pt(0)
    return shape


def add_footer(slide, module_num, module_name, page_num):
    """添加页脚"""
    # 底部装饰线
    add_shape(slide, Inches(0), Inches(7.1), SLIDE_WIDTH, Pt(2), fill_color=COLORS['border'])
    # 模块信息
    add_text_box(slide, Inches(0.5), Inches(7.15), Inches(6), Inches(0.3),
                 f"模块{module_num:02d}：{module_name}", font_size=10, color=COLORS['text_dim'])
    # 页码
    add_text_box(slide, Inches(11.5), Inches(7.15), Inches(1.5), Inches(0.3),
                 str(page_num), font_size=10, color=COLORS['text_dim'], alignment=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
# 幻灯片类型工厂
# ══════════════════════════════════════════════════════════

def make_title_slide(prs, title, subtitle, module_num=0):
    """封面页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_slide_bg(slide, COLORS['bg_dark'])

    # 顶部装饰条
    add_shape(slide, Inches(0), Inches(0), SLIDE_WIDTH, Inches(0.08), fill_color=COLORS['accent'])

    # 左侧装饰块
    add_shape(slide, Inches(0), Inches(0), Inches(0.15), Inches(3.5), fill_color=COLORS['accent'])

    # 课程标识
    add_text_box(slide, Inches(0.8), Inches(1.0), Inches(11), Inches(0.5),
                 "AI漫剧制作全流程课程 · 2026版", font_size=14, color=COLORS['accent'])

    # 主标题
    add_text_box(slide, Inches(0.8), Inches(1.8), Inches(11), Inches(1.5),
                 title, font_size=40, color=COLORS['text_white'], bold=True)

    # 装饰线
    add_accent_line(slide, Inches(0.8), Inches(3.5), Inches(3))

    # 副标题
    add_text_box(slide, Inches(0.8), Inches(3.9), Inches(10), Inches(1.0),
                 subtitle, font_size=20, color=COLORS['text_light'])

    # 模块编号（大号装饰）
    if module_num > 0:
        add_text_box(slide, Inches(9), Inches(4.5), Inches(4), Inches(2.5),
                     f"{module_num:02d}", font_size=120, color=COLORS['bg_card'], bold=True)

    # 底部信息
    add_shape(slide, Inches(0), Inches(6.8), SLIDE_WIDTH, Pt(2), fill_color=COLORS['border'])
    add_text_box(slide, Inches(0.8), Inches(6.9), Inches(5), Inches(0.4),
                 "124课时 · 17个模块 · 从零基础到独立制作", font_size=12, color=COLORS['text_dim'])


def make_toc_slide(prs, items, module_num, module_name, page_num):
    """目录页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    # 标题
    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(8), Inches(0.6),
                 "📋 本节内容", font_size=28, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    # 内容列表（两列）
    col1 = items[:len(items)//2 + len(items)%2]
    col2 = items[len(items)//2 + len(items)%2:]

    for i, item in enumerate(col1):
        y = Inches(1.5) + Inches(i * 0.7)
        add_number_badge(slide, Inches(0.8), y, i+1)
        add_text_box(slide, Inches(1.5), y, Inches(5), Inches(0.5),
                     item, font_size=16, color=COLORS['text_light'])

    for i, item in enumerate(col2):
        y = Inches(1.5) + Inches(i * 0.7)
        add_number_badge(slide, Inches(6.8), y, len(col1)+i+1, color=COLORS['accent2'])
        add_text_box(slide, Inches(7.5), y, Inches(5), Inches(0.5),
                     item, font_size=16, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_section_slide(prs, section_title, subtitle, module_num, module_name, page_num):
    """章节分隔页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    # 大号背景装饰
    add_shape(slide, Inches(0), Inches(2), Inches(0.2), Inches(3.5), fill_color=COLORS['accent2'])

    add_text_box(slide, Inches(1), Inches(2.5), Inches(10), Inches(1.0),
                 section_title, font_size=36, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(1), Inches(3.6), Inches(2.5), color=COLORS['accent2'])
    add_text_box(slide, Inches(1), Inches(4.0), Inches(10), Inches(0.8),
                 subtitle, font_size=18, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_content_slide(prs, title, bullets, module_num, module_name, page_num, accent_color=None):
    """标准内容页（带要点）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    c = accent_color or COLORS['accent']

    # 标题
    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2), color=c)

    # 要点内容
    for i, bullet in enumerate(bullets):
        y = Inches(1.5) + Inches(i * 0.85)
        if y > Inches(6.5):
            break
        # 要点标记
        add_shape(slide, Inches(0.8), y + Pt(5), Pt(8), Pt(8), fill_color=c)
        add_text_box(slide, Inches(1.2), y, Inches(11), Inches(0.7),
                     bullet, font_size=16, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_two_col_slide(prs, title, left_title, left_items, right_title, right_items, module_num, module_name, page_num):
    """双栏对比页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    # 标题
    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    # 左栏
    add_rounded_rect(slide, Inches(0.5), Inches(1.5), Inches(5.8), Inches(5.2), fill_color=COLORS['bg_card'])
    add_text_box(slide, Inches(0.8), Inches(1.7), Inches(5), Inches(0.5),
                 left_title, font_size=18, color=COLORS['accent'], bold=True)
    for i, item in enumerate(left_items):
        y = Inches(2.4) + Inches(i * 0.65)
        add_text_box(slide, Inches(1.0), y, Inches(5), Inches(0.5),
                     f"• {item}", font_size=14, color=COLORS['text_light'])

    # 右栏
    add_rounded_rect(slide, Inches(6.8), Inches(1.5), Inches(5.8), Inches(5.2), fill_color=COLORS['bg_card'])
    add_text_box(slide, Inches(7.1), Inches(1.7), Inches(5), Inches(0.5),
                 right_title, font_size=18, color=COLORS['accent2'], bold=True)
    for i, item in enumerate(right_items):
        y = Inches(2.4) + Inches(i * 0.65)
        add_text_box(slide, Inches(7.3), y, Inches(5), Inches(0.5),
                     f"• {item}", font_size=14, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_table_slide(prs, title, headers, rows, module_num, module_name, page_num):
    """表格页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    # 表格
    cols = len(headers)
    tbl_rows = len(rows) + 1
    col_w = Inches(11.5) / cols

    table_shape = slide.shapes.add_table(tbl_rows, cols, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.5 * tbl_rows))
    table = table_shape.table

    # 表头
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLORS['accent']
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(13)
            p.font.color.rgb = COLORS['text_white']
            p.font.bold = True
            p.font.name = 'Microsoft YaHei'
            p.alignment = PP_ALIGN.CENTER

    # 数据行
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i+1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLORS['bg_card'] if i % 2 == 0 else COLORS['bg_dark']
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.color.rgb = COLORS['text_light']
                p.font.name = 'Microsoft YaHei'
                p.alignment = PP_ALIGN.CENTER

    add_footer(slide, module_num, module_name, page_num)


def make_key_point_slide(prs, title, key_point, explanation, module_num, module_name, page_num):
    """重点强调页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    # 大号重点卡片
    add_rounded_rect(slide, Inches(1), Inches(2), Inches(11), Inches(2.5), fill_color=COLORS['bg_card'])
    add_shape(slide, Inches(1), Inches(2), Pt(6), Inches(2.5), fill_color=COLORS['warning'])
    add_text_box(slide, Inches(1.5), Inches(2.3), Inches(10), Inches(0.5),
                 "💡 核心要点", font_size=14, color=COLORS['warning'], bold=True)
    add_text_box(slide, Inches(1.5), Inches(2.9), Inches(10), Inches(1.2),
                 key_point, font_size=22, color=COLORS['text_white'], bold=True)

    # 详细说明
    add_text_box(slide, Inches(1), Inches(5.0), Inches(11), Inches(1.5),
                 explanation, font_size=15, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_step_slide(prs, title, steps, module_num, module_name, page_num):
    """步骤流程页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    colors = [COLORS['accent'], COLORS['accent2'], COLORS['accent3'], COLORS['success'], COLORS['warning'], COLORS['danger']]

    for i, step in enumerate(steps):
        y = Inches(1.5) + Inches(i * 1.0)
        if y > Inches(6.5):
            break
        c = colors[i % len(colors)]
        # 步骤编号圆
        add_number_badge(slide, Inches(0.8), y, i+1, color=c)
        # 连接线（非最后一步）
        if i < len(steps) - 1:
            add_shape(slide, Inches(1.03), y + Inches(0.5), Pt(2), Inches(0.5), fill_color=COLORS['border'])
        # 步骤文字
        add_text_box(slide, Inches(1.6), y, Inches(10.5), Inches(0.8),
                     step, font_size=16, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_practice_slide(prs, title, tasks, module_num, module_name, page_num):
    """实操练习页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 "🛠️ " + title, font_size=26, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2), color=COLORS['success'])

    for i, task in enumerate(tasks):
        y = Inches(1.5) + Inches(i * 0.85)
        if y > Inches(6.5):
            break
        # checkbox
        add_rounded_rect(slide, Inches(0.8), y, Inches(0.35), Inches(0.35), fill_color=COLORS['bg_card'])
        add_text_box(slide, Inches(1.4), y, Inches(11), Inches(0.7),
                     task, font_size=16, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_summary_slide(prs, takeaways, module_num, module_name, page_num):
    """总结页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
                 "📝 本节小结", font_size=28, color=COLORS['text_white'], bold=True)
    add_accent_line(slide, Inches(0.8), Inches(1.1), Inches(2))

    for i, item in enumerate(takeaways):
        y = Inches(1.5) + Inches(i * 0.85)
        if y > Inches(6.5):
            break
        add_number_badge(slide, Inches(0.8), y, i+1, color=COLORS['success'])
        add_text_box(slide, Inches(1.6), y, Inches(11), Inches(0.7),
                     item, font_size=16, color=COLORS['text_light'])

    add_footer(slide, module_num, module_name, page_num)


def make_end_slide(prs, next_module_title, module_num, module_name):
    """结束页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, COLORS['bg_dark'])

    add_shape(slide, Inches(0), Inches(3.2), SLIDE_WIDTH, Pt(2), fill_color=COLORS['accent'])
    add_text_box(slide, Inches(0), Inches(2.0), SLIDE_WIDTH, Inches(1.0),
                 "✅ 本模块结束", font_size=36, color=COLORS['text_white'], bold=True,
                 alignment=PP_ALIGN.CENTER)

    if next_module_title:
        add_text_box(slide, Inches(0), Inches(3.8), SLIDE_WIDTH, Inches(0.5),
                     f"下一模块：{next_module_title}", font_size=18, color=COLORS['text_light'],
                     alignment=PP_ALIGN.CENTER)

    add_text_box(slide, Inches(0), Inches(5.0), SLIDE_WIDTH, Inches(0.5),
                 "AI漫剧制作全流程课程 · 2026版", font_size=14, color=COLORS['text_dim'],
                 alignment=PP_ALIGN.CENTER)


def make_pr(prs):
    """设置16:9宽屏"""
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs


# ══════════════════════════════════════════════════════════
# 模块内容定义（每个模块的PPT页面内容）
# ══════════════════════════════════════════════════════════

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
            {'type': 'table', 'title': '行业数据一览',
             'headers': ['指标', '2026年数据'],
             'rows': [['市场规模', '240亿元'], ['用户规模', '2.8亿'], ['日均播放量', '15亿次'], ['创作者数量', '120万'], ['制作成本降幅', '-50%']]},
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
            {'type': 'key_point', 'key_point': 'AI漫剧的核心竞争力 = 创意 × 工具熟练度 × 量产能力',
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
                '利用AI技术（大语言模型、图像生成、视频生成、语音合成）辅助或主导制作的动态漫画短剧',
                '制作周期：传统1-3个月/集 → AI 1-3天/集（甚至1小时）',
                '制作成本：传统5-20万/集 → AI 500-5000元/集',
                '团队规模：传统10-30人 → AI 1-5人',
                '产能：传统2-4集/月 → AI 30+集/月'
            ]},
            {'type': 'content', 'title': '行业发展四个阶段', 'bullets': [
                '2023 · 沙雕漫：静态图片+配音+字幕，无动态效果',
                '2024 · 动态漫：图生视频初步应用，画面开始"动起来"',
                '2025 · AI原生漫剧：全链路AI辅助，单人可完成全流程',
                '2026 · AI仿真人剧：画面接近真人拍摄质量，与真人短剧直接竞争'
            ]},
            {'type': 'table', 'title': '2026年核心市场数据',
             'headers': ['指标', '2025年', '2026年', '增长率'],
             'rows': [['市场规模', '160亿', '240亿', '+50%'], ['用户规模', '2.0亿', '2.8亿', '+40%'], ['日均播放', '10亿次', '15亿次', '+50%'], ['创作者', '80万', '120万', '+50%']]},
            {'type': 'content', 'title': '五步工业化流程', 'bullets': [
                'Step 1 · 剧本创作：用AI（DeepSeek/文心）生成分集剧本',
                'Step 2 · 分镜设计：将剧本转化为分镜表，标注景别和提示词',
                'Step 3 · 文生图：用即梦AI生成每个镜头的关键帧画面',
                'Step 4 · 图生视频：用可灵AI/Seedance让静态画面动起来',
                'Step 5 · 剪辑配音：用剪映完成后期制作和发布'
            ]},
            {'type': 'content', 'title': '爆款案例拆解', 'bullets': [
                '《斩仙台》：玄幻题材，角色一致性标杆，单集播放破千万',
                '《气运三角洲》：都市逆袭，节奏把控精准，钩子设计教科书',
                '《霍去病》：历史题材，画面质感接近电影级别',
                '共同特点：前3秒强钩子、角色一致性好、节奏紧凑、情绪共鸣强'
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
        elif ptype == 'key_point':
            make_key_point_slide(prs, page['title'] if 'title' in page else '重点回顾',
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


# ══════════════════════════════════════════════════════════
# 主程序
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for m in MODULES:
        generate_module_ppt(m)
    print(f"\n🎉 已生成 {len(MODULES)} 个PPT文件")
