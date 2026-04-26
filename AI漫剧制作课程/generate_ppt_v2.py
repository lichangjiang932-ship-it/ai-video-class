#!/usr/bin/env python3
"""
AI漫剧制作全流程课程 PPT 生成器 v2
精美版 - 使用原生 zipfile + XML 构建完整 PPTX
"""

import zipfile
import os

# ──────────────────────────────────────────────
#  颜色方案
# ──────────────────────────────────────────────
C = {
    'bg_dark':   '0F1923',   # 深蓝黑背景
    'bg_mid':    '1A2A3A',   # 中等背景
    'accent':    'FF6B35',   # 橙色强调
    'accent2':   '00B4D8',   # 青色强调
    'accent3':   '1A936F',   # 绿色
    'accent4':   'F18F01',   # 金色
    'accent5':   'C73E1D',   # 红色
    'white':     'FFFFFF',
    'light':     'E8EDF5',   # 浅灰白
    'dim':       '8899AA',   # 暗灰
    'title':     'FFFFFF',
    'subtitle':  'FF6B35',
}

# ──────────────────────────────────────────────
#  XML 模板
# ──────────────────────────────────────────────

def content_types_xml(n):
    items = ''
    for i in range(1, n+1):
        items += f'  <Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
{items}</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''

def pres_rels_xml(n):
    items = f'''  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
'''
    for i in range(1, n+1):
        items += f'  <Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>\n'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{items}</Relationships>'''

def pres_xml(n):
    ids = ''
    for i in range(1, n+1):
        ids += f'    <p:sldId id="{255+i}" r:id="rId{i+2}"/>\n'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
{ids}  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''

THEME = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AI漫剧课程">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="0F1923"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1A2A3A"/></a:dk2>
      <a:lt2><a:srgbClr val="E8EDF5"/></a:lt2>
      <a:accent1><a:srgbClr val="FF6B35"/></a:accent1>
      <a:accent2><a:srgbClr val="00B4D8"/></a:accent2>
      <a:accent3><a:srgbClr val="1A936F"/></a:accent3>
      <a:accent4><a:srgbClr val="F18F01"/></a:accent4>
      <a:accent5><a:srgbClr val="C73E1D"/></a:accent5>
      <a:accent6><a:srgbClr val="3E1F47"/></a:accent6>
      <a:hlink><a:srgbClr val="00B4D8"/></a:hlink>
      <a:folHlink><a:srgbClr val="8899AA"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''

SLD_MASTER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>'''

SLD_MASTER_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''

SLD_LAYOUT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>'''

SLD_LAYOUT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

def slide_rels():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

# ──────────────────────────────────────────────
#  形状构建辅助函数
# ──────────────────────────────────────────────

def sp_rect(sp_id, name, x, y, cx, cy, fill_color, alpha=None):
    """矩形色块"""
    alpha_xml = ''
    if alpha is not None:
        alpha_xml = f'<a:alpha val="{alpha}"/>'
    return f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:solidFill><a:srgbClr val="{fill_color}">{alpha_xml}</a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
</p:sp>'''

def sp_line(sp_id, name, x, y, cx, cy, color, width=50000):
    """线条"""
    return f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
</p:sp>'''

def sp_text(sp_id, name, x, y, cx, cy, text, font_size=1800, color='FFFFFF', bold=False, align='l', anchor='t'):
    """文本框"""
    b_attr = ' b="1"' if bold else ''
    align_map = {'l': 'l', 'c': 'ctr', 'r': 'r'}
    a = align_map.get(align, 'l')
    return f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="{anchor}"/><a:lstStyle/>
    <a:p><a:pPr><a:algn>{a}</a:algn></a:pPr><a:r><a:rPr lang="zh-CN" sz="{font_size}"{b_attr} dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{_esc(text)}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''

def sp_bullets(sp_id, name, x, y, cx, cy, bullets, font_size=1400, color='E8EDF5', bullet_color='FF6B35', spacing=600):
    """带项目符号的文本"""
    items = ''
    for b in bullets:
        if not b:
            items += f'<a:p><a:pPr><a:spcBef><a:spcPts val="{spacing}"/></a:spcBef></a:pPr><a:r><a:rPr lang="zh-CN" sz="{font_size}" dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t></a:t></a:r></a:p>\n'
        else:
            items += f'<a:p><a:pPr><a:buFont typeface="Arial"/><a:buChar val="▸"/><a:buClr><a:srgbClr val="{bullet_color}"/></a:buClr><a:spcBef><a:spcPts val="{spacing}"/></a:spcBef></a:pPr><a:r><a:rPr lang="zh-CN" sz="{font_size}" dirty="0"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{_esc(b)}</a:t></a:r></a:p>\n'
    return f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/><a:lstStyle/>{items}</p:txBody>
</p:sp>'''

def sp_table(sp_id, name, x, y, cx, cy, rows, col_widths, header_color='FF6B35', row_color='1A2A3A', text_color='FFFFFF'):
    """简单表格（用文本框模拟）"""
    row_h = cy // len(rows)
    items = ''
    for ri, row in enumerate(rows):
        ry = y + ri * row_h
        bg = header_color if ri == 0 else (row_color if ri % 2 == 1 else '152535')
        items += sp_rect(sp_id + ri*100, f'row{ri}', x, ry, cx, row_h, bg)
        for ci, cell in enumerate(row):
            cx_pos = x + sum(col_widths[:ci])
            fw = col_widths[ci]
            fs = 1200 if ri == 0 else 1100
            b = True if ri == 0 else False
            items += sp_text(sp_id + ri*100 + ci + 10, f'cell{ri}_{ci}', cx_pos + 91440, ry, fw - 182880, row_h, cell, fs, text_color, b, 'l', 'ctr')
    return items

def _esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# ──────────────────────────────────────────────
#  幻灯片类型
# ──────────────────────────────────────────────

def slide_cover(title, subtitle, bg='0F1923'):
    """封面幻灯片"""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'topbar', 0, 0, 12192000, 120000, 'FF6B35')}
      {sp_rect(3, 'leftbar', 0, 0, 180000, 6858000, 'FF6B35')}
      {sp_rect(4, 'accent_block', 685800, 2200000, 500000, 800000, 'FF6B35', 30000)}
      {sp_text(5, 'title', 1400000, 2200000, 10000000, 1200000, title, 4800, 'FFFFFF', True, 'l', 'ctr')}
      {sp_line(6, 'divider', 1400000, 3500000, 3000000, 0, 'FF6B35', 60000)}
      {sp_text(7, 'subtitle', 1400000, 3700000, 9500000, 800000, subtitle, 2200, '8899AA', False, 'l', 't')}
      {sp_rect(8, 'bottombar', 0, 6658000, 12192000, 200000, '1A2A3A')}
      {sp_text(9, 'bottom_text', 685800, 6658000, 10820400, 200000, '2026版 · AI漫剧制作全流程课程', 1000, '8899AA', False, 'c', 'ctr')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def slide_section(title, module_num, accent='FF6B35'):
    """模块标题页"""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F1923"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'bg_accent', 0, 0, 4500000, 6858000, accent, 15000)}
      {sp_text(3, 'num', 685800, 1500000, 3200000, 1500000, f'模块 {module_num}', 7200, 'FFFFFF', True, 'l', 'ctr')}
      {sp_line(4, 'line', 685800, 3200000, 2500000, 0, 'FFFFFF', 40000)}
      {sp_text(5, 'label', 685800, 3400000, 3200000, 500000, 'MODULE', 1400, '8899AA', False, 'l', 't')}
      {sp_text(6, 'title_right', 5200000, 2000000, 6300000, 3000000, title, 3200, 'FFFFFF', True, 'l', 'ctr')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def slide_content(title, bullets, accent='FF6B35'):
    """内容幻灯片"""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F1923"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'top_accent', 0, 0, 12192000, 80000, accent)}
      {sp_rect(3, 'header_bg', 0, 80000, 12192000, 1100000, '1A2A3A')}
      {sp_text(4, 'title', 685800, 180000, 10820400, 900000, title, 3000, 'FFFFFF', True, 'l', 'ctr')}
      {sp_line(5, 'accent_line', 685800, 1230000, 2000000, 0, accent, 50000)}
      {sp_bullets(6, 'content', 685800, 1450000, 10820400, 5000000, bullets, 1600, 'E8EDF5', accent)}
      {sp_rect(7, 'bottom_bar', 0, 6658000, 12192000, 200000, '1A2A3A')}
      {sp_text(8, 'page', 10500000, 6658000, 1000000, 200000, '', 900, '8899AA', False, 'r', 'ctr')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def slide_table(title, headers, rows, accent='FF6B35'):
    """表格幻灯片"""
    n_cols = len(headers)
    total_w = 10820400
    col_w = total_w // n_cols
    col_widths = [col_w] * n_cols
    
    # Header row
    header_cells = ''
    for ci, h in enumerate(headers):
        cx = 685800 + ci * col_w
        header_cells += sp_rect(100 + ci, f'hdr{ci}', cx, 1450000, col_w - 20000, 600000, accent)
        header_cells += sp_text(200 + ci, f'htxt{ci}', cx + 91440, 1450000, col_w - 182880, 600000, h, 1300, 'FFFFFF', True, 'l', 'ctr')
    
    # Data rows
    data_cells = ''
    for ri, row in enumerate(rows):
        ry = 2100000 + ri * 650000
        bg = '1A2A3A' if ri % 2 == 0 else '152535'
        data_cells += sp_rect(300 + ri, f'rowbg{ri}', 685800, ry, total_w, 600000, bg)
        for ci, cell in enumerate(row):
            cx = 685800 + ci * col_w
            data_cells += sp_text(400 + ri * 10 + ci, f'cell{ri}_{ci}', cx + 91440, ry, col_w - 182880, 600000, cell, 1200, 'E8EDF5', False, 'l', 'ctr')
    
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F1923"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'top_accent', 0, 0, 12192000, 80000, accent)}
      {sp_rect(3, 'header_bg', 0, 80000, 12192000, 1100000, '1A2A3A')}
      {sp_text(4, 'title', 685800, 180000, 10820400, 900000, title, 3000, 'FFFFFF', True, 'l', 'ctr')}
      {sp_line(5, 'line', 685800, 1230000, 2000000, 0, accent, 50000)}
      {header_cells}
      {data_cells}
      {sp_rect(6, 'bottom_bar', 0, 6658000, 12192000, 200000, '1A2A3A')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def slide_two_col(title, left_title, left_bullets, right_title, right_bullets, accent='FF6B35'):
    """双栏幻灯片"""
    left_items = ''
    for b in left_bullets:
        left_items += f'<a:p><a:pPr><a:buChar val="▸"/><a:buClr><a:srgbClr val="{accent}"/></a:buClr><a:spcBef><a:spcPts val="500"/></a:spcBef></a:pPr><a:r><a:rPr lang="zh-CN" sz="1300" dirty="0"><a:solidFill><a:srgbClr val="E8EDF5"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{_esc(b)}</a:t></a:r></a:p>\n'
    right_items = ''
    for b in right_bullets:
        right_items += f'<a:p><a:pPr><a:buChar val="▸"/><a:buClr><a:srgbClr val="00B4D8"/></a:buClr><a:spcBef><a:spcPts val="500"/></a:spcBef></a:pPr><a:r><a:rPr lang="zh-CN" sz="1300" dirty="0"><a:solidFill><a:srgbClr val="E8EDF5"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{_esc(b)}</a:t></a:r></a:p>\n'
    
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F1923"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'top_accent', 0, 0, 12192000, 80000, accent)}
      {sp_rect(3, 'header_bg', 0, 80000, 12192000, 1100000, '1A2A3A')}
      {sp_text(4, 'title', 685800, 180000, 10820400, 900000, title, 3000, 'FFFFFF', True, 'l', 'ctr')}
      {sp_line(5, 'line', 685800, 1230000, 2000000, 0, accent, 50000)}
      {sp_rect(6, 'left_card', 457200, 1500000, 5400000, 4900000, '1A2A3A')}
      {sp_text(7, 'left_title', 650000, 1600000, 5000000, 500000, left_title, 1800, accent, True, 'l', 'ctr')}
      {sp_rect(8, 'left_bar', 650000, 2100000, 1500000, 40000, accent)}
      <p:sp><p:nvSpPr><p:cNvPr id="9" name="left_content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="650000" y="2250000"/><a:ext cx="5000000" cy="3900000"/></a:xfrm><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/><a:lstStyle/>{left_items}</p:txBody>
      </p:sp>
      {sp_rect(10, 'right_card', 6200000, 1500000, 5400000, 4900000, '1A2A3A')}
      {sp_text(11, 'right_title', 6400000, 1600000, 5000000, 500000, right_title, 1800, '00B4D8', True, 'l', 'ctr')}
      {sp_rect(12, 'right_bar', 6400000, 2100000, 1500000, 40000, '00B4D8')}
      <p:sp><p:nvSpPr><p:cNvPr id="13" name="right_content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6400000" y="2250000"/><a:ext cx="5000000" cy="3900000"/></a:xfrm><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
        <p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/><a:lstStyle/>{right_items}</p:txBody>
      </p:sp>
      {sp_rect(14, 'bottom_bar', 0, 6658000, 12192000, 200000, '1A2A3A')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

def slide_end(title, subtitle):
    """结尾幻灯片"""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F1923"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {sp_rect(2, 'accent_bg', 0, 2400000, 12192000, 2100000, 'FF6B35', 12000)}
      {sp_text(3, 'title', 685800, 2600000, 10820400, 1000000, title, 4400, 'FFFFFF', True, 'c', 'ctr')}
      {sp_line(4, 'line', 5000000, 3700000, 2192000, 0, 'FFFFFF', 40000)}
      {sp_text(5, 'subtitle', 685800, 3900000, 10820400, 600000, subtitle, 1800, '8899AA', False, 'c', 't')}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

# ──────────────────────────────────────────────
#  幻灯片数据
# ──────────────────────────────────────────────

slides = []

# 1. 封面
slides.append(('cover', {
    'title': 'AI漫剧制作全流程课程',
    'subtitle': '2026版 · 124课时 · 17个模块 · 从零基础到独立制作完整AI漫剧'
}))

# 2. 课程概览
slides.append(('content', {
    'title': '📊 课程概览',
    'bullets': [
        '课程名称：AI漫剧制作全流程实战课程 · 2026版',
        '总模块数：17个模块 · 总课时：124课时',
        '学习周期：8-12周（全日制）/ 16-24周（业余制）',
        '适用人群：零基础入门者、短视频创作者、漫画从业者、创业者',
        '前置要求：基本电脑操作能力，无需编程或美术基础',
        '行业背景：2026年市场规模240亿元，用户2.8亿，日均播放15亿次',
    ]
}))

# 3. 课程架构
slides.append(('content', {
    'title': '🏗️ 课程架构全景',
    'bullets': [
        '认知层（6课时）→ 模块1：行业认知与趋势',
        '创作层（18课时）→ 模块2-3：剧本创作 + 分镜设计',
        '图像层（16课时）→ 模块4-5：即梦AI + 其他图像工具',
        '视频层（22课时）→ 模块6-8：可灵AI + Seedance 2.0 + 海螺/Vidu',
        '核心技术（6课时）→ 模块9：角色一致性攻克',
        '音频层（8课时）→ 模块10：配音配乐与音画同步',
        '后期层（6课时）→ 模块11：剪辑与后期全流程',
        '平台层（16课时）→ 模块12-13：一站式平台 + ComfyUI',
        '工业化层（6课时）→ 模块14：工业化生产与团队协作',
        '商业层（12课时）→ 模块15-16：分发变现 + 合规规范',
        '实战层（8课时）→ 模块17：综合实战项目',
    ]
}))

# 4. 模块1
slides.append(('section', {'title': 'AI漫剧行业认知与趋势', 'module': '01'}))
slides.append(('content', {
    'title': '模块一：行业认知与趋势（6课时）',
    'bullets': [
        '🎯 2026年市场规模240亿元，用户2.8亿，创作者120万',
        '📈 四阶段演进：沙雕漫→动态漫→AI原生漫剧→AI仿真人剧',
        '🔄 五步工业化：剧本→分镜→文生图→图生视频→剪辑配音',
        '💡 爆款公式：精准题材 × 强节奏 × 高画质 × 稳更新 × 情感钩子',
        '🚀 三种创业路径：个人(1-5万) / 小团队(10-30万) / 工作室(50万+)',
        '📊 制作成本降幅50%（较2025年），日产30+集成为可能',
    ]
}))

# 5. 模块2
slides.append(('section', {'title': '大语言模型与AI剧本创作', 'module': '02'}))
slides.append(('content', {
    'title': '模块二：大语言模型与AI剧本创作（8课时）',
    'bullets': [
        '📝 五大AI模型深度使用：文心5.0 / DeepSeek / 豆包 / 通义千问 / 智谱清言',
        '🎯 AI剧本方法论：热点抓取→选题决策→故事构建→剧本落地',
        '🔥 DeepSeek杀手锏：联网搜索抓热点 + 长文本推理设计复杂剧情',
        '💡 可视化描写技巧：把抽象变成具象，用"摄影机视角"写剧本',
        '📐 AI适配要点：场景≤2-3人、动作可执行、包含完整视觉元素',
        '🎬 实战：用多AI工具协作完成12集漫剧完整剧本',
    ]
}))

# 6. 模块3
slides.append(('section', {'title': '分镜设计与提示词工程', 'module': '03'}))
slides.append(('content', {
    'title': '模块三：分镜设计与提示词工程（10课时）',
    'bullets': [
        '🎬 五大景别：远景→全景→中景→近景→特写，掌握切换节奏',
        '📐 六大构图法则：三分法 / 对角线 / 对称 / 框中框 / 前景遮挡 / 负空间',
        '✨ 提示词六要素：主体 + 动作 + 环境 + 光线 + 视角 + 风格',
        '📚 四大模板库：动作打斗 / 情感对话 / 情绪特写 / 场景展示',
        '🚫 负面提示词：系统化排除低质量/变形/模糊等AI常见问题',
        '🛠️ 工具链：Drawstory分镜 + DomoAI动态化 + 即梦AI智能多帧',
    ]
}))

# 7. 模块4
slides.append(('section', {'title': '即梦AI深度精通', 'module': '04'}))
slides.append(('content', {
    'title': '模块四：即梦AI深度精通（10课时）',
    'bullets': [
        '🖼️ 文生图：七步提示词法 + 参数调节(分辨率/风格/采样步数/CFG)',
        '🎨 图生图：风格迁移(0.6-0.7) + 精修(0.2-0.3) + 换背景(0.5-0.6)',
        '👤 角色一致性：角色参考功能 + 提示词锚定 + 图生图接力法',
        '🏛️ 场景资产库：核心场景"母版图" + 白天/黄昏/深夜变体批量生成',
        '🔀 多风格切换：日漫/国漫/赛博朋克/古风参数速查表',
        '🔌 企业级API：火山引擎接入 + 批量生产工作流',
    ]
}))

# 8. 模块5
slides.append(('section', {'title': '其他图像生成工具', 'module': '05'}))
slides.append(('content', {
    'title': '模块五：其他图像生成工具（6课时）',
    'bullets': [
        '🟠 通义万相(阿里)：中文理解强，中式美学风格突出',
        '🔵 腾讯混元：对话式迭代生成 + 数字人创作功能',
        '🔴 百度文心一格：国风古风专精，30+种预设艺术风格',
        '🟢 Kolors(快手)：开源免费，支持本地部署和深度定制',
        '🧠 LoRA训练：30-50张图 → 角色专属模型，彻底解决一致性',
        '🎯 决策矩阵：特写→SD+LoRA / 全景→MJ / 国风→文心一格 / 动作→SD+ControlNet',
    ]
}))

# 9. 模块6
slides.append(('section', {'title': '可灵AI深度精通', 'module': '06'}))
slides.append(('content', {
    'title': '模块六：可灵AI深度精通（8课时）',
    'bullets': [
        '🚀 可灵3.0：物理仿真引擎突破 + 影像级1080P/30fps输出',
        '🖼️ 图生视频：静态关键帧→动态画面，运动幅度低/中/高三档',
        '📝 文生视频：公式=主体+运动+场景+镜头+光影+氛围',
        '🎭 高级运动控制：运动笔刷(可视化轨迹) + 动作大师(骨骼级)',
        '🎯 首尾帧控制：精确设定起止状态，保障连续镜头衔接',
        '⚔️ 打斗实战：近身格斗/武器对决/超能力战斗提示词模板库',
    ]
}))

# 10. 模块7
slides.append(('section', {'title': 'Seedance 2.0深度精通', 'module': '07'}))
slides.append(('content', {
    'title': '模块七：Seedance 2.0深度精通（8课时）',
    'bullets': [
        '🎬 "导演之选"：多模态参考驱动 + 角色档案系统',
        '📹 参考视频驱动：上传真人动作视频→AI角色完美复刻',
        '📋 Draft模式：480P快速预览确认后再渲染1080P，省60%成本',
        '📈 抽卡率革命：从20%→90%+，四大策略提升可用率',
        '👤 角色档案：多角度参考图集 + 跨镜头形象锁定',
        '🔗 生态整合：与即梦AI无缝配合，关键帧→视频→剪辑全链路',
    ]
}))

# 11. 模块8
slides.append(('section', {'title': '海螺AI与Vidu', 'module': '08'}))
slides.append(('content', {
    'title': '模块八：海螺AI与Vidu（6课时）',
    'bullets': [
        '🌊 海螺AI 02(MiniMax)：高难度物理动作生成行业领先',
        '💃 复杂动作：打斗/舞蹈/极限运动等高难度场景最佳实践',
        '🎬 Vidu Q3：全球首个16秒音画直出，多角色对话支持',
        '👥 多角色同屏：2-3个角色自然对话和互动',
        '🎯 四大工具选型：对话→Vidu / 参考→Seedance / 打斗→海螺 / 普通→可灵',
    ]
}))

# 12. 模块9
slides.append(('section', {'title': '角色一致性攻克', 'module': '09'}))
slides.append(('content', {
    'title': '模块九：角色一致性攻克（6课时）',
    'bullets': [
        '🔥 最大挑战：AI每次独立生成，角色天然不一致',
        '💡 解决方案阶梯：固定提示词→参考图→IP-Adapter→LoRA微调',
        '🛠️ ComfyUI工作流：IP-Adapter FaceID Plus v2 + ControlNet精准控制',
        '🧠 LoRA训练：30-50张高质量图 → 训练角色专属模型',
        '🔧 ADetailer：自动检测面部偏移并高分辨率重渲染',
        '📦 资产库规范：多角度/多表情/多光照数据集标准化管理',
    ]
}))

# 13. 模块10
slides.append(('section', {'title': '配音、配乐与音画同步', 'module': '10'}))
slides.append(('content', {
    'title': '模块十：配音、配乐与音画同步（8课时）',
    'bullets': [
        '🎤 火山引擎TTS：中文语境深度适配，情感维度精准控制',
        '🎭 Fish Audio：表情标签系统（开心/悲伤/愤怒）精准控制',
        '✂️ 剪映AI克隆音色：30秒音频样本训练专属配音模型',
        '🎵 MiniMax Music-2.0：多风格音乐生成 + 风格混合',
        '🔊 SkyReels-V4：毫秒级口型对齐，实现精准音画同步',
        '💥 拟声字特效："轰""啊""哇"等文字特效制作技巧',
    ]
}))

# 14. 模块11
slides.append(('section', {'title': '剪辑与后期全流程', 'module': '11'}))
slides.append(('content', {
    'title': '模块十一：剪辑与后期全流程（6课时）',
    'bullets': [
        '🎯 前3秒法则：视觉冲击→冲突呈现→信息传达，层层递进',
        '🪝 黄金钩子：每集开头回顾+悬念，结尾必留钩子驱动追看',
        '🤖 剪映AI功能：智能字幕/自动踩点/AI混剪/智能修复',
        '🔧 画面修复：AI去闪烁/去变形/超分辨率/补帧',
        '🎨 调色统一：玄幻偏冷/都市自然/古风暖色/赛博霓虹',
        '📦 从素材到成片的完整剪辑工作流',
    ]
}))

# 15. 模块12
slides.append(('section', {'title': '一站式AI漫剧平台', 'module': '12'}))
slides.append(('content', {
    'title': '模块十二：一站式AI漫剧平台精讲（8课时）',
    'bullets': [
        '🏭 360纳米漫剧流水线：全链路自动化，零基础可用',
        '🎭 有戏AI：角色相似度95%+，一站式短剧创作平台',
        '🎬 商汤Seko 2.0：创编一体，连拍100集不崩',
        '📚 漫剧助手(阅文)：AI驱动+海量IP库，IP到漫剧最短路径',
        '🍊 橙星梦工厂 / 漫剧工场 / 白日梦AI 等平台横向对比',
        '🎯 选型建议：快速出片→360纳米 / 品质→有戏AI / IP→漫剧助手',
    ]
}))

# 16. 模块13
slides.append(('section', {'title': 'ComfyUI工作流', 'module': '13'}))
slides.append(('content', {
    'title': '模块十三：ComfyUI工作流进阶（8课时）',
    'bullets': [
        '⚙️ 环境搭建：Python 3.10+ / PyTorch 2.0+ / NVIDIA 8GB+ VRAM',
        '🔌 核心插件：IP-Adapter / ControlNet / ADetailer / Kolors',
        '🖼️ 文生图/图生图工作流：节点连接与参数优化详解',
        '🎮 ControlNet：骨骼姿态/深度图/线稿/边缘精准控制',
        '👤 角色一致性黄金组合：IP-Adapter + LoRA + ADetailer',
        '📚 Z-Comics：多格漫画批量生成插件实战',
    ]
}))

# 17. 模块14
slides.append(('section', {'title': '工业化生产与团队协作', 'module': '14'}))
slides.append(('content', {
    'title': '模块十四：工业化生产与团队协作（6课时）',
    'bullets': [
        '⚡ 极速流：1小时/集热点漫剧，单人日更模式',
        '💎 精品IP流：15-30天打造精品连载，团队协作模式',
        '👥 "1+4+2"团队：编剧 + 2抽卡师 + 后期 + 运营 + 美术 + 技术',
        '📦 资产沉淀复用：角色/场景/道具/模板资产库管理',
        '💰 成本控制：单集275-660元，Draft模式省60%试错成本',
        '📊 工业化评估矩阵：复杂度/质量/成本/时间四维决策',
    ]
}))

# 18. 模块15
slides.append(('section', {'title': '平台分发与商业变现', 'module': '15'}))
slides.append(('content', {
    'title': '模块十五：平台分发与商业变现（8课时）',
    'bullets': [
        '📱 抖音：完播率最重要，封面+前3秒决定生死，算法推荐',
        '🍎 红果短剧：分账模式成熟，S/A/B级评级决定收入',
        '⚡ 快手：可灵AI流量扶持，下沉市场优势明显',
        '🎮 B站：二次元氛围浓厚，高品质内容偏好',
        '🌏 出海变现：AI翻译+配音+换口型，全球化内容策略',
        '💰 变现阶梯：分账(1-10万)→付费(5-30万)→广告→IP授权→品牌联名(500万+)',
    ]
}))

# 19. 模块16
slides.append(('section', {'title': '合规要求与行业规范', 'module': '16'}))
slides.append(('content', {
    'title': '模块十六：合规要求与行业规范（4课时）',
    'bullets': [
        '📋 备案要求：所有AI漫剧需在平台完成备案登记',
        '🏷️ AIGC标识：必须在视频中标注"AI生成"，不得删除',
        '⚖️ 版权保护：纯AI生成不受著作权法保护，保留创作过程记录',
        '🚫 内容红线：暴力/色情/政治敏感/虚假信息绝对禁止',
        '🔒 数据安全：训练数据合规、用户隐私保护、API密钥安全',
        '🌱 可持续发展：从流量消耗→IP生态运营的长期主义转型',
    ]
}))

# 20. 模块17
slides.append(('section', {'title': '综合实战项目', 'module': '17'}))
slides.append(('content', {
    'title': '模块十七：综合实战项目（8课时）',
    'bullets': [
        'Step 1：选题定位 — 市场调研 + 竞品分析 + 蓝海判断',
        'Step 2：剧本创作 — DeepSeek完成3-5集剧本 + 分镜脚本',
        'Step 3：角色资产 — 即梦AI + ComfyUI建立完整角色资产库',
        'Step 4：图像生成 — 批量生成所有关键帧画面',
        'Step 5：视频生成 — Seedance 2.0 + 可灵AI完成动态镜头',
        'Step 6：音频制作 — 火山引擎配音 + BGM + 音效',
        'Step 7：剪辑成片 — 剪映专业版精剪 + 调色 + 字幕',
        'Step 8：发布运营 — 多平台分发 + 数据分析 + 迭代优化',
    ]
}))

# 21. 核心工具速查
slides.append(('two_col', {
    'title': '🛠️ 核心工具速查表',
    'left_title': '创作工具',
    'left_bullets': [
        '剧本：DeepSeek（首选）+ 文心5.0',
        '分镜：Drawstory + 即梦AI',
        '图像：即梦AI（首选）+ 通义万相',
        '视频：可灵AI + Seedance 2.0',
        '配音：火山引擎TTS + Fish Audio',
        '配乐：MiniMax Music-2.0',
    ],
    'right_title': '进阶工具',
    'right_bullets': [
        'ComfyUI：专业工作流引擎',
        '海螺AI：高难度物理动作',
        'Vidu Q3：16秒音画直出',
        'Kolors：开源图像生成',
        'LoRA训练：角色一致性终极方案',
        '剪映：全流程后期制作',
    ],
}))

# 22. 市场数据表
slides.append(('table', {
    'title': '📊 2026年AI漫剧市场数据',
    'headers': ['指标', '2025年', '2026年', '增长率'],
    'rows': [
        ['市场规模', '120亿元', '240亿元', '+100%'],
        ['用户规模', '1.8亿', '2.8亿', '+55%'],
        ['日均播放量', '8亿次', '15亿次', '+87%'],
        ['创作者数量', '50万', '120万', '+140%'],
        ['平均制作成本/集', '3000元', '1500元', '-50%'],
        ['头部作品月收入', '50万', '120万', '+140%'],
    ],
}))

# 23. 变现路径
slides.append(('content', {
    'title': '💰 变现路径与收入预期',
    'bullets': [
        'Level 1：平台分账 → 月入1-10万（入门级，抖音/红果/B站）',
        'Level 2：付费解锁 → 月入5-30万（进阶级，付费卡点设计）',
        'Level 3：广告植入 → 月入10-50万（品牌合作，剧情软植入）',
        'Level 4：IP授权 → 年入100万+（周边/游戏/影视改编）',
        'Level 5：品牌联名 → 年入500万+（跨界合作，联名产品）',
        '📊 收入来源占比：平台分账45% / 付费解锁30% / 广告15% / 其他10%',
    ]
}))

# 24. 学习路径
slides.append(('content', {
    'title': '🎯 学习路径选择',
    'bullets': [
        '⚡ 零基础速成（8周）：模块1→2→3→4→6→10→11→12→17',
        '    → 目标：快速上手，能独立产出基础漫剧',
        '💎 专业进阶（16周）：全部17个模块按顺序学习',
        '    → 目标：掌握全链路，能做精品漫剧',
        '🚀 创业导向（12周）：模块1→2→4→6/7→9→10→11→14→15→16→17',
        '    → 目标：快速建立团队，投入商业化生产',
    ]
}))

# 25. 结尾
slides.append(('end', {
    'title': '开始你的AI漫剧创作之旅！',
    'subtitle': '🎬 124课时 · 17个模块 · 从零到一 · 全流程实战 · 2026版'
}))

# ──────────────────────────────────────────────
#  生成 PPTX
# ──────────────────────────────────────────────

def build_pptx(slides, output_path):
    n = len(slides)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # 基础结构
        zf.writestr('[Content_Types].xml', content_types_xml(n))
        zf.writestr('_rels/.rels', ROOT_RELS)
        zf.writestr('ppt/presentation.xml', pres_xml(n))
        zf.writestr('ppt/_rels/presentation.xml.rels', pres_rels_xml(n))
        zf.writestr('ppt/theme/theme1.xml', THEME)
        zf.writestr('ppt/slideMasters/slideMaster1.xml', SLD_MASTER)
        zf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', SLD_MASTER_RELS)
        zf.writestr('ppt/slideLayouts/slideLayout1.xml', SLD_LAYOUT)
        zf.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', SLD_LAYOUT_RELS)
        
        # 生成每张幻灯片
        for i, (stype, data) in enumerate(slides, 1):
            if stype == 'cover':
                xml = slide_cover(data['title'], data['subtitle'])
            elif stype == 'section':
                xml = slide_section(data['title'], data['module'])
            elif stype == 'content':
                xml = slide_content(data['title'], data['bullets'])
            elif stype == 'two_col':
                xml = slide_two_col(data['title'], data['left_title'], data['left_bullets'], data['right_title'], data['right_bullets'])
            elif stype == 'table':
                xml = slide_table(data['title'], data['headers'], data['rows'])
            elif stype == 'end':
                xml = slide_end(data['title'], data['subtitle'])
            else:
                xml = slide_content('Untitled', [''])
            
            zf.writestr(f'ppt/slides/slide{i}.xml', xml)
            zf.writestr(f'ppt/slides/_rels/slide{i}.xml.rels', slide_rels())
    
    print(f"✅ PPTX 已生成: {output_path}")
    print(f"   幻灯片数量: {n}")
    print(f"   文件大小: {os.path.getsize(output_path) / 1024:.1f} KB")

if __name__ == '__main__':
    output_dir = '/root/.openclaw/workspace/ai-video-class/AI漫剧制作课程'
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, 'AI漫剧制作全流程课程-2026版.pptx')
    build_pptx(slides, output_path)
