#!/usr/bin/env python3
"""Generate AI漫剧制作课程 PPT from scratch using zipfile + XML."""

import zipfile
import os
import io

# PPTX is a ZIP with specific XML structure
# We'll create a minimal but functional PPTX

def create_pptx(output_path, slides_data):
    """Create a PPTX file with the given slides data."""
    
    # --- Content Types ---
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
'''
    for i in range(1, len(slides_data) + 1):
        content_types += f'  <Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n'
    content_types += '</Types>'

    # --- Root relationships ---
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''

    # --- Presentation relationships ---
    pres_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
'''
    for i in range(1, len(slides_data) + 1):
        pres_rels += f'  <Relationship Id="rId{i+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>\n'
    pres_rels += '</Relationships>'

    # --- Presentation XML ---
    slide_ids = ''
    for i in range(1, len(slides_data) + 1):
        slide_ids += f'    <p:sldId id="{255+i}" r:id="rId{i+2}"/>\n'
    
    presentation = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
{slide_ids}  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>'''

    # --- Theme ---
    theme = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AI漫剧课程">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="1B2A4A"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="2D3E5F"/></a:dk2>
      <a:lt2><a:srgbClr val="E8EDF5"/></a:lt2>
      <a:accent1><a:srgbClr val="FF6B35"/></a:accent1>
      <a:accent2><a:srgbClr val="004E89"/></a:accent2>
      <a:accent3><a:srgbClr val="1A936F"/></a:accent3>
      <a:accent4><a:srgbClr val="F18F01"/></a:accent4>
      <a:accent5><a:srgbClr val="C73E1D"/></a:accent5>
      <a:accent6><a:srgbClr val="3E1F47"/></a:accent6>
      <a:hlink><a:srgbClr val="004E89"/></a:hlink>
      <a:folHlink><a:srgbClr val="6B705C"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont>
      <a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>'''

    # --- Slide Master ---
    slide_master = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>'''

    slide_master_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>'''

    # --- Slide Layout ---
    slide_layout = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>'''

    slide_layout_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

    def make_slide_xml(title, bullets, bg_color="1B2A4A", title_color="FFFFFF", text_color="E8EDF5", accent_color="FF6B35", is_cover=False):
        """Generate slide XML with title and bullet points."""
        
        if is_cover:
            # Cover slide - centered title with subtitle
            return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="1714500"/><a:ext cx="10820400" cy="2057400"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="zh-CN" sz="4400" b="1" dirty="0"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{title}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Subtitle"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="685800" y="3886200"/><a:ext cx="10820400" cy="1143000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="zh-CN" sz="2000" dirty="0"><a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{bullets[0] if bullets else ""}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="AccentLine"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="5000000" y="3700000"/><a:ext cx="2192000" cy="50000"/></a:xfrm>
          <a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''
        
        # Content slide with bullets
        bullet_text = ""
        for i, bullet in enumerate(bullets):
            sz = 1800 if i == 0 else 1600
            bullet_text += f'<a:p><a:pPr><a:buChar val="●"/><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr><a:r><a:rPr lang="zh-CN" sz="{sz}" dirty="0"><a:solidFill><a:srgbClr val="{text_color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{bullet}</a:t></a:r></a:p>\n'
        
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="274320"/><a:ext cx="11277600" cy="1143000"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr lang="zh-CN" sz="3200" b="1" dirty="0"><a:solidFill><a:srgbClr val="{title_color}"/></a:solidFill><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr><a:t>{title}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="AccentBar"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="1417320"/><a:ext cx="1800000" cy="50000"/></a:xfrm>
          <a:solidFill><a:srgbClr val="{accent_color}"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="11277600" cy="4800600"/></a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
          <a:lstStyle/>
          {bullet_text}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>'''

    def make_slide_rels():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

    # --- Write PPTX ---
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', root_rels)
        zf.writestr('ppt/presentation.xml', presentation)
        zf.writestr('ppt/_rels/presentation.xml.rels', pres_rels)
        zf.writestr('ppt/theme/theme1.xml', theme)
        zf.writestr('ppt/slideMasters/slideMaster1.xml', slide_master)
        zf.writestr('ppt/slideMasters/_rels/slideMaster1.xml.rels', slide_master_rels)
        zf.writestr('ppt/slideLayouts/slideLayout1.xml', slide_layout)
        zf.writestr('ppt/slideLayouts/_rels/slideLayout1.xml.rels', slide_layout_rels)
        
        for i, slide_data in enumerate(slides_data, 1):
            title = slide_data.get('title', '')
            bullets = slide_data.get('bullets', [])
            bg = slide_data.get('bg', '1B2A4A')
            is_cover = slide_data.get('is_cover', False)
            accent = slide_data.get('accent', 'FF6B35')
            
            slide_xml = make_slide_xml(title, bullets, bg_color=bg, accent_color=accent, is_cover=is_cover)
            zf.writestr(f'ppt/slides/slide{i}.xml', slide_xml)
            zf.writestr(f'ppt/slides/_rels/slide{i}.xml.rels', make_slide_rels())

    print(f"✅ PPTX created: {output_path}")
    print(f"   Slides: {len(slides_data)}")


# ============================================================
# Define all slides
# ============================================================

slides = [
    # Cover
    {
        'title': 'AI漫剧制作全流程课程',
        'bullets': ['2026版 · 124课时 · 17个模块 · 从零基础到独立制作'],
        'is_cover': True,
        'bg': '1B2A4A',
        'accent': 'FF6B35'
    },
    # TOC
    {
        'title': '📋 课程目录',
        'bullets': [
            '模块1-3：行业认知 + 剧本创作 + 分镜设计（24课时）',
            '模块4-5：图像生成——即梦AI + 其他工具（16课时）',
            '模块6-8：视频生成——可灵AI + Seedance + 海螺/Vidu（22课时）',
            '模块9：角色一致性——核心技术攻克（6课时）',
            '模块10-11：音频制作 + 剪辑后期（14课时）',
            '模块12-13：一站式平台 + ComfyUI工作流（16课时）',
            '模块14-16：工业化生产 + 分发变现 + 合规（18课时）',
            '模块17：综合实战——完整漫剧项目（8课时）',
        ]
    },
    # Module 1
    {
        'title': '模块一：AI漫剧行业认知与趋势',
        'bullets': [
            '🎯 2026年市场规模：240亿元，用户2.8亿',
            '📈 内容形态演进：沙雕漫→动态漫→AI原生漫剧→AI仿真人剧',
            '🔄 五步工业化革命：剧本→分镜→文生图→图生视频→剪辑配音',
            '💡 爆款公式：精准题材 × 强节奏 × 高画质 × 稳更新 × 情感钩子',
            '🚀 创业路径：个人创作者(1-5万) / 小团队(10-30万) / 工作室(50万+)',
        ]
    },
    # Module 2
    {
        'title': '模块二：大语言模型与AI剧本创作',
        'bullets': [
            '📝 五大模型：文心5.0 / DeepSeek / 豆包 / 通义千问 / 智谱清言',
            '🎯 剧本方法论：热点抓取→爽点设计→可视化描写→镜头语言适配',
            '🔥 DeepSeek核心用法：联网搜索 + 长文本推理 + 多轮对话',
            '📊 实战项目：用DeepSeek+文心5.0完成12集漫剧完整剧本',
            '💡 提示词模板：[题材]+[角色]+[冲突]+[爽点]+[集数]一键生成大纲',
        ]
    },
    # Module 3
    {
        'title': '模块三：分镜设计与提示词工程',
        'bullets': [
            '🎬 分镜基础：景别(远/全/中/近/特) + 构图 + 镜头运动',
            '✨ 提示词六要素：主体描述、动作表情、环境、光线、视角、风格',
            '📚 模板库：动作打斗 / 情感对话 / 情绪特写 / 场景展示四大类型',
            '🚫 负面提示词：系统化排除低质量/变形/模糊等常见问题',
            '🛠️ 工具：Drawstory分镜 + DomoAI动态分镜 + 即梦AI多帧分镜',
        ]
    },
    # Module 4
    {
        'title': '模块四：即梦AI深度精通',
        'bullets': [
            '🖼️ 文生图：提示词编写规范 + 参数调节(分辨率/风格/采样步数)',
            '🎨 图生图：风格迁移操作 + 画面精修技巧',
            '👤 角色一致性：角色参考功能 + 多镜头形象稳定方案',
            '🏛️ 场景资产库：核心场景"母版图"批量生成方法',
            '🔀 多风格切换：国漫/日漫/赛博朋克/古风参数设置',
            '🔌 火山引擎API：企业级批量生产接入方式',
        ]
    },
    # Module 5
    {
        'title': '模块五：其他图像生成工具',
        'bullets': [
            '🟠 通义万相(阿里)：文生图/图生图，中式美学优势',
            '🔵 腾讯混元：图像生成 + 数字人创作',
            '🔴 百度文心一格：AI绘画 + 中式美学',
            '🟢 Kolors(快手可图)：开源模型，漫剧场景应用',
            '🧠 角色LoRA训练：ComfyUI环境 + 30-50张图训练专属模型',
            '🎯 决策矩阵：根据镜头类型选择最优工具',
        ]
    },
    # Module 6
    {
        'title': '模块六：可灵AI深度精通',
        'bullets': [
            '🚀 可灵AI 3.0：物理仿真突破 + 影像级1080P输出',
            '🖼️ 图生视频：静态图→动态画面，参数设置详解',
            '📝 文生视频：纯文本描述驱动视频生成',
            '🎭 高级运动控制：复杂人体动力学 + 动作大师能力',
            '🎯 首尾帧控制：精准设定起始帧与结束帧过渡',
            '⚔️ 打斗场景实战：武打/运动/高难度动作提示词模板',
        ]
    },
    # Module 7
    {
        'title': '模块七：Seedance 2.0深度精通',
        'bullets': [
            '🎬 "导演之选"：多模态参考系统 + 角色档案',
            '📹 参考驱动：上传参考视频→AI角色复刻动作',
            '📋 Draft模式：480P预览确认后再渲染高清，节省60%成本',
            '📈 抽卡率革命：从20%→90%+可用率的四大策略',
            '🔗 与即梦AI整合：全链路"AI工作室"实战',
            '⚔️ 武侠漫剧：真人功夫视频驱动漫剧侠客动作',
        ]
    },
    # Module 8
    {
        'title': '模块八：海螺AI与Vidu',
        'bullets': [
            '🌊 海螺AI 02(MiniMax)：高难度物理动作生成突破',
            '💃 复杂动作：打斗/舞蹈/运动场景最佳实践',
            '🎬 Vidu Q3：全球首个16秒音画直出模型',
            '👥 多角色对话：支持2-3个角色同屏对话',
            '🎯 选型决策：对话→Vidu / 参考→Seedance / 打斗→海螺 / 普通→可灵',
        ]
    },
    # Module 9
    {
        'title': '模块九：角色一致性——核心技术攻克',
        'bullets': [
            '🔥 最大挑战：AI每次生成都是"独立"的，角色天然不一致',
            '💡 解决方案阶梯：参考图→IP-Adapter→LoRA微调',
            '🛠️ ComfyUI工作流：IP-Adapter FaceID Plus v2 + ControlNet',
            '🧠 LoRA训练：30-50张图训练角色专属模型',
            '🔧 ADetailer：自动检测+高分辨率重渲染面部偏移',
            '📦 资产库管理：多角度/多表情/多光照数据集规范',
        ]
    },
    # Module 10
    {
        'title': '模块十：AI音频——配音、配乐与音画同步',
        'bullets': [
            '🎤 火山引擎TTS：中文语境深度适配，情感控制',
            '🎭 Fish Audio：表情标签系统精准控制情感',
            '✂️ 剪映AI克隆音色：30秒训练专属配音模型',
            '🎵 MiniMax Music-2.0：AI音乐生成+风格混合',
            '🔊 SkyReels-V4：毫秒级口型对齐音画同步',
            '💥 拟声字特效："轰""啊""哇"等文字特效制作',
        ]
    },
    # Module 11
    {
        'title': '模块十一：剪辑与后期——剪映专业版',
        'bullets': [
            '🎯 前3秒法则：视觉冲击→冲突呈现→信息传达',
            '🪝 黄金钩子：每集开头回顾+悬念，结尾必留钩子',
            '🤖 AI功能：智能字幕/自动踩点/AI混剪/智能修复',
            '🔧 画面修复：AI去闪烁/去变形/超分辨率',
            '🎨 调色统一：玄幻偏冷/都市自然/古风暖色/赛博霓虹',
        ]
    },
    # Module 12
    {
        'title': '模块十二：一站式AI漫剧平台',
        'bullets': [
            '🏭 360纳米漫剧流水线：全链路自动化，零基础可用',
            '🎭 有戏AI：角色相似度95%+，一站式短剧创作',
            '🎬 商汤Seko 2.0：创编一体，连拍100集不崩',
            '📚 漫剧助手(阅文)：AI驱动+海量IP，IP到漫剧最短路径',
            '🎯 选型建议：快速出片→360纳米 / 品质→有戏AI / IP→漫剧助手',
        ]
    },
    # Module 13
    {
        'title': '模块十三：ComfyUI工作流进阶',
        'bullets': [
            '⚙️ 环境搭建：Python 3.10+ / NVIDIA 8GB+ VRAM',
            '🔌 核心插件：IP-Adapter / ControlNet / ADetailer',
            '🖼️ 文生图/图生图工作流：节点连接与参数优化',
            '🎮 ControlNet：骨骼姿态/深度图/线稿精准控制',
            '👤 角色一致性：IP-Adapter + LoRA + ADetailer黄金组合',
            '📚 Z-Comics：多格漫画批量生成插件',
        ]
    },
    # Module 14
    {
        'title': '模块十四：工业化生产与团队协作',
        'bullets': [
            '⚡ 极速流：1小时/集热点漫剧（单人日更模式）',
            '💎 精品IP流：15-30天打造精品连载（团队模式）',
            '👥 "1+4+2"模式：编剧+2抽卡师+后期+运营+美术+技术',
            '📦 资产沉淀：角色/场景/道具/模板资产库复用',
            '💰 成本控制：单集275-660元，Draft模式省60%',
        ]
    },
    # Module 15
    {
        'title': '模块十五：平台分发与商业变现',
        'bullets': [
            '📱 抖音：完播率最重要，封面+前3秒决定生死',
            '🍎 红果短剧：分账模式成熟，评级决定收入',
            '⚡ 快手：可灵AI流量扶持，下沉市场优势',
            '🎮 B站：二次元氛围，高品质内容偏好',
            '🌏 出海变现：AI翻译+配音+换口型全球化策略',
            '💰 变现阶梯：分账→付费→广告→IP授权→品牌联名',
        ]
    },
    # Module 16
    {
        'title': '模块十六：合规要求与行业规范',
        'bullets': [
            '📋 备案要求：所有AI漫剧需在平台备案',
            '🏷️ AIGC标识：必须标注"AI生成"，不得删除',
            '⚖️ 版权：纯AI生成不受保护，保留创作过程记录',
            '🚫 内容红线：暴力/色情/政治敏感绝对禁止',
            '🌱 可持续发展：从流量消耗→IP生态运营转型',
        ]
    },
    # Module 17
    {
        'title': '模块十七：综合实战——完整漫剧项目',
        'bullets': [
            'Step 1: 选题定位（市场调研+竞品分析）',
            'Step 2: 剧本创作（DeepSeek完成3-5集剧本+分镜脚本）',
            'Step 3: 角色资产（即梦AI+ComfyUI建立资产库）',
            'Step 4: 图像生成（批量生成所有关键帧画面）',
            'Step 5: 视频生成（Seedance+可灵AI完成动态镜头）',
            'Step 6: 音频制作（配音+BGM+音效）',
            'Step 7: 剪辑成片（剪映专业版精剪+调色+字幕）',
            'Step 8: 发布运营（多平台分发+数据分析+迭代优化）',
        ]
    },
    # Tools Summary
    {
        'title': '🛠️ 核心工具速查表',
        'bullets': [
            '📝 剧本：DeepSeek（首选）+ 文心5.0 + 豆包',
            '🎬 分镜：Drawstory + 即梦AI + DomoAI',
            '🖼️ 图像：即梦AI（首选）+ 通义万相 + ComfyUI',
            '🎥 视频：可灵AI + Seedance 2.0 + 海螺AI + Vidu Q3',
            '🎤 配音：火山引擎TTS + Fish Audio + 剪映AI克隆',
            '🎵 配乐：MiniMax Music-2.0',
            '✂️ 剪辑：剪映专业版（全流程）',
        ]
    },
    # Business Model
    {
        'title': '💰 变现路径与收入预期',
        'bullets': [
            'Level 1: 平台分账 → 月入1-10万（入门级）',
            'Level 2: 付费解锁 → 月入5-30万（进阶级）',
            'Level 3: 广告植入 → 月入10-50万（品牌合作）',
            'Level 4: IP授权 → 年入100万+（周边/游戏/影视）',
            'Level 5: 品牌联名 → 年入500万+（跨界合作）',
            '📊 成本：单集275-660元 | 头部作品月收入：120万',
        ]
    },
    # Learning Paths
    {
        'title': '🎯 学习路径选择',
        'bullets': [
            '⚡ 零基础速成(8周)：模块1→2→3→4→6→10→11→12→17',
            '💎 专业进阶(16周)：全部17个模块按顺序学习',
            '🚀 创业导向(12周)：模块1→2→4→6/7→9→10→11→14→15→16→17',
            '',
            '✅ 完成课程后，你将掌握：',
            '   AI漫剧完整制作流程 + 主流AI工具使用',
            '   工业化生产能力 + 平台运营变现策略',
        ]
    },
    # End
    {
        'title': '开始你的AI漫剧创作之旅！',
        'bullets': ['🎬 124课时 · 17个模块 · 从零到一 · 全流程实战', '2026版 AI漫剧制作全流程课程'],
        'is_cover': True,
        'bg': '1B2A4A',
        'accent': 'FF6B35'
    },
]

# Generate
output_dir = '/root/.openclaw/workspace/AI漫剧制作课程'
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'AI漫剧制作全流程课程-2026版.pptx')
create_pptx(output_path, slides)
