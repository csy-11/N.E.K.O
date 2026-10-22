#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成演示样例（模拟 Windows 迁移到 Linux/WPS 后出现文字错位的典型演示文稿）。

生成两个文件：
  sample_broken.pptx   包含各类跨平台错位隐患（缺字体、自动缩放残留、
                       百分比行距、不换行文本框、表格边距缺失、竖排、组合缩放……）
  sample_clean.pptx    同一份内容，但按本工具的修复策略写好了显式属性

用途：不想拿真实课件冒险时，可以先用样例体验完整的“扫描 → 修复 → 复查”流程。
只依赖 Python 标准库。用法：
    python3 make_sample.py [输出目录]
"""

import os
import sys
import zipfile

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

DECL = (
    'xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"' % (NS_A, NS_R, NS_P)
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\
<Default Extension="xml" ContentType="application/xml"/>\
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>\
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>\
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>\
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>\
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>\
</Relationships>"""

CORE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" \
xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" \
xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
<dc:title>跨平台演示文稿样例</dc:title><dc:creator>fixture</dc:creator><cp:lastModifiedBy>fixture</cp:lastModifiedBy>\
<dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>\
<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>\
</cp:coreProperties>"""

APP = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" \
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
<Application>Microsoft Office PowerPoint</Application><PresentationFormat>宽屏</PresentationFormat>\
<Slides>2</Slides><Company></Company><AppVersion>16.0000</AppVersion></Properties>"""

THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="%s" name="Office 主题">\
<a:themeElements>\
<a:clrScheme name="Office">\
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>\
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
<a:dk2><a:srgbClr val="44546A"/></a:dk2>\
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>\
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>\
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>\
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>\
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>\
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>\
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>\
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>\
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>\
</a:clrScheme>\
<a:fontScheme name="Office">\
<a:majorFont>\
<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>\
<a:ea typeface=""/>\
<a:cs typeface=""/>\
<a:font script="Hans" typeface="等线 Light"/>\
</a:majorFont>\
<a:minorFont>\
<a:latin typeface="Calibri" panose="020F0502020204030204"/>\
<a:ea typeface=""/><a:cs typeface=""/>\
<a:font script="Hans" typeface="等线"/>\
</a:minorFont>\
</a:fontScheme>\
<a:fmtScheme name="Office">\
<a:fillStyleLst>\
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>\
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>\
</a:fillStyleLst>\
<a:lnStyleLst>\
<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\
<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\
<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>\
</a:lnStyleLst>\
<a:effectStyleLst>\
<a:effectStyle><a:effectLst/></a:effectStyle>\
<a:effectStyle><a:effectLst/></a:effectStyle>\
<a:effectStyle><a:effectLst/></a:effectStyle>\
</a:effectStyleLst>\
<a:bgFillStyleLst>\
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>\
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>\
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>\
</a:bgFillStyleLst>\
</a:fmtScheme>\
</a:themeElements>\
<a:objectDefaults/><a:extraClrSchemeLst/>\
</a:theme>""" % NS_A


def presentation(clean=False):
    if clean:
        default_style = (
            '<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN" altLang="en-US" sz="1800"/></a:defPPr>'
            '<a:lvl1pPr marL="0" indent="0" algn="l"><a:defRPr sz="1800" lang="zh-CN" altLang="en-US">'
            '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>'
            '<a:cs typeface="Liberation Sans"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>'
        )
    else:
        default_style = (
            '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US" sz="1800"/></a:defPPr>'
            '<a:lvl1pPr marL="0" indent="0" algn="l"><a:defRPr sz="1800" lang="en-US">'
            '<a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:presentation %s>'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId4"/></p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '%s'
        '</p:presentation>' % (DECL, default_style)
    )


PRESENTATION_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>\
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>\
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>\
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>\
</Relationships>"""


def master(clean=False):
    if clean:
        fonts = ('<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>'
                 '<a:cs typeface="Liberation Sans"/>')
    else:
        fonts = ('<a:latin typeface="Calibri"/><a:ea typeface=""/>'
                 '<a:cs typeface="Calibri"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sldMaster %s>'
        '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
        'hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:txStyles>'
        '<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400" b="1" lang="zh-CN" altLang="en-US">%s</a:defRPr></a:lvl1pPr></p:titleStyle>'
        '<p:bodyStyle><a:lvl1pPr marL="0" indent="0"><a:defRPr sz="2000" lang="zh-CN" altLang="en-US">%s</a:defRPr></a:lvl1pPr></p:bodyStyle>'
        '<p:otherStyle><a:defPPr><a:defRPr lang="zh-CN" altLang="en-US"/></a:defPPr></p:otherStyle>'
        '</p:txStyles>'
        '</p:sldMaster>' % (DECL, fonts, fonts)
    )


MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>\
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>\
</Relationships>"""


def _ph_sp(idx, name, ph_type, x, y, cx, cy, body, clean):
    if clean:
        bodypr = ('<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" anchor="t" wrap="square"/>')
    else:
        bodypr = '<a:bodyPr/>'
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="%s" idx="%d"/></p:nvPr></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/>%s</p:txBody></p:sp>'
        % (idx, name, ph_type, idx, x, y, cx, cy, bodypr, body)
    )


def layout(clean=False):
    title = _ph_sp(1, "标题 1", "title", 838200, 365125, 10515600, 1325563,
                   '<a:p><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>标题占位符</a:t></a:r></a:p>', clean)
    body = _ph_sp(2, "内容占位符 2", "body", 838200, 1825625, 10515600, 4351338,
                  '<a:p><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>正文占位符</a:t></a:r></a:p>', clean)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sldLayout %s type="titleAndBody" preserve="1">'
        '<p:cSld name="标题和内容"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '%s%s'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>' % (DECL, title, body)
    )


LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>\
</Relationships>"""

SLIDE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>\
</Relationships>"""


def _run(text, sz, latin=None, ea=None, lang="zh-CN", spc=None):
    attrs = ' lang="%s" altLang="en-US" sz="%d"' % (lang, sz)
    if spc is not None:
        attrs += ' spc="%d"' % spc
    font_xml = ""
    if latin:
        font_xml += '<a:latin typeface="%s"/>' % latin
    if ea:
        font_xml += '<a:ea typeface="%s"/>' % ea
    if not font_xml:
        font_xml = ""
    return '<a:r><a:rPr%s>%s</a:rPr><a:t>%s</a:t></a:r>' % (attrs, font_xml, text)


def slide1(clean=False):
    if clean:
        title_font = ('<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>')
        title_size = 4000
        body_pr = '<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" anchor="t" wrap="square"><a:noAutofit/></a:bodyPr>'
        box_pr = '<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" wrap="square"/>'
        box2_pr = '<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" wrap="square"/>'
        font_title = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>'
        font_body = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>'
        font_box = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/>'
        tblpr = ('<a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>')
        tcpr = '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="ctr"/>'
        trh = 400000
        grp = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="组合 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
               '<p:grpSpPr><a:xfrm><a:off x="762000" y="5200000"/><a:ext cx="4572000" cy="914400"/>'
               '<a:chOff x="0" y="0"/><a:chExt cx="4572000" cy="914400"/></a:xfrm></p:grpSpPr>'
               '<p:sp><p:nvSpPr><p:cNvPr id="10" name="组内文本框 10"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>'
               '<p:txBody><a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" wrap="square"/>'
               '<a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="1800">%s</a:rPr>'
               '<a:t>组合内文字</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>' % font_box)
    else:
        title_font = '<a:latin typeface="微软雅黑"/>'
        title_size = 4000
        body_pr = ('<a:bodyPr><a:normAutofit fontScale="92500" lnSpcReduction="20000"/></a:bodyPr>')
        box_pr = '<a:bodyPr wrap="none"/>'
        box2_pr = '<a:bodyPr><a:spAutoFit/></a:bodyPr>'
        font_title = '<a:latin typeface="微软雅黑"/>'
        font_body = '<a:latin typeface="宋体"/><a:ea typeface="宋体"/>'
        font_box = '<a:latin typeface="Times New Roman"/>'
        tblpr = ('<a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="4572000"/><a:gridCol w="4572000"/></a:tblGrid>')
        tcpr = '<a:tcPr/>'
        trh = 300000
        grp = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="组合 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
               '<p:grpSpPr><a:xfrm><a:off x="762000" y="5200000"/><a:ext cx="4572000" cy="914400"/>'
               '<a:chOff x="0" y="0"/><a:chExt cx="4572000" cy="1371600"/></a:xfrm></p:grpSpPr>'
               '<p:sp><p:nvSpPr><p:cNvPr id="10" name="组内文本框 10"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="1371600"/></a:xfrm></p:spPr>'
               '<p:txBody><a:bodyPr/>'
               '<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1800">%s</a:rPr>'
               '<a:t>组合内文字 text in group</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>' % font_box)

    title = (
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" altLang="en-US" sz="%d" b="1" dirty="0">%s</a:rPr>'
        '<a:t>Windows 迁移后文字错位排查</a:t></a:r></a:p></p:txBody></p:sp>'
        % ('<a:bodyPr/>' if not clean else '<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" anchor="t" wrap="square"/>',
           title_size, title_font)
    )
    body = (
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="内容占位符 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/>'
        '<a:p><a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>'
        '<a:buClr><a:srgbClr val="4472C4"/></a:buClr></a:pPr>%s</a:p>'
        '<a:p><a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>%s</a:p>'
        '<a:p>%s</a:p>'
        '<a:p>%s</a:p>'
        '</p:txBody></p:sp>'
        % (body_pr,
           _run("字体缺失会让行宽变化，从而整段换行位置改变。", 2000, '宋体' if not clean else None,
                '宋体' if not clean else None),
           _run("行距为百分比时，不同字体的度量差异会叠加成错位。", 2000, '宋体' if not clean else None,
                '宋体' if not clean else None),
           _run("段落三：包含 English words and 中文混排，用于测试中英文断行。", 2000, 'Calibri' if not clean else None,
                '等线' if not clean else None),
           _run("段落四：手动换行会在自动换行变化后出现断行偏移。", 2000, '宋体' if not clean else None,
                '宋体' if not clean else None))
    )
    box = (
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="不换行文本框 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="762000" y="4800000"/><a:ext cx="2514600" cy="457200"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>'
        % (box_pr, _run("不换行文本框：这一行文字很长，必须扩宽才不会压到别的元素。", 2400, 'Times New Roman' if not clean else None,
                        '宋体' if not clean else None))
    )
    box2 = (
        '<p:sp><p:nvSpPr><p:cNvPr id="5" name="自动调整文本框 5"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="8382000" y="4800000"/><a:ext cx="3048000" cy="914400"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>'
        % (box2_pr, _run("自动调整形状高度：WPS 与 PowerPoint 的重算结果不同。", 1800,
                         '微软雅黑' if not clean else None, '微软雅黑' if not clean else None))
    )
    table = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="表格 6"/>'
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="6096000" y="1400000"/><a:ext cx="9144000" cy="1200000"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        '<a:tbl>%s'
        '<a:tr h="%d"><a:tc>%s<a:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></a:txBody></a:tc>'
        '<a:tc>%s<a:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></a:txBody></a:tc></a:tr>'
        '<a:tr h="%d"><a:tc>%s<a:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></a:txBody></a:tc>'
        '<a:tc>%s<a:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></a:txBody></a:tc></a:tr>'
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
        % (tblpr, trh, tcpr, _run("表格第一格", 1400, '宋体' if not clean else None, None),
           tcpr, _run("第二格", 1400, '宋体' if not clean else None, None),
           trh, tcpr, _run("表格第三格：文字较多时行高会自动增长，不同实现增长量不同。", 1400,
                           '宋体' if not clean else None, None),
           tcpr, _run("第四格", 1400, '宋体' if not clean else None, None))
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sld %s>'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '%s%s%s%s%s%s'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>' % (DECL, title, body, box, box2, table, grp)
    )


def slide2(clean=False):
    if clean:
        bodypr = '<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" anchor="t" wrap="square" vert="horz"/>'
        font_xml = '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK SC"/><a:cs typeface="Liberation Sans"/>'
        text = "第二页：竖排与文字变形已规范化。"
        warp = ""
        lang = 'zh-CN'
    else:
        bodypr = '<a:bodyPr vert="eaVert"/>'
        font_xml = '<a:latin typeface="SimSun"/><a:ea typeface="SimSun"/><a:cs typeface="Times New Roman"/>'
        text = "第二页：竖排文字在字体替换后位置会偏移。"
        lang = 'en-US'
    sp = (
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="竖排文本框 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="762000" y="762000"/><a:ext cx="1828800" cy="4572000"/></a:xfrm></p:spPr>'
        '<p:txBody>%s<a:lstStyle/><a:p><a:r><a:rPr lang="%s" altLang="en-US" sz="2400">%s</a:rPr>'
        '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (bodypr, lang, font_xml, text)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sld %s>'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '%s'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sld>' % (DECL, sp)
    )


def build(clean=False):
    files = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": CORE,
        "docProps/app.xml": APP,
        "ppt/presentation.xml": presentation(clean),
        "ppt/_rels/presentation.xml.rels": PRESENTATION_RELS,
        "ppt/slideMasters/slideMaster1.xml": master(clean),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": layout(clean),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": LAYOUT_RELS,
        "ppt/slides/slide1.xml": slide1(clean),
        "ppt/slides/_rels/slide1.xml.rels": SLIDE_RELS,
        "ppt/slides/slide2.xml": slide2(clean),
        "ppt/slides/_rels/slide2.xml.rels": SLIDE_RELS,
        "ppt/theme/theme1.xml": THEME,
    }
    return files


def write(path, clean=False):
    files = build(clean)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            info = zipfile.ZipInfo(name)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.date_time = (2024, 1, 1, 0, 0, 0)
            z.writestr(info, data.encode("utf-8"))


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    os.makedirs(out, exist_ok=True)
    broken = os.path.join(out, "sample_broken.pptx")
    clean = os.path.join(out, "sample_clean.pptx")
    write(broken, clean=False)
    write(clean, clean=True)
    for p in (broken, clean):
        print("%s  (%d bytes)" % (p, os.path.getsize(p)))


if __name__ == "__main__":
    main()
