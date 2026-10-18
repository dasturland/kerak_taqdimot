#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""python-pptx'siz (faqat standart kutubxona) PROFESSIONAL PPTX generatori.

Bu skript `generate_presentation.py` bilan AYNAN bir xil kontentni ishlab
chiqaradi, ammo python-pptx o'rnatilmagan yoki internet cheklangan muhitlar
uchun mo'ljallangan. .pptx fayli OOXML (Office Open XML) formatidagi ZIP
to'plami bo'lgani uchun uni to'g'ridan-to'g'ri `zipfile` yordamida tuzamiz.

Yangilangan (professional) dizayn:
  * Aynan 10 ta slayd.
  * Har bir run'da shrift "Times New Roman" aniq o'rnatilgan.
  * Minimal shrift 12 pt (title 40, subtitle 24, slayd sarlavhasi 32, ...).
  * Har bir slaydda MAVZUGA ALOQADOR vektor illyustratsiya (DrawingML):
      1  stilizatsiya qilingan barmoq izi + skan chizig'i
      2  raqamlangan reja/mundarija ikonkasi
      3  ridge'lar + minutiae nuqtalari
      4  barmoq izi -> ikkilik/hash raqamlari
      5  LED yorug'lik nurlari va qaytish (optik)
      6  CMOS piksel panjarasi (2D tasvir)
      7  ultratovush konsentrik impulslari
      8  echo to'lqinlari + 3D relyef balandlik profili
      9  qalqon/qulf (yorug'lik + tovush birlashuvi)
     10  barmoq izi + qulf nishoni (yakun)
  * 4 qismli tuzilma uchun bo'lim indikatorlari (section pills).
  * To'q, texnologik kiberxavfsizlik rang sxemasi.

Rasmlar vektorli bo'lgani uchun hech qanday tashqi tarmoq yoki Pillow
kutubxonasi kerak emas. (Internetli muhitda `generate_presentation.py`
Unsplash fotosuratlarini yuklab oladi - URL'lar slides_content.py'da.)

Foydalanish:
    python3 build_pptx_fallback.py [chiqish.pptx]
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

from slides_content import (
    SLIDES, FONT_NAME,
    COLOR_BG_DARK, COLOR_PANEL, COLOR_ACCENT, COLOR_ACCENT2,
    COLOR_TITLE, COLOR_BODY,
    SIZE_TITLE, SIZE_SUBTITLE, SIZE_SLIDE_TITLE,
    SIZE_BODY_HEAD, SIZE_BODY, SIZE_META,
)
import illustrations as il

# EMU: 1 inch = 914400 EMU. 16:9 slayd.
EMU = 914400
SLIDE_W = int(13.333 * EMU)
SLIDE_H = int(7.5 * EMU)

PAL = il.make_palette(COLOR_PANEL, COLOR_ACCENT, COLOR_ACCENT2, COLOR_TITLE, COLOR_BG_DARK)

# 4 qismli tuzilma: har bir slayd qaysi bo'limga tegishli (indikator uchun).
#   0 = kirish (titul/reja), 1..3 = asosiy rejalar, 4 = xulosa/yakun
SECTIONS = ["KIRISH", "1-REJA", "1-REJA", "1-REJA",
            "2-REJA", "2-REJA", "3-REJA", "3-REJA", "XULOSA", "XULOSA"]
SECTION_STEP = [0, 1, 1, 1, 2, 2, 3, 3, 4, 4]  # 0..4 (5 ta nuqta indikator)


def hexc(t):
    return "%02X%02X%02X" % (t[0], t[1], t[2])


def inch(v):
    return int(v * EMU)


# ----------------------------------------------------------------------------
# XML yordamchi funksiyalari
# ----------------------------------------------------------------------------

def run_xml(text, size_pt, color, bold=False, italic=False):
    """Bitta matn run'i. Shrift (latin + cs + ea) har doim Times New Roman."""
    sz = int(size_pt * 100)
    b = ' b="1"' if bold else ''
    i = ' i="1"' if italic else ''
    return (
        '<a:r><a:rPr lang="uz-UZ" sz="%d"%s%s dirty="0">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/><a:cs typeface="%s"/><a:ea typeface="%s"/>'
        '</a:rPr><a:t>%s</a:t></a:r>'
        % (sz, b, i, hexc(color), FONT_NAME, FONT_NAME, FONT_NAME, escape(text))
    )


def para_xml(runs_xml, align=None, space_before=None, space_after=None, line_spc=None):
    inner = ''
    if line_spc is not None:
        inner += '<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % int(line_spc * 1000)
    if space_before is not None:
        inner += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(space_before * 100)
    if space_after is not None:
        inner += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % int(space_after * 100)
    algn = ' algn="%s"' % align if align else ''
    ppr = ''
    if inner or algn:
        ppr = '<a:pPr%s>%s</a:pPr>' % (algn, inner)
    return '<a:p>%s%s</a:p>' % (ppr, runs_xml)


def rect_shape(sid, name, x, y, cx, cy, fill_color, line_color=None, lw=12700, round_=False):
    prst = "roundRect" if round_ else "rect"
    if line_color:
        line = '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % (lw, hexc(line_color))
    else:
        line = '<a:ln><a:noFill/></a:ln>'
    fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % hexc(fill_color)
            if fill_color is not None else '<a:noFill/>')
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>%s%s</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (sid, name, x, y, cx, cy, prst, fill, line)
    )


def gradient_bg(sid, c1, c2):
    """Diagonal gradient orqa fon (to'q -> biroz ko'kroq) - professional ko'rinish."""
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="bg"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:gradFill><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="%s"/></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="%s"/></a:gs>'
        '</a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill>'
        '<a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (sid, SLIDE_W, SLIDE_H, hexc(c1), hexc(c2))
    )


def textbox_shape(sid, name, x, y, cx, cy, paras_xml, anchor="t"):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="%s" lIns="91440" rIns="91440">'
        '<a:normAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>'
        % (sid, name, x, y, cx, cy, anchor, paras_xml)
    )


# ----------------------------------------------------------------------------
# Umumiy dizayn elementlari
# ----------------------------------------------------------------------------

def section_indicator(nid, step):
    """O'ngdagi 5 nuqtali bo'lim indikatori (4 qismli tuzilmani ko'rsatadi)."""
    out = ''
    x = inch(12.75)
    for i in range(5):
        y = inch(2.9) + i * inch(0.42)
        active = (i == step)
        if active:
            out += rect_shape(nid(), "dotA", x - inch(0.03), y, inch(0.18), inch(0.18),
                              COLOR_ACCENT, round_=True)
        else:
            out += rect_shape(nid(), "dot", x, y, inch(0.12), inch(0.12),
                              None, line_color=(0x2A, 0x45, 0x63), lw=12700, round_=True)
    return out


def footer(nid, idx):
    """Pastki kolontitul: mavzu + slayd raqami."""
    out = rect_shape(nid(), "fline", inch(0.6), inch(7.02), inch(12.13), inch(0.016),
                     (0x24, 0x3B, 0x55))
    pL = para_xml(run_xml("Barmoq izi skanerlari \u2014 Optika va Ultratovush",
                          SIZE_META, (0x7A, 0x93, 0xAD)))
    out += textbox_shape(nid(), "fl", inch(0.6), inch(7.05), inch(8.0), inch(0.4), pL)
    pR = para_xml(run_xml("%02d / 10" % (idx + 1), SIZE_META, COLOR_ACCENT, bold=True),
                  align="r")
    out += textbox_shape(nid(), "fr", inch(9.5), inch(7.05), inch(3.23), inch(0.4), pR)
    return out


def section_pill(nid, idx):
    """Sarlavha tepasidagi kichik bo'lim yorlig'i (pill)."""
    label = SECTIONS[idx]
    pill = rect_shape(nid(), "pill", inch(0.9), inch(0.42), inch(1.75), inch(0.42),
                      COLOR_ACCENT2, round_=True)
    txt = para_xml(run_xml(label, SIZE_META, COLOR_TITLE, bold=True), align="ctr")
    pill += textbox_shape(nid(), "pilltxt", inch(0.9), inch(0.42), inch(1.75), inch(0.42),
                          txt, anchor="ctr")
    return pill


# ----------------------------------------------------------------------------
# Slayd tuzuvchilari
# ----------------------------------------------------------------------------

def build_slide_shapes(data, idx):
    """Slayd turiga qarab shakllar XML'ini qaytaradi (vektor rasm bilan)."""
    shapes = []
    sid = [10]

    def nid():
        sid[0] += 1
        return sid[0]

    # gradient orqa fon
    shapes.append(gradient_bg(nid(), COLOR_BG_DARK, COLOR_PANEL))

    if data["kind"] == "title":
        # chapda matn, o'ngda katta illyustratsiya kartasi
        shapes.append(rect_shape(nid(), "accent", inch(0.9), inch(1.5), inch(0.14), inch(2.6),
                                 COLOR_ACCENT))
        p1 = para_xml(run_xml(data["title"], SIZE_TITLE, COLOR_TITLE, bold=True), line_spc=104)
        p2 = para_xml(run_xml(data["subtitle"], SIZE_SUBTITLE, COLOR_ACCENT, italic=True),
                      space_before=16, line_spc=108)
        shapes.append(textbox_shape(nid(), "title", inch(1.25), inch(1.35), inch(6.0), inch(3.0),
                                    p1 + p2))
        metas = ''.join(para_xml(run_xml(l, SIZE_META, COLOR_BODY), space_after=7)
                        for l in data["lines"])
        shapes.append(rect_shape(nid(), "mpanel", inch(1.1), inch(4.5), inch(6.1), inch(2.3),
                                 COLOR_BG_DARK, line_color=(0x24, 0x3B, 0x55), lw=9525, round_=True))
        shapes.append(textbox_shape(nid(), "meta", inch(1.25), inch(4.6), inch(5.8), inch(2.1),
                                    metas))
        # illyustratsiya (o'ngda)
        shapes.append(il.build_illustration(idx, inch(7.9), inch(1.3), inch(4.6), inch(4.9), PAL))

    elif data["kind"] == "closing":
        # markazda illyustratsiya + matn paneli
        shapes.append(il.build_illustration(idx, inch(5.0), inch(0.75), inch(3.33), inch(3.3), PAL))
        shapes.append(rect_shape(nid(), "panel", inch(1.6), inch(4.2), inch(10.13), inch(2.5),
                                 COLOR_BG_DARK, line_color=COLOR_ACCENT, lw=15875, round_=True))
        p = para_xml(run_xml(data["title"], SIZE_TITLE, COLOR_ACCENT, bold=True), align="ctr")
        for l in data["lines"]:
            p += para_xml(run_xml(l, SIZE_META, COLOR_BODY), align="ctr", space_before=9)
        shapes.append(textbox_shape(nid(), "ctext", inch(1.9), inch(4.35), inch(9.5), inch(2.2),
                                    p, anchor="ctr"))

    else:  # content
        # bo'lim yorlig'i (pill)
        shapes.append(section_pill(nid, idx))
        # sarlavha
        pt = para_xml(run_xml(data["title"], SIZE_SLIDE_TITLE, COLOR_TITLE, bold=True),
                      line_spc=104)
        shapes.append(textbox_shape(nid(), "title", inch(0.85), inch(0.95), inch(11.5), inch(0.95),
                                    pt, anchor="ctr"))
        # sarlavha ostidagi aksent chizig'i
        shapes.append(rect_shape(nid(), "uline", inch(0.9), inch(1.95), inch(2.6), inch(0.05),
                                 COLOR_ACCENT))
        # matn paneli (chapda)
        shapes.append(rect_shape(nid(), "panel", inch(0.6), inch(2.2), inch(7.4), inch(4.6),
                                 COLOR_PANEL, line_color=(0x24, 0x3B, 0x55), lw=9525, round_=True))
        body = ''
        for l in data["lines"]:
            is_bullet = l.strip().startswith("-")
            if is_bullet:
                body += para_xml(run_xml(l, SIZE_BODY, COLOR_BODY), space_after=11, line_spc=110)
            else:
                body += para_xml(run_xml(l, SIZE_BODY_HEAD, COLOR_ACCENT, bold=True),
                                 space_after=11, line_spc=110)
        shapes.append(textbox_shape(nid(), "body", inch(0.9), inch(2.4), inch(6.9), inch(4.3),
                                    body))
        # illyustratsiya (o'ngda)
        shapes.append(il.build_illustration(idx, inch(8.35), inch(2.2), inch(4.1), inch(4.6), PAL))
        # bo'lim indikatori
        shapes.append(section_indicator(nid, SECTION_STEP[idx]))

    # kolontitul (barcha slaydlarda)
    shapes.append(footer(nid, idx))

    return ''.join(shapes)


def slide_xml(shapes_xml):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '%s</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping '
        'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" '
        'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>'
        % shapes_xml
    )


def slide_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )


# ----------------------------------------------------------------------------
# Statik to'plam qismlari (minimal, lekin to'liq yaroqli)
# ----------------------------------------------------------------------------

def presentation_xml(n):
    sldids = ''.join(
        '<p:sldId id="%d" r:id="rId%d"/>' % (256 + i, i + 1) for i in range(n)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId%d"/></p:sldMasterIdLst>'
        '<p:sldIdLst>%s</p:sldIdLst>'
        '<p:sldSz cx="%d" cy="%d" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
        % (n + 1, sldids, SLIDE_W, SLIDE_H)
    )


def presentation_rels_xml(n):
    rels = ''.join(
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        'Target="slides/slide%d.xml"/>' % (i + 1, i + 1) for i in range(n)
    )
    master = (
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>' % (n + 1)
    )
    theme = (
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/>' % (n + 2)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '%s%s%s</Relationships>' % (rels, master, theme)
    )


THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements><a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="0B1E33"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    '<a:accent1><a:srgbClr val="00C2D1"/></a:accent1><a:accent2><a:srgbClr val="3B82F6"/></a:accent2>'
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
    '</a:clrScheme><a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme><a:fmtScheme name="Office">'
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
    '</a:fmtScheme></a:themeElements></a:theme>'
)


SLIDE_MASTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0B1E33"/></a:solidFill>'
    '<a:effectLst/></p:bgPr></p:bg><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
    'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
    'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4000">'
    '<a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1600">'
    '<a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
    '<p:otherStyle/></p:txStyles></p:sldMaster>'
)

SLIDE_MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
    'Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
    'Target="../theme/theme1.xml"/></Relationships>'
)

SLIDE_LAYOUT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" '
    'bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
    'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" '
    'folHlink="folHlink"/></p:clrMapOvr></p:sldLayout>'
)

SLIDE_LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
    'Target="../slideMasters/slideMaster1.xml"/></Relationships>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
    'Target="docProps/app.xml"/></Relationships>'
)

CORE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>Barmoq izi skanerlari qanday ishlaydi?</dc:title>'
    '<dc:creator>Toshbo\u02bbriyeva Marjona</dc:creator>'
    '<cp:lastModifiedBy>Toshbo\u02bbriyeva Marjona</cp:lastModifiedBy>'
    '</cp:coreProperties>'
)

APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>python-pptx-fallback</Application><Slides>10</Slides></Properties>'
)


def content_types_xml():
    defaults = (
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
    )
    overrides = (
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    )
    for i in range(len(SLIDES)):
        overrides += (
            '<Override PartName="/ppt/slides/slide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % (i + 1)
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '%s%s</Types>' % (defaults, overrides)
    )


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "barmoq_izi_skanerlari.pptx"
    n = len(SLIDES)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml())
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("docProps/core.xml", CORE_XML)
        z.writestr("docProps/app.xml", APP_XML)
        z.writestr("ppt/presentation.xml", presentation_xml(n))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(n))
        z.writestr("ppt/theme/theme1.xml", THEME_XML)
        z.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_XML)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_XML)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        for idx, data in enumerate(SLIDES):
            print("Slayd %d/%d: %s" % (idx + 1, n, data["title"]))
            shapes = build_slide_shapes(data, idx)
            z.writestr("ppt/slides/slide%d.xml" % (idx + 1), slide_xml(shapes))
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % (idx + 1), slide_rels_xml())

    print("\nTayyor: %s" % os.path.abspath(out))
    print("Slaydlar soni: %d" % n)


if __name__ == "__main__":
    main()
