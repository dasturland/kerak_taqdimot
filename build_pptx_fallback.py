#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""python-pptx'siz (faqat standart kutubxona) PPTX generatori.

Bu skript `generate_presentation.py` bilan AYNAN bir xil kontentni ishlab
chiqaradi, ammo python-pptx o'rnatilmagan yoki internet cheklangan muhitlar
uchun mo'ljallangan. .pptx fayli OOXML (Office Open XML) formatidagi ZIP
to'plami bo'lgani uchun uni to'g'ridan-to'g'ri `zipfile` yordamida tuzamiz.

Mavzu: "Fleshkalar (USB) qanday qilib ma'lumot saqlaydi? Elektr zaryadlari
ko'rinishida axborot saqlash fizikasi".

Natija:
  * Aynan 10 ta slayd.
  * Har bir run'da shrift "Times New Roman" aniq o'rnatilgan.
  * Minimal shrift 12 pt (title 40, subtitle 24, slayd sarlavhasi 32, ...).
  * Har bir slaydga Unsplash rasmi joylanadi (yuklab olinmasa placeholder).
  * To'q, texnologik rang sxemasi (mikroelektronika / kvant fizikasi).

Foydalanish:
    python3 build_pptx_fallback.py [chiqish.pptx]
"""
import io
import os
import sys
import urllib.request
import zipfile
from xml.sax.saxutils import escape

from slides_content import (
    SLIDES, FONT_NAME,
    COLOR_BG_DARK, COLOR_PANEL, COLOR_ACCENT, COLOR_ACCENT2,
    COLOR_TITLE, COLOR_BODY,
    SIZE_TITLE, SIZE_SUBTITLE, SIZE_SLIDE_TITLE,
    SIZE_BODY_HEAD, SIZE_BODY, SIZE_META,
)
from placeholder import make_placeholder_png

# EMU: 1 inch = 914400 EMU. 16:9 slayd.
EMU = 914400
SLIDE_W = int(13.333 * EMU)
SLIDE_H = int(7.5 * EMU)


def hexc(t):
    return "%02X%02X%02X" % (t[0], t[1], t[2])


def inch(v):
    return int(v * EMU)


def fetch_image(url, idx):
    """Unsplash rasmini yuklab oladi; muvaffaqiyatsizlikda placeholder PNG."""
    headers = {"User-Agent": "Mozilla/5.0 (presentation-generator)"}
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=25) as resp:
                data = resp.read()
            if data and len(data) > 1000:
                print("  [OK] rasm yuklandi (%d bayt)" % len(data))
                return data, "jpeg"
        except Exception as exc:  # noqa: BLE001
            print("  [urinish %d] rasm yuklanmadi: %s" % (attempt + 1, exc))
    print("  [FALLBACK] placeholder rasm ishlatilmoqda")
    return make_placeholder_png(idx), "png"


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


def para_xml(runs_xml, align=None, space_before=None, space_after=None):
    ppr = ''
    inner = ''
    if space_before is not None:
        inner += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % int(space_before * 100)
    if space_after is not None:
        inner += '<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % int(space_after * 100)
    algn = ' algn="%s"' % align if align else ''
    if inner or algn:
        ppr = '<a:pPr%s>%s</a:pPr>' % (algn, inner)
    return '<a:p>%s%s</a:p>' % (ppr, runs_xml)


def rect_shape(sid, name, x, y, cx, cy, fill_color, line_color=None):
    line = ''
    if line_color:
        line = '<a:ln w="12700"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % hexc(line_color)
    else:
        line = '<a:ln><a:noFill/></a:ln>'
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (sid, name, x, y, cx, cy, hexc(fill_color), line)
    )


def textbox_shape(sid, name, x, y, cx, cy, paras_xml, anchor="t"):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="%s"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (sid, name, x, y, cx, cy, anchor, paras_xml)
    )


def pic_shape(sid, name, rid, x, y, cx, cy):
    return (
        '<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/><p:cNvPicPr>'
        '<a:picLocks noChangeAspect="0"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        % (sid, name, rid, x, y, cx, cy)
    )


# ----------------------------------------------------------------------------
# Slayd tuzuvchilari
# ----------------------------------------------------------------------------

def build_slide_shapes(data, idx, img_rid):
    """Slayd turiga qarab shakllar XML'ini qaytaradi."""
    shapes = []
    sid = [10]

    def nid():
        sid[0] += 1
        return sid[0]

    # orqa fon
    shapes.append(rect_shape(nid(), "bg", 0, 0, SLIDE_W, SLIDE_H, COLOR_BG_DARK))

    if data["kind"] == "title":
        # o'ngda rasm
        shapes.append(pic_shape(nid(), "img", img_rid, inch(7.4), 0, inch(5.933), SLIDE_H))
        # aksent chiziq
        shapes.append(rect_shape(nid(), "accent", inch(0.6), inch(1.2), inch(0.12), inch(2.4), COLOR_ACCENT))
        # sarlavha + kichik sarlavha
        p1 = para_xml(run_xml(data["title"], SIZE_TITLE, COLOR_TITLE, bold=True))
        p2 = para_xml(run_xml(data["subtitle"], SIZE_SUBTITLE, COLOR_ACCENT, italic=True), space_before=14)
        shapes.append(textbox_shape(nid(), "title", inch(0.9), inch(1.1), inch(6.1), inch(2.8), p1 + p2))
        # meta
        metas = ''.join(para_xml(run_xml(l, SIZE_META, COLOR_BODY), space_after=6) for l in data["lines"])
        shapes.append(textbox_shape(nid(), "meta", inch(0.9), inch(4.3), inch(6.1), inch(2.8), metas))

    elif data["kind"] == "closing":
        shapes.append(pic_shape(nid(), "img", img_rid, 0, 0, SLIDE_W, SLIDE_H))
        shapes.append(rect_shape(nid(), "panel", inch(1.5), inch(2.3), inch(10.33), inch(3.0), COLOR_BG_DARK, line_color=COLOR_ACCENT))
        p = para_xml(run_xml(data["title"], SIZE_TITLE, COLOR_ACCENT, bold=True), align="ctr")
        for l in data["lines"]:
            p += para_xml(run_xml(l, SIZE_META, COLOR_BODY), align="ctr", space_before=8)
        shapes.append(textbox_shape(nid(), "ctext", inch(1.8), inch(2.5), inch(9.7), inch(2.6), p, anchor="ctr"))

    else:  # content
        shapes.append(rect_shape(nid(), "topbar", 0, inch(1.15), SLIDE_W, inch(0.05), COLOR_ACCENT))
        shapes.append(rect_shape(nid(), "tbar", inch(0.6), inch(0.35), inch(0.12), inch(0.6), COLOR_ACCENT))
        pt = para_xml(run_xml(data["title"], SIZE_SLIDE_TITLE, COLOR_TITLE, bold=True))
        shapes.append(textbox_shape(nid(), "title", inch(0.9), inch(0.3), inch(11.8), inch(0.9), pt, anchor="ctr"))
        # rasm o'ngda
        shapes.append(pic_shape(nid(), "img", img_rid, inch(8.5), inch(1.6), inch(4.3), inch(5.4)))
        # matn paneli chapda
        shapes.append(rect_shape(nid(), "panel", inch(0.5), inch(1.6), inch(7.7), inch(5.4), COLOR_PANEL))
        body = ''
        for l in data["lines"]:
            is_bullet = l.strip().startswith("-")
            if is_bullet:
                body += para_xml(run_xml(l, SIZE_BODY, COLOR_BODY), space_after=10)
            else:
                body += para_xml(run_xml(l, SIZE_BODY_HEAD, COLOR_ACCENT, bold=True), space_after=10)
        shapes.append(textbox_shape(nid(), "body", inch(0.8), inch(1.85), inch(7.1), inch(4.9), body))

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


def slide_rels_xml(img_target):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="../media/%s"/>'
        '</Relationships>' % img_target
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
    '<dc:title>Fleshkalar (USB) qanday ishlaydi?</dc:title>'
    '<dc:creator>Toshbo\u02bbriyeva Marjona</dc:creator>'
    '<cp:lastModifiedBy>Toshbo\u02bbriyeva Marjona</cp:lastModifiedBy>'
    '</cp:coreProperties>'
)

APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>python-pptx-fallback</Application><Slides>10</Slides></Properties>'
)


def content_types_xml(media_exts):
    defaults = (
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
    )
    if "png" in media_exts:
        defaults += '<Default Extension="png" ContentType="image/png"/>'
    if "jpeg" in media_exts:
        defaults += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
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
    out = sys.argv[1] if len(sys.argv) > 1 else "fleshka_usb_taqdimot.pptx"
    n = len(SLIDES)

    # Rasmlarni oldindan olish
    media = []  # (filename, bytes)
    media_exts = set()
    slide_img_targets = []
    for idx, data in enumerate(SLIDES):
        print("Slayd %d/%d rasmi: %s" % (idx + 1, n, data["title"]))
        img_bytes, ext = fetch_image(data["image"], idx)
        fname = "image%d.%s" % (idx + 1, "png" if ext == "png" else "jpeg")
        media.append((fname, img_bytes))
        media_exts.add("png" if ext == "png" else "jpeg")
        slide_img_targets.append(fname)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml(media_exts))
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
        for fname, data_bytes in media:
            z.writestr("ppt/media/%s" % fname, data_bytes)
        for idx, data in enumerate(SLIDES):
            shapes = build_slide_shapes(data, idx, "rId2")
            z.writestr("ppt/slides/slide%d.xml" % (idx + 1), slide_xml(shapes))
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % (idx + 1),
                       slide_rels_xml(slide_img_targets[idx]))

    print("\nTayyor: %s" % os.path.abspath(out))
    print("Slaydlar soni: %d" % n)


if __name__ == "__main__":
    main()
