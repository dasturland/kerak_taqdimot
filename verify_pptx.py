#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Yaratilgan .pptx faylini standart kutubxona bilan tekshiradi.

Tekshiruvlar:
  * ZIP to'plami ochiladi va barcha XML qismlari to'g'ri (well-formed).
  * Aynan 10 ta slayd mavjud (presentation.xml va slide fayllari).
  * Har bir slaydda rasm (blip embed + media fayli) bor.
  * Har bir matn run'ida "Times New Roman" shrifti o'rnatilgan.
  * Hech bir shrift o'lchami 12 pt dan kichik emas.
Agar python-pptx mavjud bo'lsa, qo'shimcha ravishda u orqali ham ochib ko'radi.
"""
import re
import sys
import xml.dom.minidom as minidom
import zipfile

path = sys.argv[1] if len(sys.argv) > 1 else "fleshka_usb_taqdimot.pptx"
errors = []
ok = []

with zipfile.ZipFile(path) as z:
    names = z.namelist()

    # 1) Barcha XML qismlari well-formed
    for nm in names:
        if nm.endswith(".xml") or nm.endswith(".rels"):
            try:
                minidom.parseString(z.read(nm))
            except Exception as exc:  # noqa: BLE001
                errors.append("XML buzuq: %s -> %s" % (nm, exc))
    if not errors:
        ok.append("Barcha XML qismlari well-formed (%d ta)" % sum(
            1 for n in names if n.endswith((".xml", ".rels"))))

    # 2) Slaydlar soni
    slide_files = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    if len(slide_files) == 10:
        ok.append("Slaydlar soni = 10 (slide fayllari)")
    else:
        errors.append("Slaydlar soni %d, 10 bo'lishi kerak" % len(slide_files))

    pres = z.read("ppt/presentation.xml").decode("utf-8")
    sldids = re.findall(r"<p:sldId ", pres)
    if len(sldids) == 10:
        ok.append("presentation.xml ichida 10 ta sldId")
    else:
        errors.append("presentation.xml ichida %d ta sldId" % len(sldids))

    # 3) Har slaydda rasm
    media = [n for n in names if n.startswith("ppt/media/")]
    if len(media) == 10:
        ok.append("10 ta media (rasm) fayli joylangan")
    else:
        errors.append("Media fayllari soni %d" % len(media))
    for sf in slide_files:
        s = z.read(sf).decode("utf-8")
        if "r:embed" not in s:
            errors.append("%s da rasm (r:embed) yo'q" % sf)

    # 4) Shrift va o'lcham tekshiruvi
    bad_font = 0
    sizes = []
    runs_total = 0
    for sf in slide_files:
        s = z.read(sf).decode("utf-8")
        for m in re.finditer(r'<a:rPr[^>]*sz="(\d+)"[^>]*>(.*?)</a:rPr>', s, re.S):
            runs_total += 1
            sz = int(m.group(1))
            sizes.append(sz)
            block = m.group(0)
            if 'typeface="Times New Roman"' not in block:
                bad_font += 1
    if bad_font == 0:
        ok.append("Barcha %d ta matn run'ida 'Times New Roman' o'rnatilgan" % runs_total)
    else:
        errors.append("%d ta run'da Times New Roman yo'q" % bad_font)

    if sizes and min(sizes) >= 1200:
        ok.append("Minimal shrift o'lchami = %.0f pt (>= 12)" % (min(sizes) / 100.0))
    else:
        errors.append("Shrift o'lchami 12 pt dan kichik: %s" % (min(sizes) / 100.0 if sizes else "N/A"))

# Agar python-pptx bo'lsa, qo'shimcha tekshiruv
try:
    from pptx import Presentation
    prs = Presentation(path)
    cnt = len(prs.slides._sldIdLst)
    if cnt == 10:
        ok.append("python-pptx bilan ochildi, slaydlar soni = 10")
    else:
        errors.append("python-pptx: slaydlar soni %d" % cnt)
except ImportError:
    ok.append("(python-pptx o'rnatilmagan - faqat stdlib tekshiruvi bajarildi)")
except Exception as exc:  # noqa: BLE001
    errors.append("python-pptx ochishda xato: %s" % exc)

print("=== TEKSHIRUV NATIJALARI ===")
for o in ok:
    print("  [OK] " + o)
for e in errors:
    print("  [XATO] " + e)

if errors:
    print("\nNATIJA: MUVAFFAQIYATSIZ (%d xato)" % len(errors))
    sys.exit(1)
print("\nNATIJA: HAMMASI YAXSHI")
