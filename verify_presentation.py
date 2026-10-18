#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
taqdimot.pptx faylini tekshiruvchi (faqat standart kutubxona).

Tekshiradi:
  1) ppt/presentation.xml ichida aniq 10 ta <p:sldId> borligini.
  2) Har bir ppt/slides/slideN.xml da uchragan har bir `typeface`
     atributi "Times New Roman" ekanligini (boshqa shrift bo'lsa XATO).
  3) Har bir uchragan `sz` atributi butun son va >= 1200 (12pt) ekanligini.

PASS/FAIL xulosasini chiqaradi; xatolikda nolga teng bo'lmagan kod bilan
chiqadi.
"""

import re
import sys
import zipfile

PPTX = "taqdimot.pptx"
FONT_REQUIRED = "Times New Roman"
MIN_SZ = 1200

TYPEFACE_RE = re.compile(r'typeface="([^"]*)"')
SZ_RE = re.compile(r'\bsz="(\d+)"')
SLDID_RE = re.compile(r"<p:sldId\b")
# Matn yugurdaklari (<a:t>...</a:t>) ichidagi haqiqiy ko'rinadigan matn.
RUN_TEXT_RE = re.compile(r"<a:t>(.*?)</a:t>", re.DOTALL)

# Barcha slaydlar bo'ylab aynan uchrashi SHART bo'lgan tayanch satrlar.
# U+02BB (modifier letter turned comma) saqlanganini tekshirish uchun
# "Toshbo\u02bbriyeva" to'liq kiritilgan. Reja/Xulosa tuzilmasi ham shu yerda.
# Satrlar ataylab qisqa va barqaror tanlangan — qayta generatsiya buzilmasligi
# uchun (butun jumla emas, faqat o'zak so'z/boshlanish).
ANCHOR_STRINGS = [
    "Toshbo\u02bbriyeva",            # U+02BB apostrofli ism (1 va 10-slayd)
    "Kompyuterlar nima uchun qiziydi va qanday soviydi?",  # titul sarlavhasi
    "1-reja.",                        # 1-reja bo'lim belgisi (2-slayd)
    "2-reja.",                        # 2-reja bo'lim belgisi (2-slayd)
    "3-reja.",                        # 3-reja bo'lim belgisi (2-slayd)
    "Xulosa",                         # Xulosa (2 va 9-slayd)
]


def _xml_unescape(text):
    """OOXML <a:t> ichidagi standart XML entitylarni ochadi."""
    return (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def main():
    errors = []

    try:
        z = zipfile.ZipFile(PPTX)
    except Exception as exc:
        print("FAIL: %s faylini ochib bo'lmadi: %s" % (PPTX, exc))
        return 1

    names = z.namelist()

    # 1) Slayd soni presentation.xml bo'yicha
    if "ppt/presentation.xml" not in names:
        errors.append("ppt/presentation.xml topilmadi")
        pres = ""
    else:
        pres = z.read("ppt/presentation.xml").decode("utf-8")
    sldid_count = len(SLDID_RE.findall(pres))
    if sldid_count != 10:
        errors.append("presentation.xml da %d ta <p:sldId> topildi (10 kutilgan)"
                      % sldid_count)

    # Slayd xml fayllari soni
    slide_parts = sorted(
        n for n in names
        if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        and "/_rels/" not in n
    )
    if len(slide_parts) != 10:
        errors.append("ppt/slides/slideN.xml fayllari soni %d (10 kutilgan)"
                      % len(slide_parts))

    # 2) va 3) Har bir slaydda shrift va o'lcham
    for part in slide_parts:
        xml = z.read(part).decode("utf-8")

        for tf in TYPEFACE_RE.findall(xml):
            if tf != FONT_REQUIRED:
                errors.append("%s: ruxsat etilmagan shrift '%s' (faqat '%s')"
                              % (part, tf, FONT_REQUIRED))

        for sz in SZ_RE.findall(xml):
            try:
                val = int(sz)
            except ValueError:
                errors.append("%s: sz butun son emas: %r" % (part, sz))
                continue
            if val < MIN_SZ:
                errors.append("%s: sz=%d < %d (12pt dan kichik)"
                              % (part, val, MIN_SZ))

    # Barcha slaydlar typeface o'z ichiga olishini ham tekshiramiz
    for part in slide_parts:
        xml = z.read(part).decode("utf-8")
        if not TYPEFACE_RE.search(xml):
            errors.append("%s: hech qanday typeface topilmadi" % part)

    # 4) Deck haqiqiy matn yugurdaklariga ega bo'lishi SHART
    #    (faqat placeholderlardan iborat, matnsiz deck XATO bo'ladi).
    all_text_parts = []
    total_runs = 0
    for part in slide_parts:
        xml = z.read(part).decode("utf-8")
        for raw in RUN_TEXT_RE.findall(xml):
            text = _xml_unescape(raw)
            all_text_parts.append(text)
            if text.strip():
                total_runs += 1
    if total_runs == 0:
        errors.append("hech qanday matn yugurdagi (<a:t>) topilmadi — "
                      "deck faqat placeholderlardan iborat ko'rinadi")

    combined_text = "\n".join(all_text_parts)

    # 5) Tayanch satrlar (verbatim o'zbekcha matn + tuzilma) mavjudligini
    #    tekshiramiz — kontent regressiyasini (SLIDES jadvali o'zgarsa) ushlaydi.
    for anchor in ANCHOR_STRINGS:
        if anchor not in combined_text:
            errors.append("tayanch satr topilmadi: %r" % anchor)

    print("--- Tekshiruv natijalari ---")
    print("Slaydlar (presentation.xml <p:sldId>): %d" % sldid_count)
    print("Slayd XML fayllari: %d" % len(slide_parts))
    print("Talab qilingan shrift: %s" % FONT_REQUIRED)
    print("Eng kichik ruxsat etilgan o'lcham: %d (12pt)" % MIN_SZ)
    print("Matn yugurdaklari (<a:t>, bo'sh emas): %d" % total_runs)
    print("Tayanch satrlar tekshirildi: %d" % len(ANCHOR_STRINGS))

    if errors:
        print("\nFAIL — %d ta muammo topildi:" % len(errors))
        for e in errors:
            print("  - %s" % e)
        return 1

    print("\nPASS — 10 ta slayd, faqat 'Times New Roman', barcha o'lchamlar >= 12pt.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
