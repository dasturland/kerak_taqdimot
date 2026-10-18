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

    print("--- Tekshiruv natijalari ---")
    print("Slaydlar (presentation.xml <p:sldId>): %d" % sldid_count)
    print("Slayd XML fayllari: %d" % len(slide_parts))
    print("Talab qilingan shrift: %s" % FONT_REQUIRED)
    print("Eng kichik ruxsat etilgan o'lcham: %d (12pt)" % MIN_SZ)

    if errors:
        print("\nFAIL — %d ta muammo topildi:" % len(errors))
        for e in errors:
            print("  - %s" % e)
        return 1

    print("\nPASS — 10 ta slayd, faqat 'Times New Roman', barcha o'lchamlar >= 12pt.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
