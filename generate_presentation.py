#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kompyuterlar nima uchun qiziydi va qanday soviydi? — PPTX generator.

Bu skript 10 ta slayddan iborat professional o'zbekcha taqdimotni yaratadi.

ISHLASH TARTIBI (ikki yo'l):
  1) ASOSIY YO'L — python-pptx o'rnatilgan bo'lsa (foydalanuvchi mashinasida
     `pip install -r requirements.txt` dan keyin), taqdimot python-pptx orqali
     quriladi.
  2) ZAXIRA YO'L — python-pptx import qilinmasa (masalan, tarmoqsiz muhitda
     o'rnatib bo'lmasa), skript avtomatik ravishda faqat standart kutubxona
     (zipfile + qo'lda yozilgan OOXML/XML) yordamida ekvivalent .pptx faylini
     yaratadi.

RASMLAR: har bir slayd uchun Unsplash URL yuklab olinishga harakat qilinadi
(requests bo'lsa requests, aks holda urllib). Har qanday xatolikda (tarmoq
yopiq, 404, Pillow yo'q va h.k.) slayd mavzusiga mos bir rangli placeholder
to'rtburchak ishlatiladi — skript hech qachon qulamaydi.

Ikkala yo'l ham quyidagi qat'iy talablarni qondiradi:
  - Aniq 10 ta slayd
  - Shrift faqat "Times New Roman"
  - Eng kichik shrift o'lchami 12pt (sz=1200), sarlavhalar kattaroq
  - Qorong'u kiber/termodinamika dizayni
"""

import os
import sys
import zipfile
import urllib.request

# --------------------------------------------------------------------------- #
#  Umumiy sozlamalar
# --------------------------------------------------------------------------- #

OUTPUT_PPTX = "taqdimot.pptx"
IMAGES_DIR = "images"
DOWNLOAD_TIMEOUT = 6  # soniya

# 16:9 slayd o'lchami (EMU)
SLIDE_CX = 12192000
SLIDE_CY = 6858000

# Shrift
FONT_NAME = "Times New Roman"

# Rang palitrasi (hex, '#' siz)
BG_DARK = "0D1117"       # deyarli qora fon
BG_PANEL = "161B22"      # panel foni
TEXT_LIGHT = "F0F6FC"    # ochiq matn rangi
ACCENT_HEAT = "FF6B35"   # issiqlik — to'q sariq
ACCENT_HEAT2 = "E63946"  # issiqlik — qizil
ACCENT_COOL = "00B4D8"   # sovutish — ko'k
ACCENT_COOL2 = "48CAE4"  # sovutish — havorang
ACCENT_CYBER = "00B4D8"  # umumiy kiber urg'u — siyohrang-ko'k

# Shrift o'lchamlari (punkt)
SZ_TITLE = 40
SZ_SUBTITLE = 22
SZ_BODY = 18
SZ_META = 14  # 12 dan hech qachon past emas

# --------------------------------------------------------------------------- #
#  Slaydlar ma'lumoti (VERBATIM o'zbekcha matn)
# --------------------------------------------------------------------------- #

SLIDES = [
    {   # 1 — Titul
        "kind": "title",
        "accent": ACCENT_CYBER,
        "image": "https://images.unsplash.com/photo-1591799264318-7e6ef8ddb7ea?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Computer Processor",
        "img_label_uz": "Rasm: Protsessor",
        "title": "Kompyuterlar nima uchun qiziydi va qanday soviydi?",
        "subtitle": "Protsessor ishlayotganda issiqlik ajralib chiqishi va kulerlar yordamida sovutish fizikasi.",
        "body": [
            "Universitet: Jizzax davlat pedagogika universiteti tabiiy fanlar oliy maktabi",
            "Fan: Axborot tizimlari",
            "Guruh: 242_26",
            "Talaba: Toshbo\u02bbriyeva Marjona",
        ],
    },
    {   # 2 — Reja
        "kind": "content",
        "accent": ACCENT_CYBER,
        "image": "https://images.unsplash.com/photo-1587202372634-32705e3bf49c?auto=format&fit=crop&w=1000&q=80",
        "img_label": "PC Internals",
        "img_label_uz": "Rasm: Kompyuter ichki tuzilishi",
        "title": "Taqdimot rejasi",
        "body": [
            "1-reja. Tranzistorlar va issiqlik ajralib chiqish fizikasi (Joul-Lents qonuni).",
            "2-reja. Issiqlik o'tkazuvchanlik: Termopasta va radiatorlar vazifasi.",
            "3-reja. Konveksiya va faol sovutish (Kulerlar va suvli tizimlar).",
            "Xulosa. Issiqlik boshqaruvining zamonaviy axborot tizimlaridagi ahamiyati.",
        ],
    },
    {   # 3 — 1-reja (heat)
        "kind": "content",
        "accent": ACCENT_HEAT,
        "image": "https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Transistors Microchip",
        "img_label_uz": "Rasm: Tranzistorlar va mikrochip",
        "title": "1. Nega protsessor (CPU) qiziydi?",
        "body": [
            "Kompyuterning miyasi bo'lgan protsessor (CPU) ichida milliardlab mikroskopik tranzistorlar joylashgan.",
            "- Elektr signallari: Har doim kompyuter ma'lumotni (0 va 1) qayta ishlaganda, bu tranzistorlar orqali elektr toki o'tadi.",
            "- Energiya almashinuvi: Fizikaning energiya saqlanish qonuniga ko'ra, tranzistorlarning ishlashiga sarflangan elektr energiyasi yo'qolib ketmaydi, balki to'g'ridan-to'g'ri issiqlik energiyasiga aylanadi.",
        ],
    },
    {   # 4 — 1-reja (heat)
        "kind": "content",
        "accent": ACCENT_HEAT2,
        "image": "https://images.unsplash.com/photo-1614729939124-032f0b56c9ce?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Physics Energy",
        "img_label_uz": "Rasm: Energiya va fizika",
        "title": "Joul-Lents qonuni amalda",
        "body": [
            "Protsessordagi qizish bevosita fizikaning Joul-Lents qonuni bilan tushuntiriladi.",
            "- Qarshilik va Issiqlik: Har qanday o'tkazgich elektr tokiga ma'lum bir qarshilik ko'rsatadi. Elektr toki bu qarshilikni yengib o'tayotganda atomlar bilan to'qnashadi va kinetik energiya (issiqlik) hosil qiladi.",
            "- TDP (Thermal Design Power): Bu protsessorning maksimal yuklamada qancha issiqlik (Vatt hisobida) ajratib chiqarishini ko'rsatuvchi ko'rsatkichdir.",
        ],
    },
    {   # 5 — 2-reja (cool)
        "kind": "content",
        "accent": ACCENT_COOL,
        "image": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Thermal Paste",
        "img_label_uz": "Rasm: Termopasta",
        "title": "2. Issiqlikni uzoqlashtirish: Termopasta",
        "body": [
            "Protsessor yuzasida hosil bo'lgan kuchli issiqlikni tezda olib tashlash axborot tizimining barqaror ishlashi uchun juda muhim.",
            "- Mikroskopik bo'shliqlar: Protsessor va uning ustiga qo'yiladigan radiator yuzasi ko'zga tekis ko'rinsa-da, mikroskop ostida g'adir-budur bo'ladi. Ular orasida issiqlikni yomon o'tkazuvchi havo qolib ketadi.",
            "- Termopastaning vazifasi: Suyuq termopasta ana shu bo'shliqlarni to'ldiradi va issiqlikni protsessordan radiatorga maksimal darajada tez o'tkazib beradi.",
        ],
    },
    {   # 6 — 2-reja (cool)
        "kind": "content",
        "accent": ACCENT_COOL2,
        "image": "https://images.unsplash.com/photo-1555680202-c86f0e12f086?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Heatsink",
        "img_label_uz": "Rasm: Radiator (Heatsink)",
        "title": "Radiator (Heatsink) fizikasi",
        "body": [
            "Termopasta orqali o'tgan issiqlik mis yoki alyuminiydan yasalgan radiatorga yetib boradi.",
            "- Yuzani kattalashtirish fizikasi: Radiator nima uchun ko'plab yupqa qanotchalardan (plastinalardan) iborat? Termodinamika qonunlariga ko'ra, jismning yuzasi qanchalik katta bo'lsa, u atrof-muhitga issiqlikni shunchalik tez tarqatadi.",
            "- Mis va Alyuminiy: Mis issiqlikni tez qabul qiladi, alyuminiy esa uni havoga tez tarqatadi. Ko'p radiatorlar shu ikki metallning gibrididan yasaladi.",
        ],
    },
    {   # 7 — 3-reja (cool)
        "kind": "content",
        "accent": ACCENT_COOL,
        "image": "https://images.unsplash.com/photo-1587202372775-b22f4625b15a?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Computer Fan",
        "img_label_uz": "Rasm: Kuler (ventilyator)",
        "title": "3. Konveksiya: Kuler (Ventilyator) qanday soviydi?",
        "body": [
            "Radiator o'ziga issiqlikni olgach, uning atrofidagi havo qiziydi. Agar bu havo yangilanmasa, sovutish to'xtaydi.",
            "- Faol sovutish (Active Cooling): Kuler (ventilyator) radiator qanotchalari orasiga sovuq havoni kuch bilan purkaydi.",
            "- Majburiy konveksiya: Fizikadagi konveksiya hodisasi tufayli sovuq havo molekulalari radiatordagi issiqlikni o'ziga olib, korpusdan tashqariga uchib chiqib ketadi.",
        ],
    },
    {   # 8 — 3-reja (cool)
        "kind": "content",
        "accent": ACCENT_COOL2,
        "image": "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Liquid Cooling",
        "img_label_uz": "Rasm: Suvli sovutish",
        "title": "Suvli sovutish (Liquid Cooling) texnologiyasi",
        "body": [
            "Juda kuchli protsessorlar (masalan, Data markazlar yoki o'yin kompyuterlari) uchun oddiy havo yetarli bo'lmaydi.",
            "- Suvning issiqlik sig'imi: Fizikaga ko'ra, suvning issiqlik sig'imi havoga qaraganda ancha yuqori. Ya'ni suv o'ziga ko'proq issiqlikni yuta oladi.",
            "- Suyuqlik aylanmasi: Maxsus nasos yordamida suyuqlik protsessor ustidagi blokdan o'tib, issiqlikni yutadi va katta radiatorga olib borib, ventilyatorlar yordamida soviydi.",
        ],
    },
    {   # 9 — Xulosa
        "kind": "content",
        "accent": ACCENT_CYBER,
        "image": "https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Data Center",
        "img_label_uz": "Rasm: Ma'lumotlar markazi",
        "title": "Xulosa",
        "body": [
            "Xulosa qilib aytganda, kompyuterning ishlashi to'liq elektr toki va termodinamika qonunlariga bog'liq. Axborot qanchalik tez qayta ishlansa, shuncha ko'p elektr issiqlikka aylanadi.",
            "Samarali sovutish tizimlari (kulerlar, radiatorlar va termopasta) nafaqat kompyuterning qotib qolmasligini, balki uning ehtiyot qismlari yonib ketmasdan uzoq yillar barqaror ishlashini kafolatlovchi asosiy injenerlik yechimidir.",
        ],
    },
    {   # 10 — E'tibor uchun
        "kind": "thanks",
        "accent": ACCENT_CYBER,
        "image": "https://images.unsplash.com/photo-1517245386807-bb43f82c33c4?auto=format&fit=crop&w=1000&q=80",
        "img_label": "Abstract IT",
        "img_label_uz": "Rasm: Axborot texnologiyalari",
        "title": "E'tiboringiz uchun rahmat!",
        "body": [
            "Mavzu: Kompyuterning qizishi va sovutish fizikasi",
            "Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - guruh)",
            "Jizzax davlat pedagogika universiteti - Axborot tizimlari fani.",
        ],
    },
]

assert len(SLIDES) == 10, "Taqdimot aniq 10 ta slayddan iborat bo'lishi shart"


# --------------------------------------------------------------------------- #
#  Rasm yuklash (umumiy, ikkala yo'l uchun)
# --------------------------------------------------------------------------- #

# Qo'llab-quvvatlanadigan rasm turlari: (kengaytma, content-type).
# Noma'lum tur uchun jpg/jpeg ga qaytamiz.
_EXT_CONTENT_TYPE = {
    "jpg": "image/jpeg",
    "png": "image/png",
    "gif": "image/gif",
    "webp": "image/webp",
    "bmp": "image/bmp",
}


def sniff_image_ext(data, content_type=None, url=None):
    """Yuklab olingan baytlar asosida rasm kengaytmasini aniqlaydi.

    Avval sehrli baytlar (magic bytes) tekshiriladi, so'ng HTTP content-type,
    oxirida URL kengaytmasi. Hech biri mos kelmasa 'jpg' qaytariladi.
    """
    if data:
        if data[:3] == b"\xff\xd8\xff":
            return "jpg"
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return "png"
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return "gif"
        if len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP":
            return "webp"
        if data[:2] == b"BM":
            return "bmp"
    if content_type:
        ct = content_type.split(";")[0].strip().lower()
        for ext, known in _EXT_CONTENT_TYPE.items():
            if ct == known:
                return ext
        if ct == "image/jpg":
            return "jpg"
    if url:
        low = url.split("?")[0].lower()
        for ext in _EXT_CONTENT_TYPE:
            if low.endswith("." + ext) or low.endswith(".jpeg"):
                return "jpg" if ext == "jpg" else ext
    return "jpg"


def content_type_for_ext(ext):
    """Kengaytma uchun content-type qaytaradi (noma'lum bo'lsa image/jpeg)."""
    return _EXT_CONTENT_TYPE.get(ext, "image/jpeg")


def try_download_image(url, dest_dir, basename):
    """URL ni yuklab olishga harakat qiladi.

    Avval requests, bo'lmasa urllib ishlatiladi. Haqiqiy fayl turini (JPEG,
    PNG, WebP va h.k.) aniqlab, mos kengaytma bilan saqlaydi. Muvaffaqiyat
    bo'lsa to'liq yo'lni qaytaradi, aks holda None (skript qulamaydi).
    """
    try:
        data = None
        content_type = None
        try:
            import requests  # type: ignore
            resp = requests.get(url, timeout=DOWNLOAD_TIMEOUT)
            if resp.status_code == 200 and resp.content:
                data = resp.content
                content_type = resp.headers.get("Content-Type")
        except ImportError:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as r:
                data = r.read()
                content_type = r.headers.get("Content-Type")
        if not data:
            return None
        ext = sniff_image_ext(data, content_type, url)
        dest_path = os.path.join(dest_dir, "%s.%s" % (basename, ext))
        with open(dest_path, "wb") as f:
            f.write(data)
        return dest_path
    except Exception:
        return None


def prepare_images():
    """Har bir slayd uchun rasmni yuklashga urinadi.

    Natijalar ro'yxatini qaytaradi: har bir element (ok, path_or_None, reason).
    """
    if not os.path.isdir(IMAGES_DIR):
        os.makedirs(IMAGES_DIR, exist_ok=True)

    results = []
    for idx, slide in enumerate(SLIDES, start=1):
        path = try_download_image(slide["image"], IMAGES_DIR, "slide%02d" % idx)
        if path and os.path.exists(path) and os.path.getsize(path) > 0:
            ext = os.path.splitext(path)[1].lstrip(".").lower() or "jpg"
            results.append((True, path, ext))
            print("Slide %d: image OK (%s, %s)" % (idx, path, content_type_for_ext(ext)))
        else:
            results.append((False, None, None))
            print("Slide %d: placeholder (download failed / network blocked)" % idx)
    return results


# --------------------------------------------------------------------------- #
#  YO'L 1 — python-pptx
# --------------------------------------------------------------------------- #

def build_with_pptx(image_results):
    # ESLATMA (smoke-test): bu python-pptx yo'li tarmoqsiz sandboxda ishga
    # tushirib bo'lmaydi (python-pptx/lxml o'rnatilmaydi). Shu sababli uni
    # tarmoqli mashinada bir marta "smoke run" qilib tekshirish tavsiya
    # etiladi: `pip install -r requirements.txt && python3 generate_presentation.py`.
    # Ayniqsa quyidagi set_run ichidagi qo'lda a:latin/a:cs/a:ea qo'shilishi
    # python-pptx ning o'z font.name ishlovi bilan birga faqat BITTA toza
    # a:latin hosil qilishini (dublikatsiz) va rasm yuklash haqiqiy tur bilan
    # (JPEG/PNG/WebP) joylashishini shu yerda tasdiqlash kerak.
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE

    def rgb(hexs):
        return RGBColor.from_string(hexs)

    prs = Presentation()
    prs.slide_width = Emu(SLIDE_CX)
    prs.slide_height = Emu(SLIDE_CY)
    blank = prs.slide_layouts[6]

    def set_run(run, text, size, color, bold=False):
        run.text = text
        run.font.name = FONT_NAME
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
        # Murakkab yozuv / sharqiy-osiyo shriftini ham majburan Times New Roman
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:latin", "a:cs", "a:ea"):
            el = rPr.makeelement(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1],
                {"typeface": FONT_NAME},
            )
            # eskisini olib tashlab, yangisini qo'shamiz
            for old in rPr.findall(el.tag):
                rPr.remove(old)
            rPr.append(el)

    def fill_bg(slide, color):
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = rgb(color)

    def add_rect(slide, x, y, cx, cy, color):
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(cx), Emu(cy))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        return shp

    EMU_IN = 914400
    for idx, slide_data in enumerate(SLIDES, start=1):
        slide = prs.slides.add_slide(blank)
        fill_bg(slide, BG_DARK)
        accent = slide_data["accent"]

        # Chap urg'u chizig'i
        add_rect(slide, int(0.0 * EMU_IN), 0, int(0.18 * EMU_IN), SLIDE_CY, accent)

        # Rasm yoki placeholder (o'ng tomon)
        img_x = int(8.6 * EMU_IN)
        img_y = int(1.0 * EMU_IN)
        img_cx = int(4.2 * EMU_IN)
        img_cy = int(5.3 * EMU_IN)
        ok, path, _ext = image_results[idx - 1]
        if ok and path:
            try:
                slide.shapes.add_picture(path, Emu(img_x), Emu(img_y), Emu(img_cx), Emu(img_cy))
            except Exception:
                ok = False
        if not ok:
            ph = add_rect(slide, img_x, img_y, img_cx, img_cy, BG_PANEL)
            ph.line.color.rgb = rgb(accent)
            ph.line.width = Pt(2)
            tf = ph.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            # O'zbekcha placeholder yorlig'i (tarmoqsiz qurishda butun slayd
            # o'zbekcha qolishi uchun inglizcha img_label o'rniga ishlatiladi).
            set_run(p.add_run(), slide_data["img_label_uz"], SZ_META, accent, bold=True)

        # Sarlavha
        title_box = slide.shapes.add_textbox(
            Emu(int(0.5 * EMU_IN)), Emu(int(0.5 * EMU_IN)),
            Emu(int(7.8 * EMU_IN)), Emu(int(1.6 * EMU_IN)),
        )
        ttf = title_box.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        set_run(tp.add_run(), slide_data["title"], SZ_TITLE, accent, bold=True)

        # Kichik sarlavha (faqat titul slaydda)
        body_top = 2.3
        if slide_data.get("subtitle"):
            sub_box = slide.shapes.add_textbox(
                Emu(int(0.5 * EMU_IN)), Emu(int(2.1 * EMU_IN)),
                Emu(int(7.8 * EMU_IN)), Emu(int(1.1 * EMU_IN)),
            )
            stf = sub_box.text_frame
            stf.word_wrap = True
            sp = stf.paragraphs[0]
            set_run(sp.add_run(), slide_data["subtitle"], SZ_SUBTITLE, TEXT_LIGHT, bold=False)
            body_top = 3.3

        # Asosiy matn
        is_meta = slide_data["kind"] in ("title", "thanks")
        body_size = SZ_META if is_meta else SZ_BODY
        body_box = slide.shapes.add_textbox(
            Emu(int(0.5 * EMU_IN)), Emu(int(body_top * EMU_IN)),
            Emu(int(7.8 * EMU_IN)), Emu(int((7.0 - body_top) * EMU_IN)),
        )
        btf = body_box.text_frame
        btf.word_wrap = True
        for i, line in enumerate(slide_data["body"]):
            p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            p.space_after = Pt(8)
            set_run(p.add_run(), line, body_size, TEXT_LIGHT, bold=False)

    prs.save(OUTPUT_PPTX)


# --------------------------------------------------------------------------- #
#  YO'L 2 — faqat standart kutubxona (OOXML qo'lda)
# --------------------------------------------------------------------------- #

def xml_escape(text):
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _run_xml(text, size_pt, color_hex, bold):
    """Bitta <a:r> yaratadi — har doim Times New Roman, sz hundredths-of-point."""
    sz = int(size_pt * 100)
    b = ' b="1"' if bold else ""
    return (
        '<a:r><a:rPr lang="uz-UZ" sz="%d"%s dirty="0">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/><a:cs typeface="%s"/><a:ea typeface="%s"/>'
        '</a:rPr><a:t>%s</a:t></a:r>'
        % (sz, b, color_hex, FONT_NAME, FONT_NAME, FONT_NAME, xml_escape(text))
    )


def _para_xml(runs_xml, align="l", space_after=400):
    return (
        '<a:p><a:pPr algn="%s"><a:spcAft><a:spcPts val="%d"/></a:spcAft></a:pPr>%s</a:p>'
        % (align, space_after, runs_xml)
    )


def _textbox_sp(spid, name, x, y, cx, cy, paragraphs_xml, anchor="t"):
    return (
        '<p:sp><p:nvSpPr>'
        '<p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="%s"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (spid, name, x, y, cx, cy, anchor, paragraphs_xml)
    )


def _rect_sp(spid, name, x, y, cx, cy, fill_hex, line_hex=None, paragraphs_xml="", anchor="ctr"):
    line = ""
    if line_hex:
        line = ('<a:ln w="25400"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
                % line_hex)
    body = ('<p:txBody><a:bodyPr wrap="square" anchor="%s"><a:normAutofit/></a:bodyPr>'
            '<a:lstStyle/>%s</p:txBody>' % (anchor, paragraphs_xml or "<a:p/>"))
    return (
        '<p:sp><p:nvSpPr>'
        '<p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</p:spPr>%s</p:sp>'
        % (spid, name, x, y, cx, cy, fill_hex, line, body)
    )


def _pic_sp(spid, name, rembed, x, y, cx, cy):
    return (
        '<p:pic><p:nvPicPr>'
        '<p:cNvPr id="%d" name="%s"/><p:cNvPicPr/><p:nvPr/>'
        '</p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        % (spid, name, rembed, x, y, cx, cy)
    )


def _build_slide_xml(slide_data, image_result):
    EMU_IN = 914400
    accent = slide_data["accent"]
    shapes = []
    spid = 2

    # Fon (butun slaydni qoplaydigan to'rtburchak)
    shapes.append(_rect_sp(spid, "bg", 0, 0, SLIDE_CX, SLIDE_CY, BG_DARK))
    spid += 1
    # Chap urg'u chizig'i
    shapes.append(_rect_sp(spid, "accentbar", 0, 0, int(0.18 * EMU_IN), SLIDE_CY, accent))
    spid += 1

    # Rasm / placeholder
    img_x = int(8.6 * EMU_IN)
    img_y = int(1.0 * EMU_IN)
    img_cx = int(4.2 * EMU_IN)
    img_cy = int(5.3 * EMU_IN)
    ok, _path, _ext = image_result
    pic_rel = None
    if ok:
        pic_rel = "rIdImg"
        shapes.append(_pic_sp(spid, "picture", pic_rel, img_x, img_y, img_cx, img_cy))
        spid += 1
    else:
        # O'zbekcha placeholder yorlig'i (tarmoqsiz qurishda butun slayd o'zbekcha qoladi).
        label_p = _para_xml(_run_xml(slide_data["img_label_uz"], SZ_META, accent, True), align="ctr")
        shapes.append(_rect_sp(spid, "placeholder", img_x, img_y, img_cx, img_cy,
                               BG_PANEL, line_hex=accent, paragraphs_xml=label_p, anchor="ctr"))
        spid += 1

    # Sarlavha
    title_p = _para_xml(_run_xml(slide_data["title"], SZ_TITLE, accent, True))
    shapes.append(_textbox_sp(spid, "title", int(0.5 * EMU_IN), int(0.5 * EMU_IN),
                              int(7.8 * EMU_IN), int(1.6 * EMU_IN), title_p))
    spid += 1

    body_top = 2.3
    if slide_data.get("subtitle"):
        sub_p = _para_xml(_run_xml(slide_data["subtitle"], SZ_SUBTITLE, TEXT_LIGHT, False))
        shapes.append(_textbox_sp(spid, "subtitle", int(0.5 * EMU_IN), int(2.1 * EMU_IN),
                                  int(7.8 * EMU_IN), int(1.1 * EMU_IN), sub_p))
        spid += 1
        body_top = 3.3

    # Asosiy matn
    is_meta = slide_data["kind"] in ("title", "thanks")
    body_size = SZ_META if is_meta else SZ_BODY
    body_paras = "".join(
        _para_xml(_run_xml(line, body_size, TEXT_LIGHT, False)) for line in slide_data["body"]
    )
    shapes.append(_textbox_sp(spid, "body", int(0.5 * EMU_IN), int(body_top * EMU_IN),
                              int(7.8 * EMU_IN), int((7.0 - body_top) * EMU_IN), body_paras))
    spid += 1

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
        % "".join(shapes), pic_rel
    )


def _slide_rels_xml(has_pic, img_filename):
    rels = ('<Relationship Id="rIdLayout" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
            'Target="../slideLayouts/slideLayout1.xml"/>')
    if has_pic:
        rels += ('<Relationship Id="rIdImg" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                 'Target="../media/%s"/>' % img_filename)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '%s</Relationships>' % rels)


def build_with_stdlib(image_results):
    n = len(SLIDES)

    # --- statik qismlar ---
    content_types_overrides = []
    for i in range(1, n + 1):
        content_types_overrides.append(
            '<Override PartName="/ppt/slides/slide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % i
        )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="gif" ContentType="image/gif"/>'
        '<Default Extension="webp" ContentType="image/webp"/>'
        '<Default Extension="bmp" ContentType="image/bmp"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '%s</Types>' % "".join(content_types_overrides)
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/></Relationships>'
    )

    # presentation.xml — 10 ta slayd id
    sld_ids = "".join(
        '<p:sldId id="%d" r:id="rIdSlide%d"/>' % (255 + i, i) for i in range(1, n + 1)
    )
    presentation = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>'
        '<p:sldIdLst>%s</p:sldIdLst>'
        '<p:sldSz cx="%d" cy="%d" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
        % (sld_ids, SLIDE_CX, SLIDE_CY)
    )

    pres_rels_items = [
        '<Relationship Id="rIdMaster" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rIdTheme" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/>',
    ]
    for i in range(1, n + 1):
        pres_rels_items.append(
            '<Relationship Id="rIdSlide%d" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            'Target="slides/slide%d.xml"/>' % (i, i)
        )
    presentation_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '%s</Relationships>' % "".join(pres_rels_items)
    )

    theme = _theme_xml()

    slide_master = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
        'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout"/></p:sldLayoutIdLst>'
        '</p:sldMaster>' % BG_DARK
    )
    slide_master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdLayout" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rIdTheme" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="../theme/theme1.xml"/></Relationships>'
    )

    slide_layout = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'type="blank" preserve="1">'
        '<p:cSld name="Blank"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
        'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '</p:clrMapOvr></p:sldLayout>'
    )
    slide_layout_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdMaster" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="../slideMasters/slideMaster1.xml"/></Relationships>'
    )

    # --- zip yozamiz ---
    with zipfile.ZipFile(OUTPUT_PPTX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("ppt/presentation.xml", presentation)
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels)
        z.writestr("ppt/theme/theme1.xml", theme)
        z.writestr("ppt/slideMasters/slideMaster1.xml", slide_master)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels)

        for idx, slide_data in enumerate(SLIDES, start=1):
            result = image_results[idx - 1]
            slide_xml, pic_rel = _build_slide_xml(slide_data, result)
            z.writestr("ppt/slides/slide%d.xml" % idx, slide_xml)
            has_pic = pic_rel is not None
            ok, path, ext = result
            ext = (ext or "jpg").lower()
            img_filename = "image%d.%s" % (idx, ext)
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % idx,
                       _slide_rels_xml(has_pic, img_filename))
            if has_pic:
                with open(path, "rb") as f:
                    z.writestr("ppt/media/%s" % img_filename, f.read())


def _theme_xml():
    def clr(name, val):
        return '<a:%s><a:srgbClr val="%s"/></a:%s>' % (name, val, name)
    scheme = (
        '<a:clrScheme name="Cyber">'
        + clr("dk1", "F0F6FC") + clr("lt1", BG_DARK)
        + clr("dk2", "F0F6FC") + clr("lt2", BG_PANEL)
        + clr("accent1", ACCENT_HEAT) + clr("accent2", ACCENT_HEAT2)
        + clr("accent3", ACCENT_COOL) + clr("accent4", ACCENT_COOL2)
        + clr("accent5", ACCENT_CYBER) + clr("accent6", "7B2CBF")
        + clr("hlink", ACCENT_COOL) + clr("folHlink", ACCENT_COOL2)
        + '</a:clrScheme>'
    )
    font_scheme = (
        '<a:fontScheme name="Times New Roman">'
        '<a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface="Times New Roman"/>'
        '<a:cs typeface="Times New Roman"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface="Times New Roman"/>'
        '<a:cs typeface="Times New Roman"/></a:minorFont></a:fontScheme>'
    )
    fmt_scheme = (
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst></a:fmtScheme>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Cyber">'
        '<a:themeElements>%s%s%s</a:themeElements>'
        '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
        % (scheme, font_scheme, fmt_scheme)
    )


# --------------------------------------------------------------------------- #
#  Asosiy
# --------------------------------------------------------------------------- #

def main():
    print("=== Taqdimot generatori ===")
    print("Rasmlarni tayyorlash...")
    image_results = prepare_images()

    try:
        import pptx  # noqa: F401
        path_used = "python-pptx"
    except Exception:
        path_used = "stdlib (OOXML fallback)"

    print("\nTanlangan generatsiya yo'li: %s" % path_used)
    if path_used == "python-pptx":
        try:
            build_with_pptx(image_results)
        except Exception as exc:
            print("python-pptx yo'lida xatolik (%s); stdlib yo'liga o'tilmoqda." % exc)
            path_used = "stdlib (OOXML fallback)"
            build_with_stdlib(image_results)
    else:
        build_with_stdlib(image_results)

    size = os.path.getsize(OUTPUT_PPTX)
    print("\nFayl yaratildi: %s (%d bayt)" % (OUTPUT_PPTX, size))
    print("Ishlatilgan yo'l: %s" % path_used)
    n_img = sum(1 for ok, _, _ in image_results if ok)
    print("Rasmlar: %d ta yuklandi, %d ta placeholder." % (n_img, len(SLIDES) - n_img))
    return 0


if __name__ == "__main__":
    sys.exit(main())
