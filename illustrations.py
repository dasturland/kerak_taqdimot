# -*- coding: utf-8 -*-
"""Mavzuga aloqador vektor rasmlar (DrawingML) generatori.

Bu modul har bir slayd uchun MAVZUGA MOS vektorli illyustratsiya yaratadi.
Rasmlar PowerPoint'ning o'z DrawingML shakllari (freeform egri chiziqlar,
aylanalar, chiziqlar, matn) sifatida to'g'ridan-to'g'ri slayd XML ichiga
joylanadi. Shu sabab:

  * Hech qanday tashqi tarmoq (Unsplash yoki boshqa) kerak emas.
  * Pillow/PIL yoki boshqa kutubxona kerak emas - faqat standart kutubxona.
  * Rasmlar vektorli bo'lgani uchun istalgan o'lchamda tiniq ko'rinadi.

Har bir funksiya `(x, y, w, h, color_fns)` qabul qiladi (EMU birligida) va
`<p:grpSp>` (shakllar guruhi) XML satrini qaytaradi. Guruh ichidagi
koordinatalar 0..100000 lokal tizimda chiziladi va `chOff/chExt` orqali
berilgan to'rtburchakka moslanadi.

Palitra (slides_content bilan muvofiq, to'q-texnologik biometrika uslubi):
  * fon         : to'q ko'k panel
  * accent      : moviy / tech (yorug'lik + tovush)
  * accent2     : ko'k
  * ink         : oq / ochiq matn
"""
import math

# Lokal guruh koordinata tizimi (barcha motivlar shu kvadratda chiziladi)
U = 100000  # 100000 x 100000 lokal birlik


def _hex(t):
    return "%02X%02X%02X" % (t[0], t[1], t[2])


# ---------------------------------------------------------------------------
# Past darajali shakl yordamchilari (guruh ichida ishlatiladi)
# ---------------------------------------------------------------------------

def _sp_id():
    _sp_id.n += 1
    return _sp_id.n


_sp_id.n = 1000


def rect(x, y, w, h, fill, line=None, lw=9525, round_=False):
    prst = "roundRect" if round_ else "rect"
    ln = ("<a:ln w=\"%d\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill></a:ln>"
          % (lw, _hex(line))) if line else "<a:ln><a:noFill/></a:ln>"
    fillxml = ("<a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>" % _hex(fill)
               if fill is not None else "<a:noFill/>")
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id=\"%d\" name=\"r\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "<a:prstGeom prst=\"%s\"><a:avLst/></a:prstGeom>%s%s</p:spPr>"
        "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
        % (_sp_id(), x, y, w, h, prst, fillxml, ln)
    )


def ellipse(cx, cy, rx, ry, fill, line=None, lw=9525):
    ln = ("<a:ln w=\"%d\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill></a:ln>"
          % (lw, _hex(line))) if line else "<a:ln><a:noFill/></a:ln>"
    fillxml = ("<a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>" % _hex(fill)
               if fill is not None else "<a:noFill/>")
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id=\"%d\" name=\"e\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>%s%s</p:spPr>"
        "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
        % (_sp_id(), int(cx - rx), int(cy - ry), int(2 * rx), int(2 * ry), fillxml, ln)
    )


def line(x1, y1, x2, y2, color, lw=9525):
    x, y = int(min(x1, x2)), int(min(y1, y2))
    cx, cy = int(abs(x2 - x1)), int(abs(y2 - y1))
    flipH = "1" if x2 < x1 else "0"
    flipV = "1" if y2 < y1 else "0"
    return (
        "<p:cxnSp><p:nvCxnSpPr><p:cNvPr id=\"%d\" name=\"l\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>"
        "<p:spPr><a:xfrm flipH=\"%s\" flipV=\"%s\"><a:off x=\"%d\" y=\"%d\"/>"
        "<a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom>"
        "<a:ln w=\"%d\" cap=\"rnd\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill></a:ln>"
        "</p:spPr></p:cxnSp>"
        % (_sp_id(), flipH, flipV, x, y, max(cx, 1), max(cy, 1), lw, _hex(color))
    )


def polyline(points, color, lw=9525, closed=False, fill=None):
    """Freeform (custom geometry) egri/siniq chiziq. points: [(x,y),...]."""
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    minx, miny = min(xs), min(ys)
    w = max(max(xs) - minx, 1)
    h = max(max(ys) - miny, 1)

    def pt(p):
        return "<a:pt x=\"%d\" y=\"%d\"/>" % (int(p[0] - minx), int(p[1] - miny))

    path = "<a:moveTo>%s</a:moveTo>" % pt(points[0])
    for p in points[1:]:
        path += "<a:lnTo>%s</a:lnTo>" % pt(p)
    if closed:
        path += "<a:close/>"
    ln = ("<a:ln w=\"%d\" cap=\"rnd\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>"
          "<a:round/></a:ln>" % (lw, _hex(color)))
    fillxml = ("<a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>" % _hex(fill)
               if fill is not None else "<a:noFill/>")
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id=\"%d\" name=\"f\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>"
        "<a:rect l=\"0\" t=\"0\" r=\"%d\" b=\"%d\"/>"
        "<a:pathLst><a:path w=\"%d\" h=\"%d\">%s</a:path></a:pathLst></a:custGeom>"
        "%s%s</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
        % (_sp_id(), w, h, w, h, w, h, path, fillxml, ln)
    )


def _arc_points(cx, cy, r, a0, a1, n=28):
    pts = []
    for i in range(n + 1):
        a = math.radians(a0 + (a1 - a0) * i / n)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def arc(cx, cy, r, a0, a1, color, lw=9525, n=28):
    return polyline(_arc_points(cx, cy, r, a0, a1, n), color, lw=lw, closed=False)


def _wrap_group(x, y, w, h, inner_xml):
    """Berilgan ekran to'rtburchagiga lokal (0..U) tizimni moslaydi."""
    gid = _sp_id()
    return (
        "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"%d\" name=\"illustration\"/>"
        "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>"
        "<a:xfrm><a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/>"
        "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "</p:grpSpPr>%s</p:grpSp>"
        % (gid, int(x), int(y), int(w), int(h), U, U, inner_xml)
    )


# ---------------------------------------------------------------------------
# Umumiy fon kartasi (barcha illyustratsiyalar bir xil professional ramkada)
# ---------------------------------------------------------------------------

def _card(pal):
    """Panel foni + nozik ramka + burchak aksent."""
    s = rect(0, 0, U, U, pal["panel"], line=pal["accent"], lw=12700, round_=True)
    # yuqori-chap burchak aksenti
    s += line(3000, 7000, 20000, 7000, pal["accent"], lw=19050)
    s += line(3000, 7000, 3000, 24000, pal["accent"], lw=19050)
    # past-o'ng burchak aksenti
    s += line(80000, 93000, 97000, 93000, pal["accent2"], lw=19050)
    s += line(97000, 76000, 97000, 93000, pal["accent2"], lw=19050)
    return s


# ---------------------------------------------------------------------------
# Barmoq izi ridge'lari (bir nechta slaydda qayta ishlatiladi)
# ---------------------------------------------------------------------------

def _fingerprint_ridges(cx, cy, color, lw=5000, loops=6, minutiae_color=None):
    """Konsentrik/looping ridge egri chiziqlari (barmoq izi stilizatsiyasi)."""
    out = ""
    minutiae = []
    for k in range(loops):
        r = 7000 + k * 6000
        # yuqoridan ochiq loop (barmoq izi uslubi)
        pts = _arc_points(cx, cy + k * 900, r, 200, 520, n=30)
        out += polyline(pts, color, lw=lw)
        if k in (1, 3) and minutiae_color:
            # ridge tugashi (minutiae) nuqtasi
            end = pts[-1]
            minutiae.append(end)
    if minutiae_color:
        for mx, my in minutiae:
            out += ellipse(mx, my, 1600, 1600, minutiae_color)
    return out


# ---------------------------------------------------------------------------
# Slayd-ma-slayd motivlar
# ---------------------------------------------------------------------------

def motif_fingerprint_scan(pal):
    """1-slayd (titul): stilizatsiya qilingan barmoq izi + skan chizig'i."""
    s = _card(pal)
    cx, cy = 50000, 52000
    s += _fingerprint_ridges(cx, cy, pal["accent"], lw=4200, loops=7,
                             minutiae_color=pal["ink"])
    # skan chizig'i (gorizontal nurli chiziq)
    s += rect(14000, 46000, 72000, 2600, pal["accent2"])
    s += rect(14000, 45200, 72000, 700, pal["ink"])
    # ramkadagi burchak belgilari (skaner oynasi)
    for (bx, by) in [(14000, 16000), (78000, 16000), (14000, 80000), (78000, 80000)]:
        s += line(bx, by, bx + 8000, by, pal["accent"], lw=9000)
        s += line(bx, by, bx, by + 8000, pal["accent"], lw=9000)
    return s


def motif_agenda(pal):
    """2-slayd (reja): raqamlangan ro'yxat ikonkasi."""
    s = _card(pal)
    ys = [22000, 42000, 62000, 82000]
    cols = [pal["accent"], pal["accent2"], pal["accent"], pal["ink"]]
    for i, y in enumerate(ys):
        # raqam doirasi
        s += ellipse(24000, y, 8200, 8200, None, line=cols[i], lw=9000)
        s += num_text(16500, y - 7000, 15000, 14000, str(i + 1), pal, cols[i])
        # ro'yxat chizig'i
        s += rect(40000, y - 2600, 42000, 5200, None, line=cols[i], lw=6000, round_=True)
    return s


def motif_minutiae(pal):
    """3-slayd (biologiya): ridge'lar + ta'kidlangan minutiae nuqtalari."""
    s = _card(pal)
    # parallel egri ridge chiziqlari
    for i in range(7):
        base = 20000 + i * 9000
        pts = [(14000, base),
               (34000, base - 4000 - i * 300),
               (54000, base + 3000),
               (74000, base - 2000),
               (86000, base + 1500)]
        s += polyline(pts, pal["accent"], lw=4000)
    # minutiae: tugash va ayri (bifurcation) nuqtalari
    mpts = [(34000, 25000), (54000, 56000), (74000, 38000), (46000, 74000)]
    for mx, my in mpts:
        s += ellipse(mx, my, 3200, 3200, None, line=pal["ink"], lw=7000)
        s += ellipse(mx, my, 1100, 1100, pal["accent2"])
    return s


def motif_hashing(pal):
    """4-slayd (raqamlashtirish): barmoq izi -> ikkilik/hash raqamlari."""
    s = _card(pal)
    # chapda qisman barmoq izi
    s += _fingerprint_ridges(28000, 50000, pal["accent"], lw=3800, loops=5)
    # o'ngda strelka
    s += polyline([(50000, 50000), (62000, 50000)], pal["ink"], lw=6000)
    s += polyline([(58000, 46000), (62000, 50000), (58000, 54000)], pal["ink"], lw=6000)
    # 0/1 panjara (ikkilik kod)
    bits = "1011 0110 1001 0101 1100".split()
    for r, chunk in enumerate(bits):
        s += num_text(64000, 20000 + r * 12000, 32000, 11000, chunk, pal,
                      pal["accent2"] if r % 2 == 0 else pal["accent"], size=1400)
    # hash satri pastda
    s += num_text(16000, 86000, 68000, 10000, "SHA: A7F90B...", pal, pal["ink"], size=1300)
    return s


def motif_optical_light(pal):
    """5-slayd (optik): LED yorug'lik nurlari sirtga urilib qaytadi."""
    s = _card(pal)
    # LED manbai (pastda chap)
    led = (22000, 82000)
    s += ellipse(led[0], led[1], 5000, 5000, pal["accent"])
    # yorug'lik nurlari yuqoriga/o'ngga
    targets = [(45000, 24000), (56000, 28000), (67000, 33000), (78000, 40000)]
    for tx, ty in targets:
        s += line(led[0], led[1], tx, ty, pal["accent"], lw=4000)
        # qaytgan nur (reflection)
        s += line(tx, ty, tx + 14000, ty + 20000, pal["accent2"], lw=3200)
    # sirt (barmoq) chizig'i yuqorida
    s += polyline([(36000, 20000), (54000, 24000), (72000, 20000), (86000, 26000)],
                  pal["ink"], lw=5000)
    return s


def motif_cmos_grid(pal):
    """6-slayd (CMOS): piksel panjarasi 2D barmoq izi tasvirini hosil qiladi."""
    s = _card(pal)
    n = 9
    x0, y0, step = 20000, 18000, 7000
    cx, cy = x0 + n * step / 2, y0 + n * step / 2
    for r in range(n):
        for c in range(n):
            px = x0 + c * step
            py = y0 + r * step
            # markazga yaqin piksellar yoqilgan (barmoq izi shakli)
            d = math.hypot(px - cx + 2000, py - cy)
            on = (int(d / 6000) + c) % 2 == 0 and d < 30000
            col = pal["accent"] if on else None
            lc = pal["accent2"] if on else pal["grid"]
            s += rect(px, py, step - 1800, step - 1800, col, line=lc, lw=4000)
    return s


def motif_ultrasound(pal):
    """7-slayd (ultratovush): emitterdan konsentrik ultratovush impulslari."""
    s = _card(pal)
    emitter = (24000, 50000)
    s += rect(14000, 40000, 10000, 20000, pal["accent2"], round_=True)
    # tovush to'lqinlari (ochiq yoylar o'ngga)
    for i, r in enumerate(range(16000, 64000, 11000)):
        col = pal["accent"] if i % 2 == 0 else pal["accent2"]
        s += arc(emitter[0], emitter[1], r, -58, 58, col, lw=4200)
    # nishon barmoq (o'ngda)
    s += _fingerprint_ridges(82000, 50000, pal["ink"], lw=3200, loops=3)
    return s


def motif_echo_3d(pal):
    """8-slayd (3D xarita): echo to'lqinlari + 3D relyef balandlik profili."""
    s = _card(pal)
    # echo yoylari (chapda)
    for i, r in enumerate(range(12000, 44000, 10000)):
        col = pal["accent"] if i % 2 == 0 else pal["accent2"]
        s += arc(20000, 32000, r, -70, 70, col, lw=3600)
    # 3D relyef / balandlik profili (pastda) - ridge ko'ndalang kesimi
    base = 86000
    pts = [(14000, base)]
    xv = 14000
    for i in range(10):
        xv += 7200
        peak = base - (9000 if i % 2 == 0 else 3000)
        pts.append((xv - 3600, peak))
        pts.append((xv, base - 1500))
    pts.append((86000, base))
    s += polyline(pts, pal["ink"], lw=4200, closed=True, fill=pal["panel2"])
    # balandlik ustunlari (height-map)
    for i in range(8):
        hx = 20000 + i * 8000
        hh = 10000 + (abs(4 - i) * 3000)
        s += rect(hx, 70000 - hh, 5000, hh, pal["accent2"] if i % 2 else pal["accent"])
    return s


def motif_security_shield(pal):
    """9-slayd (xulosa): yorug'lik + tovush motivlarini birlashtirgan qalqon/qulf."""
    s = _card(pal)
    # qalqon konturi
    shield = [(50000, 14000), (80000, 24000), (80000, 54000),
              (50000, 86000), (20000, 54000), (20000, 24000)]
    s += polyline(shield, pal["accent"], lw=6000, closed=True, fill=pal["panel2"])
    # qulf tanasi
    s += rect(40000, 50000, 20000, 18000, pal["accent2"], round_=True)
    # qulf halqasi
    s += arc(50000, 50000, 9000, 180, 360, pal["ink"], lw=6000)
    # yorug'lik nuri (chap) + tovush yoyi (o'ng)
    s += line(14000, 20000, 30000, 36000, pal["accent"], lw=4000)
    s += arc(86000, 36000, 10000, 90, 270, pal["accent"], lw=4000)
    return s


def motif_badge(pal):
    """10-slayd (rahmat): toza yakuniy belgi - barmoq izi + qulf nishoni."""
    s = _card(pal)
    cx, cy = 50000, 46000
    s += ellipse(cx, cy, 26000, 26000, None, line=pal["accent2"], lw=9000)
    s += _fingerprint_ridges(cx, cy - 1000, pal["accent"], lw=3600, loops=4)
    # kichik qulf pastda
    s += rect(44000, 80000, 12000, 9000, pal["accent2"], round_=True)
    s += arc(50000, 80000, 4500, 180, 360, pal["ink"], lw=4000)
    return s


# ---------------------------------------------------------------------------
# Guruh ichidagi matn (raqam/hash) - Times New Roman, >= 12pt talab guruh
# ichidagi dekorativ raqamlar uchun ham saqlanadi (14pt+).
# ---------------------------------------------------------------------------

def num_text(x, y, w, h, text, pal, color, size=1800):
    """Guruh ichidagi dekorativ matn (raqam/hash). sz yuzlik pt (1800 = 18pt)."""
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id=\"%d\" name=\"t\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        "<p:txBody><a:bodyPr wrap=\"none\" anchor=\"ctr\"/><a:lstStyle/>"
        "<a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"uz-UZ\" sz=\"%d\" b=\"1\" dirty=\"0\">"
        "<a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>"
        "<a:latin typeface=\"Times New Roman\"/><a:cs typeface=\"Times New Roman\"/>"
        "<a:ea typeface=\"Times New Roman\"/></a:rPr><a:t>%s</a:t></a:r></a:p>"
        "</p:txBody></p:sp>"
        % (_sp_id(), int(x), int(y), int(w), int(h), size, _hex(color), text)
    )


# Slayd indeksi -> motiv funksiyasi
MOTIFS = [
    motif_fingerprint_scan,  # 1 titul
    motif_agenda,            # 2 reja
    motif_minutiae,          # 3 biologiya
    motif_hashing,           # 4 raqamlashtirish
    motif_optical_light,     # 5 optik
    motif_cmos_grid,         # 6 CMOS
    motif_ultrasound,        # 7 ultratovush
    motif_echo_3d,           # 8 3D xarita
    motif_security_shield,   # 9 xulosa
    motif_badge,             # 10 rahmat
]


def build_illustration(idx, x, y, w, h, pal):
    """idx (0..9) uchun ekran to'rtburchagiga joylangan vektor rasm XML'i."""
    inner = MOTIFS[idx % len(MOTIFS)](pal)
    return _wrap_group(x, y, w, h, inner)


# Illyustratsiya uchun palitra (slides_content ranglaridan tuziladi)
def make_palette(COLOR_PANEL, COLOR_ACCENT, COLOR_ACCENT2, COLOR_TITLE, COLOR_BG_DARK):
    return {
        "panel": COLOR_PANEL,
        "panel2": COLOR_BG_DARK,
        "accent": COLOR_ACCENT,
        "accent2": COLOR_ACCENT2,
        "ink": COLOR_TITLE,
        "grid": (0x24, 0x3B, 0x55),
    }
