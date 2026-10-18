# -*- coding: utf-8 -*-
"""Mavzuga mos (mikroelektronika / xotira / kvant fizikasi) illyustratsiyalar.

Faqat standart kutubxona (zlib, struct, math) ishlatiladi - PIL ham, tarmoq
ham kerak emas. Har bir slayd uchun o'z mazmuniga mos protsedurali rasm
chiziladi (USB konnektori, NAND katak panjarasi, kremniy kristali, suzuvchi
zulfin kesimi, elektronlarning qamalishi, kvant tunnellanish, o'chirish/yeyilish,
xulosa va yakun). Shunday qilib .pptx baribir yaroqli va "mavzuga aloqador"
rasmlarga ega bo'ladi.

PNG yozish usuli oldingi kod bilan bir xil: 8-bitli truecolor (color type 2),
har qator oldidan filter bayti 0, zlib bilan siqish, qo'lda IHDR/IDAT/IEND
chunklar + CRC32.
"""
import math
import struct
import zlib

# Slayd dizayni bilan bir xil rang palitrasidan foydalanamiz.
from slides_content import (
    COLOR_BG_DARK,
    COLOR_PANEL,
    COLOR_ACCENT,
    COLOR_ACCENT2,
    COLOR_TITLE,
    COLOR_BODY,
)

W, H = 1000, 1000

# Qo'shimcha tematik ranglar (palitra ohangida)
NAVY = COLOR_BG_DARK           # (11, 30, 51)
PANEL = COLOR_PANEL            # (18, 42, 69)
CYAN = COLOR_ACCENT            # (0, 194, 209)
BLUE = COLOR_ACCENT2           # (59, 130, 246)
WHITE = COLOR_TITLE            # (255, 255, 255)
BODY = COLOR_BODY             # (226, 232, 240)
GOLD = (251, 191, 36)         # metall konnektor / mis izlar
SILVER = (176, 190, 205)      # metall korpus
GREEN = (52, 211, 153)        # mantiqiy "1" / toza oqim
RED = (248, 113, 113)         # ogohlantirish / yeyilish
VIOLET = (167, 139, 250)      # kvant urg'u


# ---------------------------------------------------------------------------
# Chizish asboblari (stdlib, RGB piksel buferi)
# ---------------------------------------------------------------------------
class Canvas:
    """Oddiy mutable RGB piksel buferi va chizish yordamchilari."""

    def __init__(self, w=W, h=H, bg=NAVY):
        self.w = w
        self.h = h
        self.buf = bytearray(w * h * 3)
        self.fill(bg)

    # --- past daraja ---
    def fill(self, color):
        r, g, b = color
        row = bytes((r, g, b)) * self.w
        for y in range(self.h):
            off = y * self.w * 3
            self.buf[off:off + self.w * 3] = row

    def _set(self, x, y, color):
        if 0 <= x < self.w and 0 <= y < self.h:
            o = (y * self.w + x) * 3
            self.buf[o] = color[0]
            self.buf[o + 1] = color[1]
            self.buf[o + 2] = color[2]

    def _get(self, x, y):
        o = (y * self.w + x) * 3
        return self.buf[o], self.buf[o + 1], self.buf[o + 2]

    def blend(self, x, y, color, a):
        """color ni (x,y) ustiga alpha a (0..1) bilan aralashtirish."""
        if not (0 <= x < self.w and 0 <= y < self.h):
            return
        if a <= 0:
            return
        if a >= 1:
            self._set(x, y, color)
            return
        o = (y * self.w + x) * 3
        ia = 1.0 - a
        self.buf[o] = int(self.buf[o] * ia + color[0] * a)
        self.buf[o + 1] = int(self.buf[o + 1] * ia + color[1] * a)
        self.buf[o + 2] = int(self.buf[o + 2] * ia + color[2] * a)

    # --- gradient fon ---
    def vgradient(self, c1, c2):
        for y in range(self.h):
            t = y / (self.h - 1)
            r = int(c1[0] + (c2[0] - c1[0]) * t)
            g = int(c1[1] + (c2[1] - c1[1]) * t)
            b = int(c1[2] + (c2[2] - c1[2]) * t)
            row = bytes((r, g, b)) * self.w
            off = y * self.w * 3
            self.buf[off:off + self.w * 3] = row

    def dgradient(self, c1, c2):
        for y in range(self.h):
            ty = y / (self.h - 1)
            for x in range(self.w):
                t = (x / (self.w - 1) + ty) / 2.0
                o = (y * self.w + x) * 3
                self.buf[o] = int(c1[0] + (c2[0] - c1[0]) * t)
                self.buf[o + 1] = int(c1[1] + (c2[1] - c1[1]) * t)
                self.buf[o + 2] = int(c1[2] + (c2[2] - c1[2]) * t)

    def radial_glow(self, cx, cy, radius, color, strength=0.55):
        """Markazdan tashqariga so'nuvchi yorug'lik dog'i."""
        r2 = radius * radius
        y0 = max(0, int(cy - radius))
        y1 = min(self.h, int(cy + radius) + 1)
        x0 = max(0, int(cx - radius))
        x1 = min(self.w, int(cx + radius) + 1)
        for y in range(y0, y1):
            for x in range(x0, x1):
                d2 = (x - cx) ** 2 + (y - cy) ** 2
                if d2 <= r2:
                    t = 1.0 - math.sqrt(d2) / radius
                    self.blend(x, y, color, strength * t * t)

    # --- shakllar ---
    def fill_rect(self, x0, y0, x1, y1, color):
        x0, x1 = sorted((int(x0), int(x1)))
        y0, y1 = sorted((int(y0), int(y1)))
        x0 = max(0, x0); y0 = max(0, y0)
        x1 = min(self.w, x1); y1 = min(self.h, y1)
        for y in range(y0, y1):
            o = (y * self.w + x0) * 3
            for x in range(x0, x1):
                self.buf[o] = color[0]
                self.buf[o + 1] = color[1]
                self.buf[o + 2] = color[2]
                o += 3

    def rect_outline(self, x0, y0, x1, y1, color, t=3):
        self.fill_rect(x0, y0, x1, y0 + t, color)
        self.fill_rect(x0, y1 - t, x1, y1, color)
        self.fill_rect(x0, y0, x0 + t, y1, color)
        self.fill_rect(x1 - t, y0, x1, y1, color)

    def rounded_rect(self, x0, y0, x1, y1, radius, color):
        x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
        r = int(radius)
        self.fill_rect(x0 + r, y0, x1 - r, y1, color)
        self.fill_rect(x0, y0 + r, x1, y1 - r, color)
        for cx, cy in ((x0 + r, y0 + r), (x1 - r, y0 + r),
                       (x0 + r, y1 - r), (x1 - r, y1 - r)):
            self.filled_circle(cx, cy, r, color)

    def filled_circle(self, cx, cy, radius, color):
        cx, cy = float(cx), float(cy)
        r = float(radius)
        y0 = max(0, int(cy - r - 1))
        y1 = min(self.h, int(cy + r + 2))
        for y in range(y0, y1):
            for x in range(max(0, int(cx - r - 1)), min(self.w, int(cx + r + 2))):
                d = math.hypot(x - cx, y - cy)
                if d <= r - 0.5:
                    self._set(x, y, color)
                elif d <= r + 0.5:  # chekka silliqlash
                    self.blend(x, y, color, r + 0.5 - d)

    def circle_outline(self, cx, cy, radius, color, t=3):
        cx, cy = float(cx), float(cy)
        r = float(radius)
        y0 = max(0, int(cy - r - t))
        y1 = min(self.h, int(cy + r + t + 1))
        inner = r - t
        for y in range(y0, y1):
            for x in range(max(0, int(cx - r - t)), min(self.w, int(cx + r + t + 1))):
                d = math.hypot(x - cx, y - cy)
                if inner - 0.5 <= d <= r + 0.5:
                    a = min(1.0, min(d - (inner - 0.5), (r + 0.5) - d))
                    self.blend(x, y, color, max(0.0, min(1.0, a)))

    def draw_line(self, x0, y0, x1, y1, color, t=2):
        """Qalinligi t bo'lgan chiziq (masofa asosida silliqlangan)."""
        x0, y0, x1, y1 = float(x0), float(y0), float(x1), float(y1)
        dx, dy = x1 - x0, y1 - y0
        length = math.hypot(dx, dy)
        if length < 1e-6:
            self.filled_circle(x0, y0, t / 2.0, color)
            return
        nx, ny = dx / length, dy / length
        half = t / 2.0
        bx0 = int(min(x0, x1) - half - 1)
        bx1 = int(max(x0, x1) + half + 2)
        by0 = int(min(y0, y1) - half - 1)
        by1 = int(max(y0, y1) + half + 2)
        for y in range(max(0, by0), min(self.h, by1)):
            for x in range(max(0, bx0), min(self.w, bx1)):
                px, py = x - x0, y - y0
                proj = px * nx + py * ny
                if proj < 0:
                    dist = math.hypot(px, py)
                elif proj > length:
                    dist = math.hypot(x - x1, y - y1)
                else:
                    dist = abs(px * (-ny) + py * nx)
                if dist <= half - 0.5:
                    self._set(x, y, color)
                elif dist <= half + 0.5:
                    self.blend(x, y, color, half + 0.5 - dist)

    def arrow(self, x0, y0, x1, y1, color, t=4, head=16):
        self.draw_line(x0, y0, x1, y1, color, t)
        ang = math.atan2(y1 - y0, x1 - x0)
        for da in (math.radians(150), math.radians(-150)):
            hx = x1 + head * math.cos(ang + da)
            hy = y1 + head * math.sin(ang + da)
            self.draw_line(x1, y1, hx, hy, color, t)

    def to_png(self):
        rows = bytearray()
        stride = self.w * 3
        for y in range(self.h):
            rows.append(0)
            o = y * stride
            rows += self.buf[o:o + stride]

        def chunk(tag, data):
            c = tag + data
            return (struct.pack(">I", len(data)) + c +
                    struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF))

        sig = b"\x89PNG\r\n\x1a\n"
        ihdr = struct.pack(">IIBBBBB", self.w, self.h, 8, 2, 0, 0, 0)
        idat = zlib.compress(bytes(rows), 9)
        return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def _bg(cv, top=None, bottom=None):
    """Umumiy to'q ko'k vertikal gradient fon + yengil vinetka."""
    top = top or (14, 36, 60)
    bottom = bottom or (6, 18, 32)
    cv.vgradient(top, bottom)


def _draw_bits(cv, x, y, bits, color, cell=14, gap=4):
    """0/1 raqamlarini oddiy segmentli piksellar bilan chizish."""
    glyphs = {
        "0": ["111", "101", "101", "101", "111"],
        "1": ["010", "110", "010", "010", "111"],
    }
    cx = x
    for ch in bits:
        g = glyphs.get(ch)
        if g:
            for ry, rowpat in enumerate(g):
                for rx, c in enumerate(rowpat):
                    if c == "1":
                        px = cx + rx * cell
                        py = y + ry * cell
                        cv.fill_rect(px, py, px + cell - 1, py + cell - 1, color)
        cx += 3 * cell + gap


def _traces(cv, color, step=70, t=2, dots=True):
    """Mikrosxema plata izlari (fon bezagi)."""
    for gx in range(step, cv.w, step):
        cv.draw_line(gx, 0, gx, cv.h, color, t)
    for gy in range(step, cv.h, step):
        cv.draw_line(0, gy, cv.w, gy, color, t)
    if dots:
        for gx in range(step, cv.w, step):
            for gy in range(step, cv.h, step):
                cv.filled_circle(gx, gy, 3, color)


# ---------------------------------------------------------------------------
# Slayd illyustratsiyalari (0..9)
# ---------------------------------------------------------------------------
def slide0():
    """Titul: USB fleshka silueti + chip + plata izlari."""
    cv = Canvas()
    _bg(cv)
    cv.radial_glow(300, 300, 520, CYAN, 0.18)
    # orqa fon izlari
    _traces(cv, (24, 54, 86), step=90, t=1, dots=False)

    # USB fleshka korpusi (gorizontal)
    bx0, by0, bx1, by1 = 250, 430, 720, 570
    cv.rounded_rect(bx0, by0, bx1, by1, 34, PANEL)
    cv.rect_outline(bx0, by0, bx1, by1, CYAN, 4)
    # metall USB konnektor (chapda)
    cx0, cx1 = 120, 250
    cv.fill_rect(cx0, 468, cx1, 532, SILVER)
    cv.rect_outline(cx0, 468, cx1, 532, (90, 104, 120), 3)
    # konnektor ichidagi mis kontaktlar
    for i in range(4):
        yy = 480 + i * 13
        cv.fill_rect(cx0 + 14, yy, cx0 + 70, yy + 7, GOLD)
    # korpus ichidagi chip
    cv.fill_rect(430, 468, 540, 532, (10, 24, 40))
    cv.rect_outline(430, 468, 540, 532, CYAN, 3)
    for i in range(5):  # chip oyoqlari
        cv.fill_rect(430, 472 + i * 12, 420, 478 + i * 12, GOLD)
        cv.fill_rect(540, 472 + i * 12, 550, 478 + i * 12, GOLD)
    # LED nuqta
    cv.filled_circle(690, 500, 10, CYAN)
    cv.radial_glow(690, 500, 40, CYAN, 0.6)

    # elektr zaryad uchqunlari
    for (sx, sy) in ((780, 300), (840, 650), (300, 700)):
        cv.filled_circle(sx, sy, 7, GOLD)
        cv.radial_glow(sx, sy, 36, GOLD, 0.5)

    # pastki aksent chizig'i
    cv.fill_rect(0, 760, cv.w, 772, CYAN)
    cv.fill_rect(0, 776, cv.w, 782, BLUE)
    return cv.to_png()


def slide1():
    """Reja: raqamlangan 3 reja + xulosa tugunli oqim diagrammasi."""
    cv = Canvas()
    _bg(cv, (12, 28, 52), (6, 16, 30))
    cv.radial_glow(500, 500, 560, BLUE, 0.14)

    labels = ["1", "2", "3", "4"]
    colors = [CYAN, BLUE, CYAN, GREEN]
    ys = [180, 380, 580, 790]
    node_x = 300
    for i, (lab, col, y) in enumerate(zip(labels, colors, ys)):
        if i > 0:
            cv.draw_line(node_x, ys[i - 1] + 60, node_x, y - 60, (60, 90, 130), 5)
        cv.filled_circle(node_x, y, 62, PANEL)
        cv.circle_outline(node_x, y, 62, col, 6)
        # raqam glifi (katta)
        _draw_bits_num(cv, node_x - 20, y - 42, lab, col)
        # yon panel (reja matni o'rnida blok)
        cv.rounded_rect(node_x + 110, y - 42, 820, y + 42, 18, (20, 44, 72))
        cv.fill_rect(node_x + 130, y - 10, 760, y - 4, col)
        cv.fill_rect(node_x + 130, y + 10, 680, y + 16, (70, 100, 140))
    return cv.to_png()


def _draw_bits_num(cv, x, y, ch, color, cell=14):
    """1..4 raqamlarini blokli ko'rinishda chizish."""
    glyphs = {
        "1": ["010", "110", "010", "010", "111"],
        "2": ["111", "001", "111", "100", "111"],
        "3": ["111", "001", "111", "001", "111"],
        "4": ["101", "101", "111", "001", "001"],
    }
    g = glyphs.get(ch, glyphs["1"])
    for ry, rowpat in enumerate(g):
        for rx, c in enumerate(rowpat):
            if c == "1":
                px = x + rx * cell
                py = y + ry * cell
                cv.fill_rect(px, py, px + cell - 1, py + cell - 1, color)


def slide2():
    """NAND arxitekturasi: bitline/wordline panjarasi + tranzistor kataklari."""
    cv = Canvas()
    _bg(cv)
    cv.radial_glow(500, 500, 560, CYAN, 0.12)

    x0, y0 = 150, 180
    cols, rows = 6, 6
    step = 115
    # wordline (gorizontal) va bitline (vertikal) o'tkazgichlari
    for r in range(rows + 1):
        yy = y0 + r * step
        cv.draw_line(x0 - 40, yy, x0 + cols * step - step + 40, yy, BLUE, 3)
    for c in range(cols):
        xx = x0 + c * step
        cv.draw_line(xx, y0 - 40, xx, y0 + rows * step - step + 40, (60, 90, 130), 3)

    # xotira kataklari (tranzistorlar)
    for r in range(rows):
        for c in range(cols):
            xx = x0 + c * step
            yy = y0 + r * step
            cv.rounded_rect(xx - 34, yy - 26, xx + 34, yy + 26, 10, (14, 32, 54))
            cv.rect_outline(xx - 34, yy - 26, xx + 34, yy + 26, CYAN, 2)
            # ba'zi kataklarda zaryad bor (to'ldirilgan)
            filled = (r + c) % 3 == 0
            if filled:
                cv.filled_circle(xx, yy, 12, CYAN)
                cv.radial_glow(xx, yy, 26, CYAN, 0.4)
            else:
                cv.circle_outline(xx, yy, 12, (90, 120, 160), 2)
    # yon yorliq chiziqlari (WL/BL urg'u)
    cv.fill_rect(x0 - 70, y0 - 60, x0 - 58, y0 + rows * step - step + 60, CYAN)
    cv.fill_rect(x0 - 70, y0 - 60, x0 + 60, y0 - 48, BLUE)
    return cv.to_png()


def slide3():
    """Kremniy kristali / wafer + kattalashtirilgan zич katak panjarasi."""
    cv = Canvas()
    _bg(cv, (10, 26, 46), (5, 14, 28))
    # kremniy disk (wafer)
    cv.filled_circle(330, 420, 230, (28, 52, 80))
    cv.circle_outline(330, 420, 230, SILVER, 5)
    cv.radial_glow(260, 350, 180, CYAN, 0.22)
    # waferdagi die panjarasi
    for gy in range(-3, 4):
        for gx in range(-3, 4):
            dx = 330 + gx * 62
            dy = 420 + gy * 62
            if math.hypot(dx - 330, dy - 420) < 210:
                cv.fill_rect(dx - 26, dy - 26, dx + 26, dy + 26, (18, 40, 66))
                cv.rect_outline(dx - 26, dy - 26, dx + 26, dy + 26, (70, 110, 150), 1)

    # kattalashtirish lupasi -> zich mitti kataklar (o'ng pastda)
    cv.draw_line(470, 540, 650, 700, CYAN, 4)
    mx0, my0, mx1, my1 = 620, 640, 900, 900
    cv.rounded_rect(mx0, my0, mx1, my1, 20, (8, 20, 36))
    cv.rect_outline(mx0, my0, mx1, my1, CYAN, 4)
    cell = 14
    for yy in range(my0 + 20, my1 - 10, cell):
        for xx in range(mx0 + 20, mx1 - 10, cell):
            on = ((xx // cell) * 7 + (yy // cell) * 13) % 5 < 2
            col = CYAN if on else (36, 60, 90)
            cv.fill_rect(xx, yy, xx + cell - 3, yy + cell - 3, col)
    # "milliardlab" urg'u nuqtalari
    for _ in range(0):
        pass
    return cv.to_png()


def slide4():
    """Floating Gate tranzistor kesimi: qatlamlar + source/drain."""
    cv = Canvas()
    _bg(cv, (16, 30, 50), (8, 18, 32))

    left, right = 140, 860
    # substrat (eng past)
    sub_top = 700
    cv.fill_rect(left, sub_top, right, 860, (34, 48, 40))
    cv.rect_outline(left, sub_top, right, 860, (90, 110, 90), 2)
    # source va drain (n+ hududlar)
    cv.fill_rect(left, 660, left + 150, sub_top, BLUE)
    cv.fill_rect(right - 150, 660, right, sub_top, BLUE)
    # tunnel oksid (yupqa)
    cv.fill_rect(left + 150, 640, right - 150, 660, (120, 130, 150))
    # suzuvchi zulfin (Floating Gate) - qamalgan
    cv.fill_rect(left + 150, 560, right - 150, 640, GOLD)
    cv.rect_outline(left + 150, 560, right - 150, 640, (200, 150, 40), 2)
    # qamalgan elektronlar (nuqtalar)
    for i in range(7):
        cv.filled_circle(left + 230 + i * 70, 600, 10, (20, 30, 50))
        cv.filled_circle(left + 230 + i * 70, 600, 6, CYAN)
    # oksid izolyator qatlam
    cv.fill_rect(left + 150, 500, right - 150, 560, (90, 120, 160))
    # boshqaruvchi zulfin (Control Gate)
    cv.fill_rect(left + 150, 420, right - 150, 500, SILVER)
    cv.rect_outline(left + 150, 420, right - 150, 500, WHITE, 2)
    # terminal chiziqlari
    cv.draw_line(500, 420, 500, 300, CYAN, 5)
    cv.arrow(500, 360, 500, 300, CYAN, 5, 18)
    # qatlam yorliq nuqtalari (o'ng tomonda)
    for yy, col in ((460, SILVER), (530, (90, 120, 160)), (600, GOLD),
                    (650, (120, 130, 150)), (730, (90, 110, 90))):
        cv.filled_circle(right + 30, yy, 10, col)
    return cv.to_png()


def slide5():
    """0 va 1: zulfin ikki holatda - elektron qamalgan (0) vs bo'sh (1)."""
    cv = Canvas()
    _bg(cv)
    midx = 500
    cv.draw_line(midx, 120, midx, 880, (50, 76, 110), 3)

    def gate(cx, trapped, bit, bit_col):
        # tranzistor kanali
        cv.fill_rect(cx - 150, 560, cx + 150, 640, (34, 48, 40))
        cv.fill_rect(cx - 150, 540, cx - 90, 560, BLUE)   # source
        cv.fill_rect(cx + 90, 540, cx + 150, 560, BLUE)   # drain
        # floating gate
        cv.rounded_rect(cx - 120, 420, cx + 120, 520, 14, GOLD)
        cv.rect_outline(cx - 120, 420, cx + 120, 520, (200, 150, 40), 2)
        if trapped:
            for i in range(5):
                ex = cx - 80 + i * 40
                cv.filled_circle(ex, 470, 13, (20, 30, 50))
                cv.filled_circle(ex, 470, 8, CYAN)
            # oqim to'silgan -> X
            cv.draw_line(cx - 40, 600, cx + 40, 680, RED, 6)
            cv.draw_line(cx + 40, 600, cx - 40, 680, RED, 6)
        else:
            # bo'sh zulfin -> oqim bemalol o'tadi (strelka)
            cv.arrow(cx - 150, 680, cx + 150, 680, GREEN, 6, 20)
        # bit raqami
        _draw_bits(cv, cx - 24, 180, bit, bit_col, cell=22, gap=6)
        cv.circle_outline(cx, 230, 110, bit_col, 5)

    gate(275, True, "0", CYAN)
    gate(725, False, "1", GREEN)
    return cv.to_png()


def slide6():
    """Kvant tunnellanish: elektronlar oksid to'sig'idan kuchli maydon ostida o'tishi."""
    cv = Canvas()
    _bg(cv, (24, 18, 46), (10, 10, 26))
    cv.radial_glow(500, 460, 560, VIOLET, 0.18)

    # oksid to'sig'i (vertikal panel)
    bx0, bx1 = 460, 560
    cv.fill_rect(bx0, 160, bx1, 820, (120, 130, 170))
    cv.rect_outline(bx0, 160, bx1, 820, WHITE, 3)
    # "to'siq" potensial grafigi (yuqorida)
    cv.draw_line(120, 760, bx0, 760, CYAN, 3)
    cv.draw_line(bx0, 760, bx0, 300, CYAN, 3)
    cv.draw_line(bx0, 300, bx1, 300, CYAN, 3)
    cv.draw_line(bx1, 300, bx1, 760, CYAN, 3)
    cv.draw_line(bx1, 760, 880, 760, CYAN, 3)

    # kuchli maydon strelkalari (chapdan o'ngga)
    for yy in (240, 560, 700):
        cv.arrow(140, yy, 440, yy, VIOLET, 4, 16)

    # elektronlar: ba'zilari to'siqdan o'tyapti (tunnel)
    electrons = [(250, 420), (330, 500), (420, 460),
                 (bx0 + 50, 460), (650, 440), (760, 500)]
    for ex, ey in electrons:
        cv.filled_circle(ex, ey, 16, (20, 20, 40))
        cv.filled_circle(ex, ey, 11, CYAN)
        cv.radial_glow(ex, ey, 40, CYAN, 0.45)
    # tunnel yo'li (nuqtali chiziq to'siq ichidan)
    for t in range(0, 11):
        px = 430 + t * 10
        cv.filled_circle(px, 460, 4, GOLD)
    return cv.to_png()


def slide7():
    """O'chirish va yeyilish: teskari maydon elektronlarni tortadi + yorilgan oksid + sikl hisoblagich."""
    cv = Canvas()
    _bg(cv, (12, 26, 44), (6, 14, 26))

    # zulfin bloki
    cv.rounded_rect(300, 420, 700, 540, 16, GOLD)
    cv.rect_outline(300, 420, 700, 540, (200, 150, 40), 2)
    # elektronlar tashqariga so'rilyapti (yuqoriga strelkalar)
    for i in range(5):
        ex = 360 + i * 70
        cv.filled_circle(ex, 480, 9, CYAN)
        cv.arrow(ex, 420, ex, 300, CYAN, 4, 14)
    # teskari kuchlanish yorlig'i (yuqori panel)
    cv.rounded_rect(360, 210, 640, 290, 14, PANEL)
    cv.fill_rect(390, 245, 610, 252, RED)

    # yorilgan oksid qatlam (yeyilish) - pastki zigzag
    cv.fill_rect(300, 560, 700, 620, (90, 120, 160))
    zx = 300
    pts = [(300, 590)]
    while zx < 700:
        zx += 40
        pts.append((zx, 575 if (zx // 40) % 2 else 605))
    for i in range(len(pts) - 1):
        cv.draw_line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], RED, 4)

    # sikl hisoblagich (o'ng pastda): 10,000 - 100,000
    cv.rounded_rect(620, 700, 900, 860, 18, (18, 36, 60))
    cv.rect_outline(620, 700, 900, 860, RED, 3)
    _draw_bits(cv, 650, 740, "10", RED, cell=16, gap=6)
    cv.fill_rect(650, 810, 870, 818, (120, 60, 60))
    cv.fill_rect(650, 830, 760, 838, (120, 60, 60))

    # ogohlantirish uchburchagi (chap pastda)
    cv.draw_line(170, 820, 110, 720, GOLD, 5)
    cv.draw_line(170, 820, 230, 720, GOLD, 5)
    cv.draw_line(110, 720, 230, 720, GOLD, 5)
    cv.fill_rect(166, 745, 174, 795, GOLD)
    cv.filled_circle(170, 810, 5, GOLD)
    return cv.to_png()


def slide8():
    """Xulosa: USB + chip + elektron/kvant uchquni - mavzuni jamlovchi kompozitsiya."""
    cv = Canvas()
    _bg(cv)
    cv.radial_glow(500, 420, 520, CYAN, 0.18)
    _traces(cv, (24, 54, 86), step=100, t=1, dots=False)

    # markaziy chip
    cv.rounded_rect(360, 300, 640, 560, 20, (10, 24, 42))
    cv.rect_outline(360, 300, 640, 560, CYAN, 4)
    # chip ichidagi mitti panjara
    for yy in range(330, 540, 24):
        for xx in range(390, 620, 24):
            cv.fill_rect(xx, yy, xx + 16, yy + 16, (20, 46, 74))
    cv.circle_outline(500, 430, 70, CYAN, 4)
    # chip oyoqlari
    for i in range(8):
        cv.fill_rect(360 - 26, 320 + i * 30, 360, 334 + i * 30, GOLD)
        cv.fill_rect(640, 320 + i * 30, 640 + 26, 334 + i * 30, GOLD)
    # USB mini (yuqorida)
    cv.rounded_rect(420, 150, 580, 230, 14, PANEL)
    cv.fill_rect(360, 175, 420, 205, SILVER)
    # kvant uchqun (pastda)
    for ang in range(0, 360, 45):
        a = math.radians(ang)
        cv.draw_line(500, 680, 500 + 90 * math.cos(a), 680 + 90 * math.sin(a), VIOLET, 3)
    cv.filled_circle(500, 680, 18, CYAN)
    cv.radial_glow(500, 680, 70, CYAN, 0.5)
    # 0/1 oqim
    _draw_bits(cv, 150, 650, "10", CYAN, cell=14, gap=6)
    _draw_bits(cv, 760, 650, "01", GREEN, cell=14, gap=6)
    return cv.to_png()


def slide9():
    """Yakun: 'rahmat' temasidagi sodda grafika - chip + moviy nur."""
    cv = Canvas()
    _bg(cv, (12, 28, 52), (6, 14, 28))
    cv.radial_glow(500, 460, 560, CYAN, 0.28)
    _traces(cv, (22, 50, 82), step=120, t=1, dots=False)

    # markaziy chip (sodda, nur bilan)
    cv.rounded_rect(370, 330, 630, 590, 24, (10, 26, 46))
    cv.rect_outline(370, 330, 630, 590, CYAN, 4)
    for i in range(7):
        cv.fill_rect(344, 350 + i * 34, 370, 364 + i * 34, GOLD)
        cv.fill_rect(630, 350 + i * 34, 656, 364 + i * 34, GOLD)
    # yurak/pulse chizig'i (minnatdorchilik urg'usi)
    pts = [(150, 470), (300, 470), (350, 390), (420, 560),
           (480, 470), (850, 470)]
    # chip ustidan o'tmasin uchun ikki qism
    cv.draw_line(150, 470, 360, 470, CYAN, 4)
    cv.draw_line(640, 470, 850, 470, CYAN, 4)
    # markaziy halqa (teshik)
    cv.circle_outline(500, 460, 90, CYAN, 5)
    cv.filled_circle(500, 460, 12, CYAN)
    cv.radial_glow(500, 460, 50, CYAN, 0.6)
    # pastki aksent chiziqlari
    cv.fill_rect(0, 780, cv.w, 790, CYAN)
    cv.fill_rect(0, 796, cv.w, 802, BLUE)
    return cv.to_png()


_RENDERERS = [slide0, slide1, slide2, slide3, slide4,
              slide5, slide6, slide7, slide8, slide9]


def make_placeholder_png(idx):
    """idx (0..9) uchun mavzuga mos illyustratsiya PNG baytlarini qaytaradi.

    Barqaror kirish nuqtasi: build_pptx_fallback.py va generate_presentation.py
    shuni chaqiradi.
    """
    return _RENDERERS[idx % len(_RENDERERS)]()


if __name__ == "__main__":
    for i in range(10):
        with open("placeholder_%d.png" % i, "wb") as f:
            f.write(make_placeholder_png(i))
    print("10 ta mavzuga mos illyustratsiya yaratildi")
