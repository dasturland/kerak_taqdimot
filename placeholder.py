# -*- coding: utf-8 -*-
"""Mavzuga mos (mikroelektronika/kvant fizikasi) placeholder PNG yaratish.

Faqat standart kutubxona (zlib, struct) ishlatiladi. Bu modul Unsplash
rasmlari yuklab olinmagan holatda zaxira sifatida chiroyli gradientli,
to'lqinli rasmlar beradi, shunda .pptx baribir yaroqli bo'ladi.
"""
import math
import struct
import zlib

# Har bir slayd uchun mavzuga mos rang palitralari (ko'k/moviy texnologik -
# mikrosxema, kremniy kristali, elektr zaryadi va kvant tunnellanish mavzusi)
PALETTES = [
    ((11, 30, 51), (0, 194, 209)),    # 1 titul (USB fleshka)
    ((9, 20, 40), (59, 130, 246)),    # 2 reja
    ((20, 14, 40), (0, 194, 209)),    # 3 NAND arxitekturasi
    ((5, 25, 35), (34, 197, 94)),     # 4 kremniy kristali / mikro dunyo
    ((40, 25, 10), (251, 191, 36)),   # 5 suzuvchi zulfin (Floating Gate)
    ((15, 20, 35), (99, 179, 237)),   # 6 zaryad: 0 va 1
    ((25, 10, 40), (167, 139, 250)),  # 7 kvant tunnellanish effekti
    ((8, 28, 44), (0, 212, 255)),     # 8 o'chirish / yeyilish
    ((11, 30, 51), (0, 194, 209)),    # 9 xulosa
    ((20, 10, 30), (59, 130, 246)),   # 10 yakun
]

W, H = 1000, 1000


def _lerp(a, b, t):
    return int(a + (b - a) * t)


def make_placeholder_png(idx):
    """idx (0..9) uchun gradient + to'lqin naqshli PNG baytlarini qaytaradi."""
    c1, c2 = PALETTES[idx % len(PALETTES)]
    rows = bytearray()
    for y in range(H):
        rows.append(0)  # filter type 0 (None) har qator oldidan
        ty = y / (H - 1)
        for x in range(W):
            tx = x / (W - 1)
            # diagonal gradient
            t = (tx + ty) / 2.0
            r = _lerp(c1[0], c2[0], t)
            g = _lerp(c1[1], c2[1], t)
            b = _lerp(c1[2], c2[2], t)
            # yorug'lik/tovush to'lqinlari naqshi
            wave = math.sin((x * 0.03) + (y * 0.015) + idx) * math.cos(y * 0.02)
            boost = int(28 * wave)
            r = max(0, min(255, r + boost))
            g = max(0, min(255, g + boost))
            b = max(0, min(255, b + boost))
            rows += bytes((r, g, b))

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", W, H, 8, 2, 0, 0, 0)  # 8-bit, truecolor RGB
    idat = zlib.compress(bytes(rows), 9)
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


if __name__ == "__main__":
    for i in range(10):
        with open("placeholder_%d.png" % i, "wb") as f:
            f.write(make_placeholder_png(i))
    print("10 ta placeholder yaratildi")
