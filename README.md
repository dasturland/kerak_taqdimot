# Fleshkalar (USB) qanday ishlaydi? (PPTX taqdimot generatori)

Ushbu repozitoriy **"Fleshkalar (USB) qanday qilib ma'lumot saqlaydi? Elektr
zaryadlari ko'rinishida axborot saqlash fizikasi"** mavzusidagi 10 slaydli
professional PowerPoint taqdimotini avtomatik generatsiya qiluvchi
skriptlardan iborat.

Tayyor fayl: **[`fleshka_usb_taqdimot.pptx`](fleshka_usb_taqdimot.pptx)** —
uni to'g'ridan-to'g'ri yuklab olib ochishingiz mumkin.

## Qat'iy talablarga muvofiqlik

| Talab | Holati |
|-------|--------|
| Aynan **10 ta slayd** | ✅ |
| Matn sof, professional, xatosiz **o'zbek tilida** (verbatim) | ✅ |
| Shrift faqat **"Times New Roman"** (har bir run'da aniq o'rnatilgan) | ✅ |
| Minimal shrift **12 pt** (title 40, subtitle 24, slayd sarlavhasi 32, matn 16, meta 14) | ✅ (eng kichik 14 pt) |
| Tuzilma **4 qism**: 3 ta asosiy reja + 1 ta xulosa | ✅ |
| Har bir slaydda mavzuga mos **Unsplash rasmi** (ichiga joylangan) | ✅ |
| Dizayn: **kiber / mikroelektronika / xotira qurilmalari / kvant fizikasi** (to'q, texnologik ko'k-moviy) | ✅ |

## Fayllar tuzilishi

| Fayl | Vazifasi |
|------|----------|
| `slides_content.py` | Barcha slayd matni, rasm URL'lari, rang va shrift sozlamalari (yagona "haqiqat manbai"). |
| `generate_presentation.py` | **Asosiy skript** — `python-pptx` yordamida .pptx yaratadi (talab bo'yicha). |
| `build_pptx_fallback.py` | Zaxira generator — faqat Python standart kutubxonasi bilan (`python-pptx` yoki internet bo'lmaganda) aynan bir xil .pptx tuzadi. |
| `placeholder.py` | Unsplash rasmi yuklab olinmaganda mavzuga mos gradientli PNG placeholder yaratadi. |
| `verify_pptx.py` | Yaratilgan faylni tekshiradi (slayd soni, shrift, o'lcham, rasm). |
| `requirements.txt` | Bog'liqliklar (`python-pptx`). |
| `fleshka_usb_taqdimot.pptx` | Tayyor, generatsiya qilingan taqdimot. |

## Ishga tushirish

### 1-usul: python-pptx bilan (tavsiya etiladi)

```bash
pip install -r requirements.txt
python3 generate_presentation.py
# yoki boshqa nom bilan:
python3 generate_presentation.py mening_taqdimotim.pptx
```

Skript har bir slayd uchun Unsplash rasmini yuklab olib, slaydga **joylaydi**
(havola emas, balki rasmning o'zi). Agar rasm yuklab olinmasa (internet yo'q
yoki URL ishlamasa), mavzuga mos rangli placeholder ishlatiladi, shunda fayl
baribir yaroqli 10 slaydli .pptx bo'lib qoladi.

### 2-usul: internet/python-pptx bo'lmagan muhit uchun

```bash
python3 build_pptx_fallback.py
```

Bu variant hech qanday tashqi kutubxonaga muhtoj emas (faqat `zipfile`,
`zlib` kabi standart modullar). `.pptx` fayli OOXML ZIP to'plami bo'lgani
uchun to'g'ridan-to'g'ri tuziladi.

> **Eslatma:** Ushbu repozitoriyga biriktirilgan `fleshka_usb_taqdimot.pptx`
> cheklangan tarmoqli muhitda `build_pptx_fallback.py` orqali generatsiya
> qilingan, shuning uchun undagi rasmlar mavzuga mos **placeholder**lardir.
> Internetli muhitda yuqoridagi skriptlardan birini qayta ishga tushirsangiz,
> haqiqiy Unsplash rasmlari yuklab olinib joylanadi.

## Tekshirish

```bash
python3 verify_pptx.py fleshka_usb_taqdimot.pptx
```

Chiqishda slaydlar soni (10), har bir run'da "Times New Roman", minimal
shrift o'lchami (>= 12 pt) va 10 ta joylangan rasm tasdiqlanadi.

## Taqdimot tuzilmasi (4 qism)

1. **1-reja:** NAND xotira arxitekturasi va harakatlanuvchi qismlarning yo'qligi (3- va 4-slaydlar).
2. **2-reja:** Suzuvchi zulfin (Floating Gate) tranzistorlari fizikasi (5- va 6-slaydlar).
3. **3-reja:** Axborotni yozish va o'chirish — Kvant tunnellanish effekti (7- va 8-slaydlar).
4. **Xulosa:** Flesh xotiralarning zamonaviy axborot tizimlaridagi ahamiyati (9-slayd) + minnatdorchilik slaydi (10-slayd).

Titul (1-slayd) va mundarija (2-slayd) ushbu qismlarni ochib beradi.
