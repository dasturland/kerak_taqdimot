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
| Minimal shrift **12 pt** (title 40, subtitle 24, slayd sarlavhasi 32, matn 16, meta/futer 14) | ✅ (eng kichik 14 pt) |
| Tuzilma **4 qism**: 3 ta asosiy reja + 1 ta xulosa | ✅ |
| Har bir slaydda mavzuga mos **rasm** (ichiga joylangan, cho'zilishsiz) | ✅ |
| Dizayn: **professional, korporativ** (sarlavha bandi, aksent chiziq, matn paneli, futer) | ✅ |
| Mavzu dizayni: **kiber / mikroelektronika / xotira qurilmalari / kvant fizikasi** (to'q, texnologik ko'k-moviy) | ✅ |

## Fayllar tuzilishi

| Fayl | Vazifasi |
|------|----------|
| `slides_content.py` | Barcha slayd matni, rasm URL'lari, rang va shrift sozlamalari (yagona "haqiqat manbai"). |
| `generate_presentation.py` | **Asosiy skript** — `python-pptx` yordamida .pptx yaratadi (talab bo'yicha). |
| `build_pptx_fallback.py` | Zaxira generator — faqat Python standart kutubxonasi bilan (`python-pptx` yoki internet bo'lmaganda) aynan bir xil .pptx tuzadi. |
| `placeholder.py` | Faqat standart kutubxona (`zlib`, `struct`, `math`) bilan mavzuga **aloqador protsessual illyustratsiyalar** chizadi (USB konnektori, NAND katak panjarasi, kremniy kristali, suzuvchi zulfin kesimi, kvant tunnellanish va h.k.) - PIL ham, internet ham kerak emas. |
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
> qilingan. Shu sababli undagi rasmlar yuklab olinmagan, balki
> `placeholder.py` tomonidan **faqat standart kutubxona bilan chizilgan,
> mavzuga aloqador protsessual illyustratsiyalar**dir (har bir slayd uchun
> alohida: USB tayoqchasi, NAND katak panjarasi, kremniy kristali, suzuvchi
> zulfin kesimi, elektronlarning qamalishi, kvant tunnellanish, o'chirish va
> yeyilish, xulosa va yakun). Bular tekis gradientli oddiy fon emas, balki
> har bir slayd mazmuniga mos chizmalardir. Internetli muhitda yuqoridagi
> skriptlardan birini qayta ishga tushirsangiz, haqiqiy Unsplash rasmlari
> yuklab olinib joylanadi (protsessual illyustratsiya faqat zaxira sifatida
> ishlatiladi).

## Professional dizayn

Slaydlar korporativ, tartibli ko'rinishga ega:

- **Sarlavha bandi**: har bir mazmun slaydida oq, qalin sarlavha, chapda
  vertikal aksent chizig'i va sarlavha ostida nozik moviy aksent chizig'i.
- **Matn paneli**: chapda ochiqroq ko'k panel, keng ichki chekkalar bilan
  (matn panel chetiga yopishmaydi).
- **Rasmlar cho'zilmaydi**: kvadrat (1000x1000) illyustratsiya kvadrat ramka
  ichiga aspekt nisbatini saqlagan holda, markazlashtirib joylanadi
  (letterbox), shu sababli hech qanday buzilish bo'lmaydi.
- **Vizual ierarxiya**: sarlavha qatorlari (chiziqcha bilan boshlanmaydiganlar)
  qalin va aksent rangda; chiziqcha (`-`) bilan boshlanadigan qatorlar asosiy
  matn rangida va kichik chekinish bilan beriladi.
- **Nozik futer**: mazmun slaydlari pastida mavzu nomi (chapda) va slayd
  raqami `N / 10` (o'ngda), Times New Roman, 14 pt.
- **Titul va yakuniy slayd** ham moslashtirilgan: titulda rasm paneli,
  yakuniy slaydda markaziy aksent ramkali panel.

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
