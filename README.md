# Barmoq izi skanerlari qanday ishlaydi? (PPTX taqdimot generatori)

Ushbu repozitoriy **"Barmoq izi skanerlari qanday ishlaydi? Optik va ultratovush
to'lqinlari fizikasi"** mavzusidagi 10 slaydli professional PowerPoint
taqdimotini avtomatik generatsiya qiluvchi skriptlardan iborat.

Tayyor fayl: **[`barmoq_izi_skanerlari.pptx`](barmoq_izi_skanerlari.pptx)** —
uni to'g'ridan-to'g'ri yuklab olib ochishingiz mumkin.

## Qat'iy talablarga muvofiqlik

| Talab | Holati |
|-------|--------|
| Aynan **10 ta slayd** | ✅ |
| Matn sof, professional, xatosiz **o'zbek tilida** (verbatim) | ✅ |
| Shrift faqat **"Times New Roman"** (har bir run'da aniq o'rnatilgan) | ✅ |
| Minimal shrift **12 pt** (title 40, subtitle 24, slayd sarlavhasi 32, matn 16, meta 14) | ✅ (eng kichik 13 pt) |
| Tuzilma **4 qism**: 3 ta asosiy reja + 1 ta xulosa | ✅ |
| Har bir slaydda mavzuga **aloqador vizual** (ichiga joylangan) | ✅ |
| Dizayn: **professional** kiberxavfsizlik / biometrika / fizika (to'q, texnologik ko'k-moviy) | ✅ |

## Professional dizayn va mavzuga aloqador rasmlar

Taqdimot **professional dizaynda** qayta ishlab chiqilgan va har bir slaydda
**mavzuga aloqador** vizual mavjud. Umumiy, umumlashtirilgan gradient
placeholderlar o'rniga endi har bir slaydning rasmi uning mazmunini bevosita
tasvirlaydi.

Rasmlar PowerPoint'ning o'z **vektor shakllari (DrawingML)** sifatida slayd
ichida chiziladi — ya'ni ular istalgan o'lchamda tiniq ko'rinadi va hech
qanday tashqi tarmoq yoki qo'shimcha kutubxona talab qilmaydi. Har bir
slaydning vizuali:

| Slayd | Vizual (mavzuga aloqador) |
|-------|---------------------------|
| 1. Titul | Stilizatsiya qilingan barmoq izi (looping ridge'lar) + skan chizig'i va skaner oynasi burchaklari |
| 2. Reja | Raqamlangan mundarija ikonkasi (4 qismli tuzilma) |
| 3. Biologiya | Ridge chiziqlari + ta'kidlangan **minutiae** nuqtalari (tugash/ayri) |
| 4. Raqamlashtirish | Barmoq izi → strelka → ikkilik (0/1) panjara va hash satri |
| 5. Optik | LED **yorug'lik nurlari** sirtga urilib qaytishi (reflection) |
| 6. CMOS | **Piksel panjarasi** (sensor matritsasi) 2D tasvir hosil qiladi |
| 7. Ultratovush | Emitterdan **konsentrik ultratovush impulslari** (ovoz yoylari) |
| 8. 3D xarita | **Echo** to'lqinlari + 3D relyef balandlik profili (ekolokatsiya) |
| 9. Xulosa | **Qalqon/qulf** — yorug'lik nuri va tovush yoyi birlashuvi (xavfsizlik) |
| 10. Rahmat | Barmoq izi + qulf nishoni (toza yakuniy belgi) |

Professional elementlar: diagonal gradientli to'q fon, izchil sarlavha
bandi va aksent chizig'i, bo'lim yorliqlari (**pill**'lar: 1-reja / 2-reja /
3-reja / Xulosa), o'ngdagi 5 nuqtali bo'lim indikatori hamda pastki
kolontitul (mavzu + slayd raqami).

> **Tarmoq cheklovi haqida eslatma:** ushbu muhitda tashqi tarmoq (Unsplash
> va boshqalar) bloklangan, shuning uchun rasmlar dasturiy ravishda vektor
> shaklida generatsiya qilinadi. Ochiq tarmoqli muhitda esa
> `generate_presentation.py` haqiqiy Unsplash fotosuratlarini yuklab oladi —
> ularning URL'lari `slides_content.py` ichida saqlangan.

## Fayllar tuzilishi

| Fayl | Vazifasi |
|------|----------|
| `slides_content.py` | Barcha slayd matni, rasm URL'lari, rang va shrift sozlamalari (yagona "haqiqat manbai"). |
| `generate_presentation.py` | **Asosiy skript** — `python-pptx` yordamida .pptx yaratadi (talab bo'yicha). |
| `build_pptx_fallback.py` | **Professional generator** — faqat Python standart kutubxonasi bilan (`python-pptx` yoki internet bo'lmaganda) .pptx tuzadi; vektor illyustratsiyalar bilan. |
| `illustrations.py` | Har bir slayd uchun **mavzuga aloqador vektor rasm** (DrawingML) generatori. |
| `placeholder.py` | (Eski) Unsplash rasmi yuklab olinmaganda ishlatilgan gradientli PNG placeholder — endi `generate_presentation.py` tomonidagi zaxira sifatida saqlanadi. |
| `verify_pptx.py` | Yaratilgan faylni tekshiradi (slayd soni, shrift, o'lcham, rasm). |
| `requirements.txt` | Bog'liqliklar (`python-pptx`). |
| `barmoq_izi_skanerlari.pptx` | Tayyor, generatsiya qilingan taqdimot. |

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

> **Eslatma:** Ushbu repozitoriyga biriktirilgan `barmoq_izi_skanerlari.pptx`
> cheklangan tarmoqli muhitda `build_pptx_fallback.py` orqali generatsiya
> qilingan. Undagi rasmlar — mavzuga **aloqador vektor illyustratsiyalar**
> (DrawingML), umumiy placeholderlar emas. Ochiq tarmoqli muhitda
> `generate_presentation.py` haqiqiy Unsplash fotosuratlarini yuklab olib
> joylaydi (URL'lar `slides_content.py` ichida).

## Tekshirish

```bash
python3 verify_pptx.py barmoq_izi_skanerlari.pptx
```

Chiqishda slaydlar soni (10), har bir run'da "Times New Roman", minimal
shrift o'lchami (>= 12 pt) va 10 ta joylangan rasm tasdiqlanadi.

## Taqdimot tuzilmasi (4 qism)

1. **1-reja:** Barmoq izi biologiyasi va uni raqamlashtirish (3- va 4-slaydlar).
2. **2-reja:** Optik skanerlar fizikasi — yorug'lik nurlari va 2D tasvir (5- va 6-slaydlar).
3. **3-reja:** Ultratovushli skanerlar — ovoz to'lqinlari orqali 3D xarita (7- va 8-slaydlar).
4. **Xulosa:** Biometrik xavfsizlikning ahamiyati (9-slayd) + minnatdorchilik slaydi (10-slayd).

Titul (1-slayd) va mundarija (2-slayd) ushbu qismlarni ochib beradi.
