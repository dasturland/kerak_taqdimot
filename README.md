# Kompyuterlar nima uchun qiziydi va qanday soviydi? — Taqdimot generatori

Ushbu loyiha **10 ta slayddan** iborat professional o'zbekcha PowerPoint
taqdimotini avtomatik yaratadi. Mavzu: protsessor (CPU) ishlayotganda nega
issiqlik ajralib chiqadi va kulerlar yordamida sovutish fizikasi
(termodinamika va kompyuter arxitekturasi).

## Maqsad

`generate_presentation.py` skripti `python-pptx` kutubxonasi yordamida
`taqdimot.pptx` faylini quradi. Taqdimot qat'iy talablarga javob beradi:

- Aniq **10 ta slayd** (3 ta asosiy reja + Xulosa tuzilmasi).
- Butun matn sof **o'zbek tilida**, verbatim ko'chirilgan.
- Shrift faqat **Times New Roman**, eng kichik o'lcham **12pt** (sarlavhalar
  kattaroq).
- Qorong'u **kiber / kompyuter arxitekturasi / termodinamika** dizayni:
  issiqlik slaydlarida to'q sariq/qizil urg'u, sovutish slaydlarida ko'k/havorang.
- Har bir slaydga mavzuga mos rasm (Unsplash) yoki placeholder.

## Talablar (dependencies)

`requirements.txt` quyidagilarni o'z ichiga oladi:

- `python-pptx` — PPTX yaratish uchun asosiy kutubxona
- `requests` — Unsplash rasmlarini yuklab olish uchun
- `Pillow` — python-pptx tomonidan rasm bilan ishlashda talab qilinadi

## Ishga tushirish

```bash
pip install -r requirements.txt
python3 generate_presentation.py
```

Natijada joriy katalogda `taqdimot.pptx` fayli hosil bo'ladi. Skript har bir
slayd uchun rasm yuklanganmi yoki placeholder ishlatilganmi va qaysi
generatsiya yo'li tanlanganini chop etadi.

## Tekshirish (verification)

```bash
python3 verify_presentation.py
```

Tekshiruvchi `taqdimot.pptx` ni ochadi va quyidagilarni tasdiqlaydi:

- `ppt/presentation.xml` da aniq **10 ta** slayd mavjudligini,
- har bir slaydda shrift faqat **Times New Roman** ekanligini,
- hech qanday shrift o'lchami **12pt (sz=1200)** dan past emasligini.

Muvaffaqiyatda `PASS`, aks holda `FAIL` chop etiladi va nolga teng bo'lmagan
kod bilan chiqadi.

## Rasmlar: yuklab olish va placeholder zaxirasi

Har bir slayd uchun skript tegishli Unsplash URL ni yuklab olishga harakat
qiladi (`requests` mavjud bo'lsa u orqali, aks holda `urllib`). Yuklab olingan
rasmlar `images/` katalogiga saqlanadi (bu katalog `.gitignore` da).

Agar yuklab olish muvaffaqiyatsiz bo'lsa (tarmoq yopiq, 404, timeout va h.k.),
skript **qulamaydi** — o'rniga slayd mavzusiga mos bir rangli
**placeholder to'rtburchak** (nomli ramka) qo'yiladi.

## Tarmoqsiz muhit uchun avtomatik zaxira yo'li (stdlib fallback)

`generate_presentation.py` ikki yo'lni qo'llab-quvvatlaydi:

1. **Asosiy yo'l — `python-pptx`.** `pip install -r requirements.txt`
   bajarilgan oddiy mashinada taqdimot `python-pptx` orqali quriladi.
2. **Zaxira yo'l — faqat standart kutubxona.** Agar `import pptx`
   muvaffaqiyatsiz bo'lsa (masalan, tarmoqsiz muhitda o'rnatib bo'lmasa),
   skript **avtomatik ravishda** faqat Python standart kutubxonasi
   (`zipfile` + qo'lda yozilgan OOXML/XML) yordamida ekvivalent to'g'ri
   `.pptx` faylini yaratadi.

Ikkala yo'l ham barcha qat'iy talablarni (10 slayd, Times New Roman, >= 12pt,
qorong'u kiber/termodinamika dizayni) bir xil qondiradi. Skript oxirida qaysi
yo'l ishlatilganini chop etadi.

> Eslatma: tarmoqsiz muhitda barcha rasmlar placeholder bilan almashtiriladi
> va stdlib zaxira yo'li faollashadi — bu kutilgan va to'g'ri xatti-harakat.
