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
- Har bir slaydga mavzuga mos rasm (Unsplash) yoki chizilgan motivli placeholder.

## Professional dizayn elementlari

Taqdimot shunchaki tekis bitta urg'u chizig'idan iborat emas — har bir
kontent slaydida:

- Yuqorida **urg'u rangli sarlavha bandi** (header band) va sarlavha ostida
  ingichka **urg'u chizig'i** (underline).
- Pastda **footer tasma**: chapda deck yorlig'i
  ("Kompyuter qizishi va sovutish fizikasi"), o'ngda slayd raqami
  (masalan `3 / 10`).
- Ro'yxat elementlari oldida rangli **kvadrat marker**.
- Rasm uchun **urg'u ramkali panel**: haqiqiy rasm mavjud bo'lsa panelni
  to'liq qoplaydi (crop-to-fit, cho'zilmaydi), aks holda chizilgan
  **chip yoki radiator motivli** (konsentrik to'rtburchaklar / qanotchalar)
  professional placeholder karta va o'zbekcha izoh chiqadi.
- Titul (1-slayd) va rahmat (10-slayd) alohida **markazlashtirilgan hero**
  ko'rinishida.

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
- hech qanday shrift o'lchami **12pt (sz=1200)** dan past emasligini,
- deck haqiqiy matn yugurdaklariga (`<a:t>`) ega ekanligini (faqat
  placeholderlardan iborat, matnsiz deck XATO beradi),
- bir nechta verbatim tayanch satr mavjudligini: U+02BB apostrofli
  **"Toshbo\u02bbriyeva"**, uchta reja belgisi (**1-reja. / 2-reja. /
  3-reja.**) va **"Xulosa"**.

Muvaffaqiyatda `PASS`, aks holda `FAIL` chop etiladi va nolga teng bo'lmagan
kod bilan chiqadi.

## Rasmlar: yuklab olish va placeholder zaxirasi

Har bir slayd uchun skript tegishli Unsplash URL ni yuklab olishga harakat
qiladi (`requests` mavjud bo'lsa u orqali, aks holda `urllib`). Yuklab olingan
rasmlar `images/` katalogiga saqlanadi (bu katalog `.gitignore` da).

Rasm topilsa, u rasm panelini to'liq qoplaydigan qilib (crop-to-fit)
joylashtiriladi — rasm cho'zilmaydi, balki panelga sig'maydigan chetlari
qirqiladi.

Agar yuklab olish muvaffaqiyatsiz bo'lsa (tarmoq yopiq, 404, timeout va h.k.),
skript **qulamaydi** — o'rniga slayd mavzusiga mos, chizilgan chip/radiator
motivli **professional placeholder karta** (urg'u ramkali) qo'yiladi.
Placeholder izohi ham **o'zbek tilida** ("Rasm: ...") bo'ladi, shunda
tarmoqsiz qurishda butun taqdimot sof o'zbekcha qoladi.

Yuklab olingan rasm turi (JPEG, PNG, WebP, GIF, BMP) baytlar sehrli belgisi
(magic bytes), so'ng HTTP `Content-Type`, oxirida URL kengaytmasi bo'yicha
aniqlanadi va fayl mos kengaytma bilan to'g'ri `Content-Type` ostida
joylashtiriladi (Unsplash `auto=format` orqali PNG/WebP qaytarishi mumkin).

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

> Smoke-test eslatmasi: `python-pptx` asosiy yo'li tarmoqsiz sandboxda ishga
> tushirib bo'lmaydi (kutubxona o'rnatilmaydi), shuning uchun uni tarmoqli
> mashinada bir marta `pip install -r requirements.txt && python3
> generate_presentation.py` orqali tekshirib ko'rish tavsiya etiladi.

> Eslatma: tarmoqsiz muhitda barcha rasmlar placeholder bilan almashtiriladi
> va stdlib zaxira yo'li faollashadi — bu kutilgan va to'g'ri xatti-harakat.
