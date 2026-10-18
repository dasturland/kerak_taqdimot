# -*- coding: utf-8 -*-
"""Taqdimot kontenti (matn, rasm URL'lari va tuzilma).

Ushbu modul bitta "haqiqat manbai" (single source of truth) sifatida ishlaydi.
Ham `generate_presentation.py` (python-pptx), ham `build_pptx_fallback.py`
(faqat standart kutubxona) shu kontentdan foydalanadi.

Barcha matn aynan foydalanuvchi bergan ko'rinishda (o'zbek tilida, apostroflar
bilan) saqlangan. Hech qanday o'zgartirish yoki qayta ifodalash yo'q.
"""

# Umumiy dizayn sozlamalari (kiberxavfsizlik / biometrika / fizika mavzusi)
FONT_NAME = "Times New Roman"

# Rang sxemasi (RGB) - to'q, texnologik ko'k/moviy urg'u bilan
COLOR_BG_DARK = (0x0B, 0x1E, 0x33)      # to'q ko'k fon
COLOR_PANEL = (0x12, 0x2A, 0x45)        # matn paneli uchun biroz ochiqroq
COLOR_ACCENT = (0x00, 0xC2, 0xD1)       # moviy / tech aksent (yorug'lik + tovush)
COLOR_ACCENT2 = (0x3B, 0x82, 0xF6)      # ko'k aksent
COLOR_TITLE = (0xFF, 0xFF, 0xFF)        # oq sarlavha
COLOR_BODY = (0xE2, 0xE8, 0xF0)         # ochiq kulrang matn

# Shrift o'lchamlari (pt). Minimal 12 pt talabiga rioya qilinadi.
SIZE_TITLE = 40        # titul sarlavhasi
SIZE_SUBTITLE = 24     # kichik sarlavha
SIZE_SLIDE_TITLE = 32  # oddiy slayd sarlavhasi
SIZE_BODY_HEAD = 18    # matn ichidagi kalit so'zlar (qalin)
SIZE_BODY = 16         # asosiy matn (12 dan katta)
SIZE_META = 14         # titul/yakuniy slayddagi ma'lumot (12 dan katta)

# Har bir slayd: type (chap/o'ng joylashuv va kontent turi), sarlavha,
# ixtiyoriy kichik sarlavha, matn qatorlari va Unsplash rasm URL'i.
SLIDES = [
    {
        "kind": "title",
        "image": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?auto=format&fit=crop&w=1000&q=80",
        "title": "Barmoq izi skanerlari qanday ishlaydi?",
        "subtitle": "Optik va ultratovush to'lqinlari yordamida biometrik axborotni tanib olish fizikasi.",
        "lines": [
            "Universitet: Jizzax davlat pedagogika universiteti tabiiy fanlar oliy maktabi",
            "Fan: Axborot tizimlari",
            "Guruh: 242_26",
            "Talaba: Toshbo\u02bbriyeva Marjona",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?auto=format&fit=crop&w=1000&q=80",
        "title": "Taqdimot rejasi",
        "lines": [
            "1-reja. Barmoq izi biologiyasi va uni raqamlashtirish (Minutiae nuqtalari).",
            "2-reja. Optik skanerlar fizikasi: Yorug'lik nurlari va 2D tasvir.",
            "3-reja. Ultratovushli (Ultrasonic) skanerlar: Ovoz to'lqinlari orqali 3D xarita.",
            "Xulosa. Biometrik xavfsizlikning zamonaviy axborot tizimlaridagi ahamiyati.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1614729939124-032f0b56c9ce?auto=format&fit=crop&w=1000&q=80",
        "title": "1. Biologiya va axborotning birlashuvi",
        "lines": [
            "Barmoq izimiz \u2014 bu inson tanasidagi tabiiy axborot tashuvchi noyob \"kod\" hisoblanadi.",
            "- Tepaliklar va chuqurliklar (Ridges and Valleys): Barmog'imiz terisi tekis emas. U qabariq chiziqlar (tepaliklar) va ular orasidagi bo'shliqlardan (chuqurliklar) iborat.",
            "- Minutiae nuqtalari: Skaner barmoqning umumiy rasmini emas, balki chiziqlar uzilgan, birlashgan yoki tugagan maxsus kesishuv nuqtalarini (Minutiae) izlaydi va ularni matematik algoritmga o'giradi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?auto=format&fit=crop&w=1000&q=80",
        "title": "Biometrik axborotni raqamlashtirish",
        "lines": [
            "Siz barmog'ingizni ekranga qo'yganingizda, axborot tizimi uni to'g'ridan-to'g'ri rasm sifatida saqlamaydi.",
            "- Shifrlash (Hashing): Topilgan kesishuv nuqtalari orasidagi masofa va burchaklar hisoblab chiqiladi. Bu geometrik ma'lumotlar murakkab shifr (Xesh kod - masalan, \"A7F90B...\") ko'rinishiga keltiriladi.",
            "- Tekshirish: Telefonni ochishda joriy barmoq izidan olingan shifr, xotiradagi (Secure Enclave'dagi) shifr bilan solishtiriladi. Ikkisi mos kelsa, telefon ochiladi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1516035069371-29a1b244cc32?auto=format&fit=crop&w=1000&q=80",
        "title": "2. Optik skanerlar: Yorug'lik fizikasi",
        "lines": [
            "Optik skanerlar qadimiyroq, ammo juda keng tarqalgan texnologiya bo'lib, ular asosan yorug'lik fizikasi yordamida ishlaydi.",
            "- Chaqnash (Illumination): Barmog'ingizni ekranga yoki tugmaga qo'yganingizda, ostki qismdan kuchli yorug'lik (LED) sochiladi. Bu yorug'lik barmoq izini to'liq yoritib beradi.",
            "- To'liq ichki qaytish: Barmoq terisi yorug'likni qaytaradi, bo'shliqlar (chuqurliklar) esa qorong'i bo'lib qoladi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1592434448555-520e5c9a752f?auto=format&fit=crop&w=1000&q=80",
        "title": "Optik tasvirni muhrlash (CMOS)",
        "lines": [
            "Yorug'lik barmoqdan qaytgach, xuddi kameradagi kabi maxsus sensorga (CMOS yoki CCD) borib uriladi.",
            "- 2D Tasvir: Sensor qaytgan yorug'lik asosida barmoq izining oq-qora 2 o'lchamli (2D) rasm xaritasini chizadi.",
            "- Kamchiligi: Optik skanerlar shunchaki rasm olgani uchun, ularni ba'zan yuqori sifatli fotonusxalar yoki yasalgan sun'iy barmoqlar yordamida aldash mumkin bo'ladi. Uning xavfsizlik darajasi o'rtacha.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1518932945647-7a1c969f8be2?auto=format&fit=crop&w=1000&q=80",
        "title": "3. Ultratovush (Ultrasonic) skanerlar",
        "lines": [
            "Zamonaviy flagman smartfonlarda (masalan, Samsung Galaxy S seriyasi) ishlatiladigan bu texnologiya yorug'likdan emas, balki tovush to'lqinlaridan foydalanadi.",
            "- Pyezoelektrik effekt: Skaner ostidagi datchik inson qulog'i eshitmaydigan yuqori chastotali (ultratovush) impuls tarqatadi.",
            "- Bu to'lqinlar displey orqali o'tib, to'g'ridan-to'g'ri barmog'ingizga borib uriladi va aks-sado (Echo) bo'lib orqaga qaytadi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=1000&q=80",
        "title": "Ovoz to'lqinlari orqali 3D xarita (Ekolokatsiya)",
        "lines": [
            "Ultratovush to'lqinlarining qaytish vaqti fizikaning ajoyib qonuniyatiga bo'ysunadi.",
            "- Qaytish farqi: To'lqin barmoqning bo'rtib chiqqan qismiga urilsa tezroq qaytadi, chuqurlik (havo bo'shlig'i) qismiga urilsa kechroq qaytadi.",
            "- 3D Model: Sensor ushbu vaqt farqlarini (millisoniyalarni) o'lchab, barmoq izining o'ta aniq 3D relyef xaritasini tuzadi. Bu usul qorong'ida ham, barmoq nam yoki kirlangan bo'lsa ham mukammal ishlaydi va uni rasmlar bilan aldab bo'lmaydi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&w=1000&q=80",
        "title": "Xulosa",
        "lines": [
            "Xulosa qilib aytganda, barmoq izi skanerlari inson biologiyasi va fizika qonunlarining IT bilan mukammal uyg'unligidir.",
            "Optik skanerlar yorug'lik nurlari (fotonlar) yordamida vizual xarita tuzsa, ultratovush skanerlar ovoz to'lqinlarining aks-sadosi yordamida o'ta aniq 3D model yaratadi. Bu texnologiyalar shaxsiy ma'lumotlarimiz va axborot tizimlarining xavfsizligini ta'minlashda eng ishonchli biometrik qulf bo'lib xizmat qilmoqda.",
        ],
    },
    {
        "kind": "closing",
        "image": "https://images.unsplash.com/photo-1517245386807-bb43f82c33c4?auto=format&fit=crop&w=1000&q=80",
        "title": "E'tiboringiz uchun rahmat!",
        "lines": [
            "Mavzu: Barmoq izi skanerlari (Optika va Ultratovush)",
            "Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - guruh)",
            "Jizzax davlat pedagogika universiteti - Axborot tizimlari fani.",
        ],
    },
]

assert len(SLIDES) == 10, "Taqdimot aynan 10 ta slayddan iborat bo'lishi shart"
