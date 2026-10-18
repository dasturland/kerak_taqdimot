# -*- coding: utf-8 -*-
"""Taqdimot kontenti (matn, rasm URL'lari va tuzilma).

Ushbu modul bitta "haqiqat manbai" (single source of truth) sifatida ishlaydi.
Ham `generate_presentation.py` (python-pptx), ham `build_pptx_fallback.py`
(faqat standart kutubxona) shu kontentdan foydalanadi.

Barcha matn aynan foydalanuvchi bergan ko'rinishda (o'zbek tilida, apostroflar
bilan) saqlangan. Hech qanday o'zgartirish yoki qayta ifodalash yo'q.
"""

# Umumiy dizayn sozlamalari (kiber / mikroelektronika / xotira qurilmalari /
# kvant fizikasi mavzusi)
FONT_NAME = "Times New Roman"

# Rang sxemasi (RGB) - to'q, texnologik ko'k/moviy urg'u bilan
# (mikrosxema platasi, kremniy kristali va kvant tunnellanish mavzusiga mos)
COLOR_BG_DARK = (0x0B, 0x1E, 0x33)      # to'q ko'k fon
COLOR_PANEL = (0x12, 0x2A, 0x45)        # matn paneli uchun biroz ochiqroq
COLOR_ACCENT = (0x00, 0xC2, 0xD1)       # moviy / tech aksent (elektr zaryadi)
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
        "image": "https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&w=1000&q=80",
        "title": "Fleshkalar (USB) qanday ishlaydi?",
        "subtitle": "Hech qanday harakatlanuvchi qismlarsiz ma'lumotlarni elektr zaryadlari ko'rinishida saqlash fizikasi.",
        "lines": [
            "Universitet: Jizzax davlat pedagogika universiteti tabiiy fanlar oliy maktabi",
            "Fan: Axborot tizimlari",
            "Guruh: 242_26",
            "Talaba: Toshbo\u02bbriyeva Marjona",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1555949963-aa79dcee981c?auto=format&fit=crop&w=1000&q=80",
        "title": "Taqdimot rejasi",
        "lines": [
            "1-reja. NAND xotira arxitekturasi va harakatlanuvchi qismlarning yo'qligi.",
            "2-reja. Suzuvchi zulfin (Floating Gate) tranzistorlari fizikasi.",
            "3-reja. Axborotni yozish va o'chirish: Kvant tunnellanish effekti.",
            "Xulosa. Flesh xotiralarning zamonaviy axborot tizimlaridagi ahamiyati.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1517077304055-6e89abbf09b0?auto=format&fit=crop&w=1000&q=80",
        "title": "1. NAND Flash xotira arxitekturasi",
        "lines": [
            "Fleshkalar (USB) va SSD disklar qattiq disklar (HDD) kabi aylanuvchi magnit plastinalarga ega emas. Ularning ichida mutlaqo harakatlanuvchi qismlar yo'q.",
            "Solid-State:",
            "- Axborot faqatgina elektr toklari yordamida qattiq jismli (Solid-State) mikrosxemalarda saqlanadi.",
            "NAND darvozalari:",
            "- Fleshkaning ichidagi asosiy xotira chipi \"NAND\" deb ataluvchi milliardlab mitti tranzistorlardan tashkil topgan bo'lib, ular barcha ma'lumotlarni 0 va 1 ko'rinishida saqlaydi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1591405351990-4726e331f141?auto=format&fit=crop&w=1000&q=80",
        "title": "Mikro dunyodagi minglab rasmlar",
        "lines": [
            "Kichkina USB qurilmasiga minglab yuqori sifatli rasmlar qanday sig'adi?",
            "- Buning siri tranzistorlarning o'lchamidadir. Zamonaviy fleshkalardagi bitta tranzistor (ma'lumot saqlovchi katakcha) inson sochidan minglab marta ingichka bo'ladi.",
            "- Ularning milliardlab donasi bitta kremniy (Silicon) kristaliga zichlashtiriladi. Kompyuterdan kelgan har bir rasm va hujjat aynan shu mitti katakchalarga elektr zaryadi tariqasida taqsimlanadi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1631556097152-c39479bbff93?auto=format&fit=crop&w=1000&q=80",
        "title": "2. Suzuvchi zulfin (Floating Gate)",
        "lines": [
            "Fleshkadagi ma'lumot elektr toki uzilganda (kompyuterdan sug'urilganda) nega o'chib ketmaydi?",
            "Sababi oddiy tranzistorlardan farqli o'laroq, fleshkalar \"Suzuvchi zulfin\" (Floating Gate) deb nomlanuvchi maxsus tranzistorlardan foydalanadi.",
            "Izolyatsiya:",
            "- Bu zulfin elektr tokini o'tkazmaydigan o'ta yupqa oksid qatlami bilan har tomondan to'liq o'rab olingan (qamalgan).",
            "Qopqonga tushgan elektronlar:",
            "- Agar elektronlar bu zulfin ichiga kiritilsa, ular u yerdan chiqa olmay, yillab qamalib qoladi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?auto=format&fit=crop&w=1000&q=80",
        "title": "Zaryadni saqlash: 0 va 1 larni ifodalash",
        "lines": [
            "Zulfin ichiga elektronlar kiritilgani yoki kiritilmagani axborotni o'qish imkonini beradi.",
            "Mantiqiy \"0\":",
            "- Suzuvchi zulfinda elektronlar qamalgan bo'lsa, ular tranzistor orqali asosiy elektr tokining o'tishiga to'sqinlik qiladi. Kompyuter buni \"0\" deb o'qiydi.",
            "Mantiqiy \"1\":",
            "- Zulfinda elektronlar yo'q bo'lsa, elektr toki bemalol o'tadi va kompyuter buni \"1\" deb qabul qiladi. Shu tariqa fayllar fizik holatga ko'chadi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1635070041078-e363dbe005cb?auto=format&fit=crop&w=1000&q=80",
        "title": "3. Kvant tunnellanish effekti",
        "lines": [
            "Agar zulfin izolyator bilan to'liq o'ralgan bo'lsa, elektronlar uning ichiga qanday kiradi?",
            "Aynan shu yerda Kvant mexanikasi ishga tushadi! Kompyuter fleshkaga fayl yozayotganda tranzistorga kuchli kuchlanish (Volt) beradi.",
            "Fauler-Nordxeym (Fowler-Nordheim) tunnellanishi:",
            "- Kuchli elektr maydoni ta'sirida elektronlar \"kutilmaganda\" o'tib bo'lmas oksid to'sig'idan teshib o'tib, zulfin ichiga tushib qoladi. Bu mo'jizaviy kvant hodisasi axborotni fleshkaga yozish imkonini beradi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1517976487492-5750f3195933?auto=format&fit=crop&w=1000&q=80",
        "title": "Ma'lumotlarni o'chirish va fleshkaning yeyilishi",
        "lines": [
            "Ma'lumotni o'chirish uchun fleshka zulfinga teskari kuchlanish beradi va elektronlarni kvant tunnellanish orqali yana tashqariga so'rib oladi.",
            "Eskirish (Wear out):",
            "- Har safar elektronlar oksid qatlamidan kuch bilan teshib o'tganda, bu izolyatsiya qatlami sekin-asta yemiriladi.",
            "Hayot sikli:",
            "- Shu sababli USB fleshkalarning ma'lum bir yozish/o'chirish tsikli bor (odatda 10,000 dan 100,000 martagacha). Bu limit tugagach, u axborot saqlay olmay qoladi.",
        ],
    },
    {
        "kind": "content",
        "image": "https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&w=1000&q=80",
        "title": "Xulosa",
        "lines": [
            "Xulosa qilib aytganda, qo'limizdagi oddiy USB fleshka shunchaki plastik bo'lak emas. U ichida milliardlab mitti qopqonlarni saqlovchi kvant fizikasining mo'jizasidir.",
            "Kvant tunnellanish effekti kabi murakkab fizik qonuniyatlar insoniyatga ma'lumotlarni harakatlanuvchi mexanizmlarsiz, sof elektr zaryadlari ko'rinishida saqlash va bugungi tezkor axborot tizimlarini yaratish imkonini berdi.",
        ],
    },
    {
        "kind": "closing",
        "image": "https://images.unsplash.com/photo-1526628953301-3e589a6a8b74?auto=format&fit=crop&w=1000&q=80",
        "title": "E'tiboringiz uchun rahmat!",
        "lines": [
            "Mavzu: Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?",
            "Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - guruh)",
            "Jizzax davlat pedagogika universiteti - Axborot tizimlari fani.",
        ],
    },
]

assert len(SLIDES) == 10, "Taqdimot aynan 10 ta slayddan iborat bo'lishi shart"
