# -*- coding: utf-8 -*-
"""Uygulama sabitleri: gün/saat eşlemesi, şablonlar, yayın listesi, dosya adları."""

# ---- Dosya ----
JSON_DOSYA = "program_kayitlari.json"
SAYFA_ADI = "Koçluk Çizelgesi"

# ---- Excel grid: gün -> sütun, saat -> satır ----
DAY_TO_COL = {
    "Pazartesi": 2,
    "Salı": 3,
    "Çarşamba": 4,
    "Perşembe": 5,
    "Cuma": 6,
    "Cumartesi": 7,
    "Pazar": 8,
}
COL_TO_DAY = {v: k for k, v in DAY_TO_COL.items()}

TIME_TO_ROW = {
    "00:00": 3,
    "01:00": 4,
    "02:00": 5,
    "03:00": 6,
    "04:00": 7,
    "05:00": 8,
    "06:00": 9,
    "07:00": 10,
    "08:00": 11,
    "09:00": 12,
    "10:00": 13,
    "11:00": 14,
    "12:00": 15,
    "13:00": 16,
    "14:00": 17,
    "15:00": 18,
    "16:00": 19,
    "17:00": 20,
    "18:00": 21,
    "19:00": 22,
    "20:00": 23,
    "21:00": 24,
    "22:00": 25,
    "23:00": 26,
}
ROW_TO_TIME = {v: k for k, v in TIME_TO_ROW.items()}

GUNLER = list(DAY_TO_COL.keys())
SAATLER = list(TIME_TO_ROW.keys())

# ---- Şablonlar ----
HAZIR_METINLER = {
    "TYT Paragraf": "20 paragraf çalışması",
    "TYT Matematik": "TYT Mat Soru Çözümü",
    "AYT Fizik": "AYT Fizik Konu Tekrarı",
    "Genel Tekrar": "Genel tekrar & önceki sorular",
    "dinlenme": "Boş aktivite",
    "Okul zamanı": "Okul zamanı",
}
SABLON_SECIMLERI = ["(Şablon seç)"] + list(HAZIR_METINLER.keys())
DEFAULT_9_TEXT = "20 paragraf 20 problem"

# ---- Yayın (kaynak) listesi ----
KAYNAK_LISTESI = [
    "Üç Dört Beş Yayınları",
    "Marka Yayınları",
    "Limit Yayınları",
    "3D Yayınları",
    "Bilgi Sarmalı Yayınları",
    "Biyotiği Yayınları",
    "Aydın Yayınları",
    "Orbital Yayınları",
    "4K Yayınları",
    "Toprak Yayınları",
    "Endemik Yayınları",
]

# ---- Seviye (sınav / sınıf) combo ----
SINAV_SECIMLERI = [
    "TYT", "AYT",
    "5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf",
    "9. Sınıf", "10. Sınıf", "11. Sınıf", "12. Sınıf",
]

# ---- PDF font (Windows) ----
FONT_PATH = "C:/Windows/Fonts/arial.ttf"
