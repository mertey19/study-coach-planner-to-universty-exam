# -*- coding: utf-8 -*-
"""Application constants: day/hour mapping and defaults."""

# ---- Dosya ----
JSON_DOSYA = "program_kayitlari.json"
SAYFA_ADI = "Coaching Schedule"

# ---- Excel grid: gün -> sütun, saat -> satır ----
DAY_TO_COL = {
    "Monday": 2,
    "Tuesday": 3,
    "Wednesday": 4,
    "Thursday": 5,
    "Friday": 6,
    "Saturday": 7,
    "Sunday": 8,
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

# ---- Templates ----
HAZIR_METINLER = {
    "TYT Paragraph": "20 paragraph practice",
    "TYT Math": "TYT math problem solving",
    "AYT Physics": "AYT physics topic review",
    "General Review": "General review and previous questions",
    "Rest": "Break activity",
    "School time": "School time",
}
SABLON_SECIMLERI = ["(Select template)"] + list(HAZIR_METINLER.keys())
DEFAULT_9_TEXT = "20 paragraphs 20 problems"

# ---- Source list ----
KAYNAK_LISTESI = [
    "Uc Dort Bes",
    "Marka",
    "Limit",
    "3D",
    "Bilgi Sarmali",
    "Biyotigi",
    "Aydin",
    "Orbital",
    "4K",
    "Toprak",
    "Endemik",
]

# ---- Level (exam/grade) combo ----
SINAV_SECIMLERI = [
    "TYT", "AYT",
    "5th Grade", "6th Grade", "7th Grade", "8th Grade",
    "9th Grade", "10th Grade", "11th Grade", "12th Grade",
]

# ---- PDF font (Windows) ----
FONT_PATH = "C:/Windows/Fonts/arial.ttf"
