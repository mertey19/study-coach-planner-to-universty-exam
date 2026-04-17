from __future__ import annotations

import copy
import os
import sys
import ctypes
import re
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from typing import Optional

# Yüksek DPI (Windows)
if sys.platform.startswith("win"):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# Proje kökü (PyInstaller .exe: veriler exe'nin bulunduğu klasöre yazılır)
def _app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_SCRIPT_DIR = _app_base_dir()
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from constants import (
    DEFAULT_9_TEXT,
    GUNLER as gunler,
    HAZIR_METINLER,
    KAYNAK_LISTESI,
    SABLON_SECIMLERI,
    SAATLER as saatler,
    SAYFA_ADI,
)
from storage import JsonStorage
from services import ExcelService, PdfService
from config import get_pdf_font_path, get_sheet_name, get_theme, load_settings, save_settings

# Eski sabit isimleri (bu dosyada hâlâ kullanılan yerler için)
from constants import COL_TO_DAY, DAY_TO_COL, ROW_TO_TIME, TIME_TO_ROW

# Panel / tema renkleri (birkaç örnek)
THEMES = {
    "Gri (varsayılan)": {
        "bg": "#f0f0f0", "fg": "#1a1a2e", "entry_bg": "#ffffff",
        "btn_bg": "#e8e8e8", "btn_active": "#d0d0d0", "accent": "#2563eb", "label_muted": "#555",
    },
    "Açık mavi": {
        "bg": "#e8f4fc", "fg": "#1a365d", "entry_bg": "#ffffff",
        "btn_bg": "#bae0f7", "btn_active": "#7cc4ea", "accent": "#2563eb", "label_muted": "#4a5568",
    },
    "Yeşil": {
        "bg": "#e8f5e9", "fg": "#1b2e1b", "entry_bg": "#ffffff",
        "btn_bg": "#c8e6c9", "btn_active": "#a5d6a7", "accent": "#2e7d32", "label_muted": "#555",
    },
    "Sıcak bej": {
        "bg": "#faf5ef", "fg": "#3e2723", "entry_bg": "#ffffff",
        "btn_bg": "#efe0d0", "btn_active": "#ddccb8", "accent": "#8d6e63", "label_muted": "#5d4e37",
    },
    "Koyu": {
        "bg": "#2d2d2d", "fg": "#e0e0e0", "entry_bg": "#404040",
        "btn_bg": "#505050", "btn_active": "#606060", "accent": "#64b5f6", "label_muted": "#b0b0b0",
    },
    "Lacivert": {
        "bg": "#e8ecf4", "fg": "#1a237e", "entry_bg": "#ffffff",
        "btn_bg": "#c5cae9", "btn_active": "#9fa8da", "accent": "#3949ab", "label_muted": "#3f51b5",
    },
    "Mor": {
        "bg": "#f3e5f5", "fg": "#4a148c", "entry_bg": "#ffffff",
        "btn_bg": "#e1bee7", "btn_active": "#ce93d8", "accent": "#7b1fa2", "label_muted": "#6a1b9a",
    },
}
THEME_KEYS = list(THEMES.keys())

# ---- TYT / AYT DERS & KONU VERİSİ ----
KONU_VERISI = {
    "TYT": {
        "Türkçe": [
            "Sözcükte Anlam",
            "Deyimler ve Atasözleri",
            "Cümlede Anlam",
            "Paragrafta Anlam",
            "Paragrafta Anlatım Teknikleri",
            "Paragrafta Yapı ve Akış",
            "Ses Bilgisi",
            "Yazım Kuralları",
            "Noktalama İşaretleri",
            "Sözcük Türleri",
            "Fiiller",
            "Fiilimsi",
            "Cümle Türleri",
            "Cümle Ögeleri",
            "Anlatım Bozuklukları"
        ],
        "Matematik": [
            "Temel Kavramlar",
            "Sayı Basamakları",
            "Bölme ve Bölünebilme",
            "EBOB - EKOK",
            "Rasyonel Sayılar",
            "Basit Eşitsizlikler",
            "Mutlak Değer",
            "Üslü Sayılar",
            "Köklü Sayılar",
            "Çarpanlara Ayırma",
            "Oran Orantı",
            "Denklem Çözme",
            "Problemler - Sayı Problemleri",
            "Problemler - Kesir Problemleri",
            "Problemler - Yaş Problemleri",
            "Problemler - Hareket Hız Problemleri",
            "Problemler - İşçi Emek Problemleri",
            "Problemler - Yüzde Problemleri",
            "Problemler - Kar Zarar Problemleri",
            "Problemler - Karışım Problemleri",
            "Problemler - Grafik Problemleri",
            "Problemler - Rutin Olmayan Problemler",
            "Kümeler - Kartezyen Çarpım",
            "Mantık",
            "Fonksiyonlar",
            "Polinomlar",
            "2. Dereceden Denklemler",
            "Permütasyon ve Kombinasyon",
            "Olasılık",
            "Veri - İstatistik"
        ],
        "Geometri": [
            "Temel Kavramlar ve Doğruda Açılar",
            "Üçgenler",
            "Açıortay ve Kenarortay",
            "Üçgende Eşlik ve Benzerlik",
            "Üçgende Alanlar",
            "Çokgenler",
            "Dörtgenler (paralelkenar, kare, dikdörtgen, yamuk)",
            "Çember ve Daire",
            "Katı Cisimler",
            "Analitik Geometri (Doğru)"
        ],
        "Fizik": [
            "Fizik Bilimine Giriş",
            "Büyükler ve Birimler",
            "Hareket ve Kuvvet",
            "İş, Güç ve Enerji",
            "Isı ve Sıcaklık",
            "Madde ve Isı",
            "Basınç",
            "Kaldırma Kuvveti",
            "Dalgalar",
            "Optik",
            "Elektrik ve Manyetizma"
        ],
        "Kimya": [
            "Kimya Bilimi",
            "Atomun Yapısı",
            "Periyodik Sistem",
            "Kimyasal Türler Arası Etkileşimler",
            "Maddenin Halleri",
            "Kimyanın Temel Kanunları",
            "Mol Kavramı ve Kimyasal Hesaplamalar",
            "Karışımlar",
            "Asit, Baz ve Tuzlar",
            "Kimya Her Yerde"
        ],
        "Biyoloji": [
            "Biyoloji Bilimi",
            "Canlıların Ortak Özellikleri",
            "Canlıların Temel Bileşenleri",
            "Hücre ve Yapısı",
            "Hücre Zarından Madde Geçişi",
            "Canlılarda Enerji Dönüşümleri",
            "Fotosentez",
            "Solunum",
            "Kalıtımın Temel İlkeleri",
            "Ekosistem Ekoloji",
            "İnsan Fizyolojisine Giriş"
        ],
        "Tarih": [
            "Tarih ve Zaman",
            "İnsanlığın İlk Dönemleri",
            "İlk ve Orta Çağlarda Türk Dünyası",
            "İslam Medeniyetinin Doğuşu",
            "İlk Türk İslam Devletleri",
            "Türkiye Tarihi (Beylikten Devlete)",
            "Osmanlı Devleti'nin Kuruluşu",
            "Osmanlı Yükselme Dönemi",
            "Klasik Çağda Osmanlı Toplum ve Ekonomi",
            "18. ve 19. yüzyılda Osmanlı",
            "I. Dünya Savaşı ve Sonuçları",
            "Milli Mücadele Dönemi",
            "İnkılaplar ve Atatürkçülük",
            "Çağdaş Türk ve Dünya Tarihi'ne Giriş"
        ],
        "Coğrafya": [
            "Doğa ve İnsan",
            "Dünyanın Şekli ve Hareketleri",
            "Coğrafi Konum",
            "Harita Bilgisi",
            "İklim Bilgisi",
            "İç ve Dış Kuvvetler",
            "Nüfus ve Yerleşme",
            "Türkiye'nin Yer Şekilleri",
            "Türkiye'nin İklimi ve Bitki Örtüsü",
            "Türkiye'de Nüfus ve Yerleşme",
            "Ekonomik Faaliyetler",
            "Bölgeler ve Ülkeler"
        ],
        "Felsefe": [
            "Felsefeye Giriş",
            "Bilgi Felsefesi",
            "Varlık Felsefesi",
            "Ahlak Felsefesi",
            "Siyaset Felsefesi",
            "Sanat Felsefesi",
            "Din Felsefesi",
            "Bilim Felsefesi"
        ],
        "Din Kültürü": [
            "Bilgi ve İnanç",
            "İslam ve İbadet",
            "Kur’an ve Yorumu",
            "Hz. Muhammed'in Hayatı ve Örnekliği",
            "Ahlak ve Değerler",
            "Dinler ve Mezhepler",
            "Güncel Dini Meseleler"
        ]
    },
    "AYT": {
        "Matematik": [
            "Sayı Kümeleri ve Sayı Sistemleri",
            "Fonksiyonlar",
            "Polinomlar",
            "Karmaşık Sayılar",
            "2. Dereceden Denklemler",
            "Eşitsizlikler",
            "Kümeler ve Fonksiyonlarda Uygulamalar",
            "Permütasyon",
            "Kombinasyon",
            "Binom",
            "Olasılık",
            "Trigonometri",
            "Analitik Geometri (Doğru ve Daire)",
            "Diziler",
            "Logaritma",
            "Limit",
            "Süreklilik",
            "Türev ve Uygulamaları",
            "İntegral ve Uygulamaları"
        ],
        "Geometri": [
            "Temel Kavramlar ve Açı",
            "Üçgenler ve Özellikleri",
            "Çokgenler ve Dörtgenler",
            "Çember ve Daire",
            "Analitik Geometri",
            "Katı Cisimler",
            "Çemberde Açı ve Uzunluk",
            "Dönüşümler ve Geometrik Yer"
        ],
        "Fizik": [
            "Vektörler",
            "Kuvvet, Tork ve Denge",
            "Kütle Merkezi",
            "Basit Makineler",
            "Hareket",
            "Newton'un Hareket Yasaları",
            "İş, Güç ve Enerji",
            "Atışlar",
            "İtme ve Momentum",
            "Elektrik Alan ve Potansiyel",
            "Paralel Levhalar ve Sığaçlar",
            "Manyetik Alan ve Manyetik Kuvvet",
            "Elektromanyetik İndüksiyon",
            "Alternatif Akım",
            "Dalgalar ve Optik",
            "Çembersel Hareket",
            "Dönme, Yuvarlanma, Açısal Momentum",
            "Kütle Çekim ve Kepler Yasaları",
            "Modern Fizik ve Uygulamaları"
        ],
        "Kimya": [
            "Modern Atom Teorisi",
            "Kimyasal Hesaplamalar",
            "Gazlar",
            "Sıvı Çözeltiler ve Çözünürlük",
            "Kimyasal Tepkimelerde Enerji",
            "Tepkimelerde Hız",
            "Kimyasal Denge",
            "Asit - Baz Dengeleri",
            "Çözelti Dengeleri",
            "Redoks Tepkimeleri ve Elektrokimya",
            "Karbon Kimyasına Giriş",
            "Organik Bileşikler",
            "Polimerler ve Günlük Hayatta Kimya"
        ],
        "Biyoloji": [
            "Sinir Sistemi",
            "Endokrin Sistem",
            "Duyu Organları",
            "Destek ve Hareket Sistemi",
            "Sindirim Sistemi",
            "Dolaşım ve Bağışıklık",
            "Solunum Sistemi",
            "Boşaltım Sistemi",
            "Üreme Sistemi ve Embriyo Gelişimi",
            "Genetik Bilgi ve DNA",
            "Protein Sentezi",
            "Genetik Çeşitlilik ve Evrim",
            "Ekosistem Ekoloji ve Enerji Akışı",
            "Bitki Biyolojisi ve Fotosentez",
            "Komünite Ekolojisi ve Biyomlar"
        ],
        "Edebiyat": [
            "Güzel Sanatlar ve Edebiyat",
            "Söz Sanatları",
            "Edebi Metin Türleri",
            "İslamiyet Öncesi Türk Edebiyatı",
            "Geçiş Dönemi Eserleri",
            "Divan Edebiyatı",
            "Halk Edebiyatı",
            "Tanzimat Edebiyatı",
            "Servet-i Fünun Edebiyatı",
            "Fecriati Edebiyatı",
            "Milli Edebiyat",
            "Cumhuriyet Dönemi Türk Edebiyatı (Şiir)",
            "Cumhuriyet Dönemi Türk Edebiyatı (Roman, Hikaye)",
            "Cumhuriyet Dönemi Tiyatro",
            "Dünya Edebiyatından Seçmeler"
        ],
        "Tarih-1": [
            "Tarih ve Zaman",
            "Tarih Yazıcılığı",
            "İlk Çağ Medeniyetleri",
            "İlk Türk Devletleri",
            "İslam Tarihi ve Uygarlığı",
            "Türk - İslam Devletleri",
            "Türkiye Selçuklu Devleti",
            "Beylikten Devlete Osmanlı",
            "Dünya Gücü Osmanlı",
            "Arayış Yılları ve Islahatlar",
            "19. Yüzyılda Osmanlı",
            "I. Dünya Savaşı ve Sonuçları",
            "Milli Mücadele Dönemi",
            "Atatürk Dönemi İç ve Dış Politika",
            "II. Dünya Savaşı ve Sonrası",
            "Soğuk Savaş Dönemi",
            "Küreselleşen Dünya"
        ],
        "Coğrafya-1": [
            "Doğa ve İnsan",
            "Harita Bilgisi ve Coğrafi Konum",
            "Atmosfer ve İklim",
            "Yerin Şekillenmesi",
            "Jeolojik Zamanlar",
            "Nüfus ve Yerleşme",
            "Göçler ve Şehirleşme",
            "Ekonomik Faaliyetler",
            "Türkiye'nin Fiziki Özellikleri",
            "Türkiye'de İklim ve Bitki Örtüsü",
            "Türkiye'de Nüfus ve Yerleşme",
            "Türkiye'de Ekonomik Faaliyetler"
        ],
        "Tarih-2": [
            "Türk Dünyası ve Türk Devletleri",
            "Türk Kültürü ve Medeniyeti",
            "Osmanlı Devleti'nde Siyaset ve Toplum",
            "Osmanlı'da Ekonomik Hayat",
            "Avrupa'da Değişim ve Dönüşüm",
            "Sömürgecilik ve Sanayi Devrimi",
            "20. Yüzyıl Başlarında Dünya",
            "I. Dünya Savaşı Sonrası Dünya Düzeni",
            "II. Dünya Savaşı",
            "Soğuk Savaş Dönemi",
            "Küreselleşme ve Günümüz Dünya Siyasi Yapısı"
        ],
        "Coğrafya-2": [
            "Doğal Sistemler",
            "Beşeri Sistemler",
            "Ekonomik Coğrafya",
            "Bölgeler ve Ülkeler",
            "Jeopolitik ve Türkiye'nin Jeopolitik Konumu",
            "Küresel Çevre Sorunları",
            "Küresel ve Bölgesel Örgütler",
            "Çatışma Bölgeleri ve Göç"
        ],
        "Felsefe Grubu": [
            "Felsefeye Giriş",
            "Bilgi Felsefesi",
            "Varlık Felsefesi",
            "Ahlak Felsefesi",
            "Siyaset Felsefesi",
            "Sanat Felsefesi",
            "Bilim Felsefesi",
            "Din Felsefesi",
            "Psikolojiye Giriş",
            "Sosyolojiye Giriş",
            "Mantığa Giriş"
        ],
        "Din Kültürü ve Ahlak Bilgisi": [
            "İnanç ve İbadet",
            "Kur’an ve Yorumu",
            "Hz. Muhammed’in Hayatı ve Örnekliği",
            "Ahlak ve Değerler",
            "İslam Düşüncesinde Yoruma Açık Alanlar",
            "Din ve Kültür İlişkisi",
            "Güncel Dini Meseleler"
        ]
    },

    "5. Sınıf": {},
    "6. Sınıf": {},
    "7. Sınıf": {
        "Matematik": [
            "Tam Sayılar",
            "Tam Sayılarla Toplama İşlemi",
            "Tam Sayılarla Toplama İşleminin Özellikleri",
            "Tam Sayılarla Çıkarma İşlemi",
            "Tam Sayılarla Çarpma İşlemi",
            "Tam Sayılarla Çarpma İşleminin Özellikleri",
            "Tam Sayılarla Bölme İşlemi",
            "Tam Sayıların Kuvveti",
            "Tam Sayı Problemleri",
            "Rasyonel Sayılar",
            "Rasyonel Sayıların Ondalık Gösterimleri",
            "Ondalık Gösterimleri Rasyonel Sayıya Çevirme",
            "Rasyonel Sayılarda Sıralama",
            "Rasyonel Sayılarla Toplama-Çıkarma İşlemleri",
            "Rasyonel Sayılarla Toplama İşleminin Özellikleri",
            "Rasyonel Sayılarla Çarpma İşlemleri",
            "Rasyonel Sayılarla Çarpma İşleminin Özellikleri",
            "Rasyonel Sayılarla Bölme İşlemi",
            "Rasyonel Sayılarla Bölme İşleminde 0, 1 ve -1'in Etkisi",
            "Rasyonel Sayılarla Çok Adımlı İşlemler",
            "Rasyonel Sayıların Kareleri ve Küpleri",
            "Rasyonel Sayı Problemleri",
            "Cebirsel İfadeler",
            "Cebirsel İfadelerle Toplama ve Çıkarma İşlemi",
            "Bir Doğal Sayıyı Bir Cebirsel İfade ile Çarpma",
            "Örüntüler ve İlişkiler",
            "Birinci Dereceden Bir Bilinmeyenli Denklemler",
            "Eşitliğin Korunumu",
            "Denklem Çözme",
            "Birinci Dereceden Bir Bilinmeyenli Denklem Problemleri",
            "Oran ve Orantı",
            "Oran",
            "Orantı",
            "Doğru Orantı",
            "Doğru Orantılı İki Çokluğa Ait Orantı Sabiti",
            "Ters Orantı",
            "Doğru ve Ters Orantı Problemleri",
            "Yüzdeler",
            "Bir Çokluğun Belirtilen Yüzdesini Bulma",
            "Bir Çokluğu Diğer Bir Çokluğun Yüzdesi Olarak Hesaplama",
            "Bir Çokluğu Belirli Bir Yüzde ile Arttırma ve Azaltma",
            "Yüzde Problemleri",
            "Açılar",
            "Açıortay",
            "Aynı Düzlemde Üç Doğrunun Birbirine Göre Durumları",
            "Paralel İki Doğrunun Bir Kesene ile Yaptığı Açılar",
            "Çokgenlerin İç ve Dış Açıları",
            "Düzgün Çokgenler",
            "Dörtgenler",
            "Eşkenar Dörtgenin Alanı",
            "Yamuk Alanı",
            "Dörtgenlerin Alanları ile İlgili Problemler",
            "Çevre Alan İlişkisi",
            "Çemberde Merkez Açılar ve Bu Açılarının Gördüğü Yaylar",
            "Çemberin Çevre Uzunluğu",
            "Çember Parçasının Uzunluğu",
            "Dairenin Alanı",
            "Daire Diliminin Alanı",
            "Çizgi Grafiği",
            "Yanlış Yorumlamalara Neden Olabilecek Çizgi Grafikleri",
            "Ortalama, Ortanca, Tepe Değer",
            "Daire Grafiği",
            "Verilere Uygun Grafik Belirleme",
            "Farklı Yönlerden Görünümler",
            "Farklı Yönlerden Görünümleri Verilen Yapıları Oluşturma"
        ],
        "Fen": [
            "Uzay Araştırmaları",
            "Uzay Teknolojileri",
            "Uzay Kirliliği",
            "Teknoloji ve Uzay Araştırmaları",
            "Teleskop",
            "Güneş Sistemi Ötesi: Gök Cisimleri",
            "Bulutsu (Nebula)",
            "Yıldızlar",
            "Galaksiler",
            "Hücre",
            "Hücrenin Temel Kısımları",
            "DNA, Gen, Kromozom",
            "Geçmişten Günümüze Hücre",
            "Hücre-Doku-Organ-Sistem-Organizma",
            "Mitoz",
            "Hücre Bölünmesi (Mitoz)",
            "Mitoz Bölünmenin Canlılar İçin Önemi",
            "Mitoz Bölünmenin Evreleri",
            "Mayoz",
            "Mayoz Bölünme",
            "Mitoz ve Mayoz Bölünme Arasındaki Farklar",
            "Kütle ve Ağırlık İlişkisi",
            "Ağırlık Bir Kuvvettir",
            "Kütle ve Ağırlık Farklı Kavramlardır",
            "Kuvvet, İş ve Enerji İlişkisi",
            "Kuvvet ve İş",
            "Enerji ve Enerji Çeşitleri",
            "Enerji Dönüşümleri",
            "Kinetik ve Potansiyel Enerji Dönüşümleri",
            "Sürtünme Kuvveti ve Kinetik Enerji",
            "Maddenin Tanecikli Yapısı",
            "Atomun Yapısı",
            "Geçmişten Günümüze Atom Kavramı",
            "Moleküller",
            "Saf Maddeler",
            "Elementler ve Sembolleri",
            "Bileşik Formülleri",
            "Karışımlar",
            "Karışımların Ayrılması",
            "Karışımların Ayrılması (Ayırma Yöntemleri)",
            "Evsel Atıklar ve Geri Dönüşüm",
            "Evsel Atıklar ve Geri Dönüşümün Önemi",
            "Işığın Soğurulması",
            "Renklerin Oluşumu",
            "Güneş Enerjisinin Kullanım Alanları",
            "Aynalar",
            "Aynalar ve Kullanım Alanları",
            "Aynalarda Görüntü Oluşumu",
            "Işığın Kırılması ve Mercekler",
            "Işığın Kırılması",
            "Mercekler ve Kullanım Alanları",
            "İnsanda Üreme, Büyüme ve Gelişme",
            "İnsanda Üremeyi Sağlayan Yapı ve Organlar",
            "Bitki ve Hayvanlarda Üreme, Büyüme ve Gelişme",
            "Üreme",
            "Bitkilerde Büyüme ve Gelişme",
            "Hayvanlarda Büyüme ve Gelişme",
            "Ampullerin Bağlanma Şekilleri",
            "Ampullerin Seri ve Paralel Bağlanması",
            "Elektrik Akımı",
            "Akım Şiddeti ve Gerilim"
        ]
    },
    "8. Sınıf": {},
    "9. Sınıf": {},
    "10. Sınıf": {
        "Matematik": [
            "Üçgenler ve Trigonometri",
            "Dik Üçgende Trigonometrik Oranlar",
            "Trigonometrik Özdeşlikler",
            "Sinüs ve Kosinüs Teoremleri",
            "Nokta ve Doğrunun Analitiği",
            "İki Kategorik Değişken İçeren Dağılımlar",
            "İstatistiksel Araştırma Süreci",
            "İki Kategorik Değişkenin İlişkililiği",
            "Veri Toplama ve Analiz",
            "Bağımlı ve Bağımsız Olaylar",
            "Koşullu Olasılık",
            "Bir Doğal Sayı ile Asal Çarpanları ve Bölenleri Arasındaki İlişkiler",
            "Bir Doğal Sayının OBEB ve OKEK’i",
            "Bir Doğal Sayının Belirli Doğal Sayılara Bölümünden Kalanlar",
            "Sayma Stratejileri",
            "Algoritmik Dil ve Problem Çözme",
            "Fonksiyonun Tanımı ve Fonksiyon Olma Şartları",
            "Gerçek Sayılarda Tanımlı Fonksiyonların Temel Özellikleri",
            "Doğrusal Fonksiyonlar",
            "Karesel Fonksiyonlar",
            "Karekök Fonksiyonlar",
            "Rasyonel Fonksiyonlar",
            "Doğrusal, Karesel, Karekök ve Rasyonel Fonksiyonların Ters Fonksiyonları",
            "Doğrusal, Karesel, Karekök ve Rasyonel Fonksiyonları İçeren Denklem ve Eşitsizlikler",
            "Gerçek Sayılarda Tanımlı Kareköklü Fonksiyonların Özellikleri",
            "Gerçek Sayılarda Tanımlı Karekök Fonksiyonları ile İlgili Problemler",
            "Gerçek Sayılarda Tanımlı Rasyonel Fonksiyonların Özellikleri",
            "Gerçek Sayılarda Tanımlı Rasyonel Fonksiyonlar ile İlgili Problemler",
            "Dik Koordinat Sisteminde İki Nokta Arasındaki Uzaklık",
            "Dik Koordinat Sisteminde Doğrunun Eğim ve Analitik Özellikleri",
            "Bir Doğru Parçasını Oranla Bölen Noktanın Koordinatları",
            "Koşullu Olasılık",
            "Bağımlı ve Bağımsız Olaylar",
            "Bayes Teoremi ve Uygulamaları"
        ],
        "Kimya": [
            "Kimyasal Tepkimelerin Oluşumu",
            "Kimyasal Tepkime Türleri",
            "Mol Kavramı",
            "Kimyasal Tepkimelerde Stokiyometrik Hesaplamalar",
            "Gazların Özellikleri",
            "Gaz Yasaları",
            "İdeal Gaz Yasası",
            "Graham Difüzyon Yasası",
            "Graham Efüzyon Yasası",
            "Çözünme Süreci",
            "Derşim (Konsantrasyon) Birimleri",
            "Çözünürlük ve Çözünürlüğe Etki Eden Faktörler",
            "Çözeltilerin Sınıflandırılması",
            "Koligatif Özellikler",
            "Çevresel Sürdürülebilirlik",
            "Ekonomik Sürdürülebilirlik"
        ],
        "Fizik": [
            "Sabit Hızlı Hareket",
            "Sabit İvmeli Hareket",
            "Bir Boyutta Sabit İvmeli Hareket (Serbest Düşme)",
            "İki Boyutta Sabit İvmeli Hareket (Serbest Düşme)",
            "İş, Enerji ve Güç",
            "Enerji Çeşitleri ve Mekanik Enerji",
            "Dalgaların Genel Özellikleri",
            "Yay Dalgaları",
            "Ses Dalgaları",
            "Su Dalgaları",
            "Elektromanyetik Dalgalar"
        ],
    },
    "11. Sınıf": {},
    "12. Sınıf": {}
}


# --------------------------------------------------------------
#                      SINIF TABANLI UYGULAMA
# --------------------------------------------------------------


class CoachingApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Haftalık Koçluk Çizelgesi Doldurucu")
        self.root.resizable(True, True)

        try:
            self.root.state("zoomed")
        except Exception:
            self.root.update_idletasks()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.root.geometry(f"{sw}x{sh}+0+0")

        # Durum
        # ogrenciler[ogrenci_adi][hafta_adi][gun][saat] = {"text": str, "done": bool}
        self.ogrenciler: dict[str, dict] = {}
        self.ogrenci_notlari: dict[str, str] = {}
        self.ogrenci_alanlari: dict[str, str] = {}  # ogrenci_adi -> Sayısal / EA / Sözel
        self.ogrenci_tercihleri: dict[str, dict] = {}  # ogrenci_adi -> son seçimler
        self.denemeler: dict[str, list] = {}  # ogrenci_adi -> [{"tarih","ad","puan","tur","ayt_alan","turkce","mat","fen","sosyal","sos1","sos2"}, ...]
        self.aktif_ogrenci: Optional[str] = None
        self.aktif_hafta: Optional[str] = None

        self.list_items_map: list = []
        self.edit_mode: bool = False
        self.edit_prev_key: Optional[tuple] = None  # (gun, saat)

        self.excel_dosya_yolu: Optional[str] = None

        # Widget referansları
        self.ogrenci_combo = None
        self.ogrenci_alan_combo = None
        self.ogrenci_entry = None
        self.listbox = None
        self.liste_filtre_entry = None
        self.liste_filtre_var = None
        self.sadece_yapilmayan_var = None
        self.liste_sayac_label = None
        self.oneri_combo = None
        self.oneri_btn = None
        self.otoplan_btn = None
        self._last_auto_text = ""
        self.metin_entry = None
        self.sablon_combo = None
        self.excel_label = None
        self.ogrenci_not_text = None
        self.sinav_combo = None
        self.ders_combo = None
        self.konu_combo = None
        self.tablo_tree = None
        self.canvas = None
        self.content_frame = None
        self.hafta_combo = None
        self.kaynak_combo = None  # yayın seçimi
        self.saat_multi_listbox = None  # çoklu saat seçimi
        self.deneme_tur_combo = None
        self.deneme_tur_filtre = "Tümü"
        self.deneme_zaman_combo = None
        self.deneme_zaman_filtre = "Tümü"

        # Butonlar (state yönetimi için)
        self.ekle_btn = None
        self.dokuz_btn = None
        self.sil_btn = None
        self.duzenle_btn = None
        self.temizle_btn = None
        self.excel_btn = None
        self.pdf_btn = None
        self.ozet_btn = None
        self.istatistik_btn = None
        self.not_kaydet_btn = None
        self.yapildi_btn = None
        self.oneri_btn = None
        self.otoplan_btn = None

        # Veri (storage + servisler)
        self._storage = JsonStorage(_SCRIPT_DIR)
        self._excel_service = ExcelService(sheet_name=get_sheet_name(_SCRIPT_DIR))
        self._pdf_service = PdfService(font_path=get_pdf_font_path(_SCRIPT_DIR))
        self._yukle_ve_uygula()

        # UI
        self._setup_ui_style()
        self.build_ui()
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self.uygulamayi_kapat)

        # Otomatik yedekleme (5 dakikada bir)
        self._yedekleme_aralik_ms = 5 * 60 * 1000
        self._otomatik_yedekleme()

        # Başlangıç durumu
        self.ogrenci_combo["values"] = list(self.ogrenciler.keys())
        if self.aktif_ogrenci:
            self.ogrenci_combo.set(self.aktif_ogrenci)
            if self.ogrenci_alan_combo is not None:
                self.ogrenci_alan_combo.set(self._ogrenci_ayt_alani(self.aktif_ogrenci))
            if self.ogrenci_not_text is not None:
                self.ogrenci_not_text.insert(
                    "1.0", self.ogrenci_notlari.get(self.aktif_ogrenci, "")
                )

            ogr_data = self.ogrenciler[self.aktif_ogrenci]
            haftalar = list(ogr_data.keys())
            if not self.aktif_hafta or self.aktif_hafta not in haftalar:
                self.aktif_hafta = haftalar[0] if haftalar else "Hafta 1"
            if self.hafta_combo is not None:
                self.hafta_combo["values"] = haftalar
                self.hafta_combo.set(self.aktif_hafta)

        self.listeyi_guncelle()
        self.guncelle_ders_combo()
        self._ogrenci_tercihini_uygula()
        self._onerileri_yenile()
        self.update_button_states()

    # -------------- KALICI KAYIT -----------------

    def parse_entry(self, entry):
        """Kayıt değerini (dict/string/None) -> (text, done) çevirir."""
        if isinstance(entry, dict):
            return entry.get("text", ""), bool(entry.get("done", False))
        if entry is None:
            return "", False
        return str(entry), False

    @staticmethod
    def _normalize_saat_label(saat: str) -> str:
        """Saat etiketini HH:MM biçimine getirir (örn. 9:00 -> 09:00)."""
        parca = str(saat).split(":")
        if len(parca) == 2 and parca[0].isdigit() and parca[1].isdigit():
            return f"{int(parca[0]):02d}:{int(parca[1]):02d}"
        return str(saat)

    def _yukle_ve_uygula(self):
        """Storage'dan yükle; konu_verisi boşsa bu dosyadaki KONU_VERISI ile doldur ve kaydet."""
        global KONU_VERISI
        data = self._storage.load()
        self.ogrenciler = data["ogrenciler"]
        self.ogrenci_notlari = data["ogrenci_notlari"]
        self.ogrenci_alanlari = data.get("ogrenci_alanlari", {})
        self.ogrenci_tercihleri = data.get("ogrenci_tercihleri", {})
        self.denemeler = data.get("denemeler", {})
        self.aktif_ogrenci = data["aktif_ogrenci"]
        self.aktif_hafta = data["aktif_hafta"]
        if not data.get("konu_verisi"):
            data["konu_verisi"] = dict(KONU_VERISI)
            self._storage.save(data)
        KONU_VERISI = data["konu_verisi"]
        if self.ogrenciler and self.aktif_ogrenci not in self.ogrenciler:
            self.aktif_ogrenci = list(self.ogrenciler.keys())[0]
        # Alan bilgisini normalize/otomatik tamamla (eski veriler için)
        for ad in self.ogrenciler.keys():
            self.ogrenci_alanlari[ad] = self._ogrenci_ayt_alani(ad)

    def kaydet_diske(self):
        try:
            data = {
                "ogrenciler": self.ogrenciler,
                "ogrenci_notlari": self.ogrenci_notlari,
                "ogrenci_alanlari": self.ogrenci_alanlari,
                "ogrenci_tercihleri": self.ogrenci_tercihleri,
                "denemeler": self.denemeler,
                "konu_verisi": KONU_VERISI,
                "aktif_ogrenci": self.aktif_ogrenci,
                "aktif_hafta": self.aktif_hafta,
            }
            self._storage.save(data)
        except Exception as e:
            print("Kaydetme hatası:", e)

    def _otomatik_yedekleme(self):
        """5 dakikada bir diske kaydedip yedek dosyası oluşturur."""
        try:
            self.kaydet_diske()
            self._storage.backup()
        except Exception:
            pass
        try:
            self.root.after(self._yedekleme_aralik_ms, self._otomatik_yedekleme)
        except Exception:
            pass

    def disa_aktar_json(self):
        """Veriyi dışa aktar (json)."""
        default_name = f"program_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        path = filedialog.asksaveasfilename(
            title="Veriyi dışa aktar (JSON)",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        try:
            self.kaydet_diske()
            self._storage.export_to(path)
            messagebox.showinfo("Dışa Aktar", f"Veri dışa aktarıldı:\n{path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma başarısız:\n{e}")

    def ice_aktar_json(self):
        """Dışarıdan json verisini içe aktar."""
        path = filedialog.askopenfilename(
            title="Veri içe aktar (JSON)",
            filetypes=[("JSON", "*.json"), ("Tümü", "*.*")],
        )
        if not path:
            return
        if not messagebox.askyesno(
            "Onay",
            "Mevcut verinin üzerine yazılacak.\n"
            "Devam etmeden önce otomatik yedek alınacak.\n\n"
            "Devam edilsin mi?",
        ):
            return
        try:
            self.kaydet_diske()
            self._storage.backup()
            self._storage.import_from(path)
            self._yukle_ve_uygula()
            self._ui_data_refresh()
            messagebox.showinfo("İçe Aktar", "Veri başarıyla içe aktarıldı.")
        except Exception as e:
            messagebox.showerror("Hata", f"İçe aktarma başarısız:\n{e}")

    def yedekten_don(self):
        """Tarihli yedeklerden birini seçip geri yükler."""
        backups = self._storage.list_backups()
        if not backups:
            messagebox.showwarning("Yedek", "Henüz tarihli yedek bulunamadı.")
            return

        win = tk.Toplevel(self.root)
        win.title("Yedekten Dön")
        win.transient(self.root)
        tk.Label(win, text="Geri yüklenecek yedeği seç:", font=self._font_label).pack(padx=10, pady=(10, 5), anchor="w")
        lb = tk.Listbox(win, width=65, height=10, exportselection=False)
        for p in backups:
            lb.insert(tk.END, os.path.basename(p))
        lb.pack(padx=10, pady=5, fill="both", expand=True)
        lb.selection_set(0)

        def uygula():
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            target = backups[idx]
            if not messagebox.askyesno("Onay", f"Şu yedek geri yüklensin mi?\n{os.path.basename(target)}", parent=win):
                return
            try:
                self.kaydet_diske()
                self._storage.backup()
                self._storage.restore_backup(target)
                self._yukle_ve_uygula()
                self._ui_data_refresh()
                win.destroy()
                messagebox.showinfo("Yedek", "Yedek başarıyla geri yüklendi.")
            except Exception as e:
                messagebox.showerror("Hata", f"Yedek geri yüklenemedi:\n{e}", parent=win)

        btn = ttk.Frame(win)
        btn.pack(fill="x", padx=10, pady=10)
        tk.Button(btn, text="Geri Yükle", command=uygula, **self._btn_opts).pack(side="left")
        tk.Button(btn, text="İptal", command=win.destroy, **self._btn_opts).pack(side="right")

    def veriyi_sifirla(self):
        """Tüm kayıtları sıfırlar (önce otomatik yedek alır)."""
        if not messagebox.askyesno(
            "Veriyi Sıfırla",
            "Tüm öğrenciler, programlar, notlar ve deneme verileri sıfırlanacak.\n"
            "Önce otomatik yedek alınacak.\n\nDevam edilsin mi?",
        ):
            return
        try:
            self.kaydet_diske()
            self._storage.backup()
            self._storage.reset_data()
            self._yukle_ve_uygula()
            self._ui_data_refresh()
            messagebox.showinfo("Sıfırlandı", "Tüm veri sıfırlandı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Veri sıfırlanamadı:\n{e}")

    def _ui_data_refresh(self):
        """Storage'dan yüklenen veriyi arayüze yeniden uygula."""
        if self.ogrenci_combo is not None:
            self.ogrenci_combo["values"] = list(self.ogrenciler.keys())
            self.ogrenci_combo.set(self.aktif_ogrenci or "")
        if self.ogrenci_alan_combo is not None:
            aktif_alan = self._ogrenci_ayt_alani(self.aktif_ogrenci or "")
            self.ogrenci_alan_combo.set(aktif_alan)
        if self.hafta_combo is not None and self.aktif_ogrenci in self.ogrenciler:
            haftalar = list(self.ogrenciler[self.aktif_ogrenci].keys())
            self.hafta_combo["values"] = haftalar
            self.hafta_combo.set(self.aktif_hafta or (haftalar[0] if haftalar else ""))
        if self.ogrenci_not_text is not None:
            self.ogrenci_not_text.delete("1.0", tk.END)
            if self.aktif_ogrenci:
                self.ogrenci_not_text.insert("1.0", self.ogrenci_notlari.get(self.aktif_ogrenci, ""))
        self.listeyi_guncelle()
        self.guncelle_ders_combo()
        self._ogrenci_tercihini_uygula()
        self._onerileri_yenile()
        self.update_button_states()

    # -------------- YARDIMCI -----------------

    def aktif_program(self):
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci ekleyip seçmelisin.")
            return None

        ogr_data = self.ogrenciler[self.aktif_ogrenci]  # {"Hafta 1": {...}, ...}

        # Aktif hafta yoksa veya öğrencide yoksa ilk haftayı seç
        if not self.aktif_hafta or self.aktif_hafta not in ogr_data:
            if ogr_data:
                self.aktif_hafta = sorted(ogr_data.keys())[0]
            else:
                self.aktif_hafta = "Hafta 1"
                ogr_data[self.aktif_hafta] = {g: {} for g in gunler}

            if self.hafta_combo is not None:
                self.hafta_combo["values"] = list(ogr_data.keys())
                self.hafta_combo.set(self.aktif_hafta)

        return ogr_data[self.aktif_hafta]

    def _grafik_verileri(self):
        """Grafikler için veri: (bu hafta günlük plan/yapılan, son haftalar tamamlanma %)."""
        gunluk_plan, gunluk_yapilan = [], []
        hafta_isimleri, hafta_oranlari = [], []
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            return gunler, gunluk_plan, gunluk_yapilan, hafta_isimleri, hafta_oranlari
        ogr_data = self.ogrenciler[self.aktif_ogrenci]
        program = self.aktif_program()
        if program is None:
            return gunler, gunluk_plan, gunluk_yapilan, hafta_isimleri, hafta_oranlari
        gt, _, _, _, gy, _ = self.gunluk_istatistik(program)
        for g in gunler:
            gunluk_plan.append(gt.get(g, 0))
            gunluk_yapilan.append(gy.get(g, 0))
        # Aktif haftadan itibaren (maks 8) haftaları göster
        def _hafta_key(hafta_adi: str):
            ad = str(hafta_adi).strip()
            parca = ad.split()
            if len(parca) >= 2 and parca[0].lower() == "hafta" and parca[1].isdigit():
                return (0, int(parca[1]))
            return (1, ad.lower())

        haftalar = sorted(ogr_data.keys(), key=_hafta_key)
        if self.aktif_hafta in haftalar:
            idx = haftalar.index(self.aktif_hafta)
            haftalar = haftalar[idx:]
        haftalar = haftalar[:8]

        for h in haftalar:
            p = ogr_data[h]
            top = sum(len(p.get(g, {})) for g in gunler)
            yap = sum(1 for g in gunler for _, entry in p.get(g, {}).items() if self.parse_entry(entry)[1])
            hafta_isimleri.append(h)
            hafta_oranlari.append(round(100 * yap / top) if top else 0)
        return gunler, gunluk_plan, gunluk_yapilan, hafta_isimleri, hafta_oranlari

    def _grafikleri_ciz(self):
        """Sağ panelde günlük, haftalık, deneme puanı ve net trendi grafiklerini çizer."""
        try:
            self.grafik_canvas.delete("all")
        except Exception:
            return
        gunler_etiket = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]
        günler, plan, yapilan, hafta_isim, hafta_oran = self._grafik_verileri()
        deneme_liste = []
        if self.aktif_ogrenci and self.aktif_ogrenci in self.denemeler:
            tum = sorted(self.denemeler[self.aktif_ogrenci], key=lambda x: x.get("tarih", ""))
            deneme_liste = self._deneme_listesi_filtrele(tum)[-10:]
        if not plan and not hafta_oran and not deneme_liste:
            self.grafik_canvas.create_text(200, 120, text="Öğrenci ve hafta seçin;\nistatistik burada görünecek.\n\nDeneme eklemek için\n'Deneme ekle' butonunu kullanın.", fill=self._fg, font=self._font_label)
            return
        w = self.grafik_canvas.winfo_width() or 340
        h = self.grafik_canvas.winfo_height() or 620
        if w < 100:
            w = 340
        if h < 100:
            h = 620
        pad_left, pad_right, pad_top = 50, 30, 36
        chart_w = w - pad_left - pad_right
        chart_h = (h - pad_top - 90) // 4
        max_plan = max(plan) if plan else 1
        max_plan = max(max_plan, 1)
        bar_w = max(8, (chart_w // len(gunler)) - 6) if gunler else 20
        # 1) Bu hafta günlük
        self.grafik_canvas.create_text(pad_left + chart_w // 2, 14, text="Bu hafta — günlük (plan / yapılan)", fill=self._fg, font=("Segoe UI", 9, "bold"))
        y0 = pad_top + chart_h - 22
        for i, g in enumerate(gunler):
            x = pad_left + i * (chart_w // len(gunler)) + 4
            p_val = plan[i] if i < len(plan) else 0
            y_val = yapilan[i] if i < len(yapilan) else 0
            h_plan = int((p_val / max_plan) * (chart_h - 42)) if max_plan else 0
            h_yap = int((y_val / max_plan) * (chart_h - 42)) if max_plan else 0
            self.grafik_canvas.create_rectangle(x, y0 - h_plan, x + bar_w, y0, outline=self._label_muted, fill=self._btn_bg)
            self.grafik_canvas.create_rectangle(x, y0 - h_yap, x + bar_w, y0, outline=self._accent, fill=self._accent)
            self.grafik_canvas.create_text(x + bar_w // 2, y0 + 12, text=gunler_etiket[i] if i < 7 else "", fill=self._fg, font=("Segoe UI", 8))
        # 2) Haftalık tamamlanma %
        self.grafik_canvas.create_text(pad_left + chart_w // 2, pad_top + chart_h + 14, text="Haftalık tamamlanma oranı (%)", fill=self._fg, font=("Segoe UI", 9, "bold"))
        y0_2 = pad_top + chart_h + chart_h - 22
        n_hafta = len(hafta_isim)
        if n_hafta:
            bar_w2 = max(12, (chart_w // n_hafta) - 4)
            for i in range(n_hafta):
                x = pad_left + i * (chart_w // n_hafta) + 4
                oran = hafta_oran[i] if i < len(hafta_oran) else 0
                h_bar = int((oran / 100.0) * (chart_h - 42)) if oran else 0
                self.grafik_canvas.create_rectangle(x, y0_2, x + bar_w2, y0_2 - h_bar, outline=self._accent, fill=self._accent)
                if i < len(hafta_isim):
                    ad = str(hafta_isim[i])
                    if ad.lower().startswith("hafta "):
                        parca = ad.split()
                        lbl = f"Hafta\n{parca[1]}" if len(parca) > 1 else ad
                    else:
                        lbl = ad if len(ad) <= 10 else (ad[:10] + "…")
                else:
                    lbl = ""
                self.grafik_canvas.create_text(
                    x + bar_w2 // 2,
                    y0_2 + 16,
                    text=lbl,
                    fill=self._fg,
                    font=("Segoe UI", 8),
                    justify="center",
                )
                if h_bar > 12:
                    self.grafik_canvas.create_text(x + bar_w2 // 2, y0_2 - h_bar - 8, text=f"%{oran}", fill=self._fg, font=("Segoe UI", 8))
        # 3) Deneme toplam neti
        y0_3 = pad_top + 2 * chart_h + chart_h - 22
        y3_top = pad_top + 2 * chart_h + 28  # başlıkla çakışmaması için üst boşluk
        self.grafik_canvas.create_text(pad_left + chart_w // 2, pad_top + 2 * chart_h + 14, text="Deneme toplam neti", fill=self._fg, font=("Segoe UI", 9, "bold"))
        if deneme_liste:
            aktif_alan = self._ogrenci_ayt_alani()
            puans = [self._deneme_toplam_net(d, aktif_alan) for d in deneme_liste]
            max_puan = max(puans) if puans else 100
            max_puan = max(max_puan, 1)
            n_d = len(deneme_liste)
            bar_w3 = max(14, (chart_w // n_d) - 4)
            usable_h3 = max(26, y0_3 - y3_top)
            for i in range(n_d):
                x = pad_left + i * (chart_w // n_d) + 4
                puan = puans[i]
                h_bar3 = int((puan / max_puan) * usable_h3) if max_puan else 0
                self.grafik_canvas.create_rectangle(x, y0_3, x + bar_w3, y0_3 - h_bar3, outline=self._accent, fill=self._accent)
                lbl = (deneme_liste[i].get("ad") or deneme_liste[i]["tarih"][-5:])[:8]
                self.grafik_canvas.create_text(x + bar_w3 // 2, y0_3 + 12, text=lbl, fill=self._fg, font=("Segoe UI", 7))
                if h_bar3 > 10:
                    y_val = max(y3_top + 7, y0_3 - h_bar3 - 8)
                    self.grafik_canvas.create_text(x + bar_w3 // 2, y_val, text=f"{puan:.1f}", fill=self._fg, font=("Segoe UI", 8))
        else:
            ham_var = bool(self.aktif_ogrenci and self.aktif_ogrenci in self.denemeler and self.denemeler[self.aktif_ogrenci])
            msg = (
                "Bu dönem / tür için deneme yok.\nFiltreleri 'Tümü' yapın veya\nbaşka aralık seçin."
                if ham_var
                else "Henüz deneme yok.\n'Deneme ekle' ile ekleyin."
            )
            self.grafik_canvas.create_text(pad_left + chart_w // 2, y0_3 - (chart_h - 42) // 2, text=msg, fill=self._label_muted, font=self._font_label)

        # 4) Ders net trendi (son 6 deneme)
        y0_4 = pad_top + 4 * chart_h - 22
        self.grafik_canvas.create_text(pad_left + chart_w // 2, pad_top + 3 * chart_h + 14, text="Ders net trendi (son 6 deneme)", fill=self._fg, font=("Segoe UI", 9, "bold"))
        trend = deneme_liste[-6:]
        if trend:
            is_ayt_view = self.deneme_tur_filtre == "AYT"
            if is_ayt_view:
                first_subject_values = [float(exam.get("mat", 0) or 0) for exam in trend]
                second_subject_values = [float(exam.get("fen", 0) or 0) for exam in trend]
                third_subject_values = [float(exam.get("sos1", 0) or 0) for exam in trend]
                fourth_subject_values = [float(exam.get("sos2", 0) or 0) for exam in trend]
                legend = [("Mat", "#1e88e5"), ("Fen", "#43a047"), ("Edb-S1", "#8e24aa"), ("Sos-2", "#fb8c00")]
                max_net = max(first_subject_values + second_subject_values + third_subject_values + fourth_subject_values) if trend else 1
            else:
                first_subject_values = [float(exam.get("turkce", 0) or 0) for exam in trend]
                second_subject_values = [float(exam.get("mat", 0) or 0) for exam in trend]
                third_subject_values = [float(exam.get("fen", 0) or 0) for exam in trend]
                fourth_subject_values = [float(exam.get("sosyal", 0) or 0) for exam in trend]
                legend = [("Tr", "#8e24aa"), ("Mat", "#1e88e5"), ("Fen", "#43a047"), ("Sos", "#fb8c00")]
                max_net = max(first_subject_values + second_subject_values + third_subject_values + fourth_subject_values) if trend else 1
            max_net = max(max_net, 1)
            n_t = len(trend)
            slot_w = max(28, chart_w // max(n_t, 1))
            bar_w4 = max(4, (slot_w - 8) // 4)

            # Legend
            lx = pad_left + 25
            ly = pad_top + 3 * chart_h + 30
            for j, (name, color) in enumerate(legend):
                self.grafik_canvas.create_text(lx + j * 55, ly, text=f"■ {name}", fill=color, font=("Segoe UI", 8), anchor="w")

            # Son denemelerin ders ders net özeti
            summary_lines = 0
            if self.deneme_tur_filtre == "Tümü":
                son_tyt = next((d for d in reversed(trend) if (d.get("tur") or "TYT") == "TYT"), None)
                son_ayt = next((d for d in reversed(trend) if (d.get("tur") or "TYT") == "AYT"), None)
                y_ozet = pad_top + 3 * chart_h + 44
                if son_tyt is not None:
                    ozet_tyt = (
                        f"Son TYT -> Tr: {float(son_tyt.get('turkce', 0) or 0):.1f} | "
                        f"Mat: {float(son_tyt.get('mat', 0) or 0):.1f} | "
                        f"Fen: {float(son_tyt.get('fen', 0) or 0):.1f} | "
                        f"Sos: {float(son_tyt.get('sosyal', 0) or 0):.1f}"
                    )
                    self.grafik_canvas.create_text(
                        pad_left + chart_w // 2,
                        y_ozet,
                        text=ozet_tyt,
                        fill=self._fg,
                        font=("Segoe UI", 7),
                    )
                    summary_lines += 1
                    y_ozet += 12
                if son_ayt is not None:
                    ayt_alan = self._ogrenci_ayt_alani()
                    ozet_ayt = (
                        f"Son AYT ({ayt_alan}) -> Mat: {float(son_ayt.get('mat', 0) or 0):.1f} | "
                        f"Fen: {float(son_ayt.get('fen', 0) or 0):.1f} | "
                        f"Edb-S1: {float(son_ayt.get('sos1', 0) or 0):.1f} | "
                        f"Sos-2: {float(son_ayt.get('sos2', 0) or 0):.1f} | Toplam: {self._deneme_toplam_net(son_ayt, ayt_alan):.1f}"
                    )
                    self.grafik_canvas.create_text(
                        pad_left + chart_w // 2,
                        y_ozet,
                        text=ozet_ayt,
                        fill=self._fg,
                        font=("Segoe UI", 7),
                    )
                    summary_lines += 1
            else:
                son = trend[-1]
                if is_ayt_view:
                    ayt_alan = self._ogrenci_ayt_alani()
                    ozet = (
                        f"Son deneme ({ayt_alan}) -> Mat: {float(son.get('mat', 0) or 0):.1f} | "
                        f"Fen: {float(son.get('fen', 0) or 0):.1f} | "
                        f"Edb-S1: {float(son.get('sos1', 0) or 0):.1f} | "
                        f"Sos-2: {float(son.get('sos2', 0) or 0):.1f} | Toplam: {self._deneme_toplam_net(son, ayt_alan):.1f}"
                    )
                else:
                    ozet = (
                        f"Son deneme netleri -> Tr: {float(son.get('turkce', 0) or 0):.1f} | "
                        f"Mat: {float(son.get('mat', 0) or 0):.1f} | "
                        f"Fen: {float(son.get('fen', 0) or 0):.1f} | "
                        f"Sos: {float(son.get('sosyal', 0) or 0):.1f} | Toplam: {self._deneme_toplam_net(son):.1f}"
                    )
                self.grafik_canvas.create_text(
                    pad_left + chart_w // 2,
                    pad_top + 3 * chart_h + 44,
                    text=ozet,
                    fill=self._fg,
                    font=("Segoe UI", 7),
                )
                summary_lines = 1

            baseline = y0_4
            usable_h = chart_h - (74 + (summary_lines * 12))
            for i in range(n_t):
                x0 = pad_left + i * slot_w + 6
                h1 = int((first_subject_values[i] / max_net) * usable_h)
                h2 = int((second_subject_values[i] / max_net) * usable_h)
                h3 = int((third_subject_values[i] / max_net) * usable_h)
                h4 = int((fourth_subject_values[i] / max_net) * usable_h)
                c1, c2, c3, c4 = legend[0][1], legend[1][1], legend[2][1], legend[3][1]
                self.grafik_canvas.create_rectangle(x0, baseline, x0 + bar_w4, baseline - h1, outline=c1, fill=c1)
                self.grafik_canvas.create_rectangle(x0 + bar_w4 + 1, baseline, x0 + 2 * bar_w4 + 1, baseline - h2, outline=c2, fill=c2)
                self.grafik_canvas.create_rectangle(x0 + 2 * bar_w4 + 2, baseline, x0 + 3 * bar_w4 + 2, baseline - h3, outline=c3, fill=c3)
                self.grafik_canvas.create_rectangle(x0 + 3 * bar_w4 + 3, baseline, x0 + 4 * bar_w4 + 3, baseline - h4, outline=c4, fill=c4)
                # Her dersin netini bar üstünde göster
                sayi_goster = (n_t <= 4 and bar_w4 >= 8)
                if sayi_goster and h1 >= 10:
                    self.grafik_canvas.create_text(x0 + bar_w4 // 2, baseline - h1 - 7, text=f"{first_subject_values[i]:.1f}", fill=self._fg, font=("Segoe UI", 6))
                if sayi_goster and h2 >= 10:
                    self.grafik_canvas.create_text(x0 + (3 * bar_w4) // 2 + 1, baseline - h2 - 7, text=f"{second_subject_values[i]:.1f}", fill=self._fg, font=("Segoe UI", 6))
                if sayi_goster and h3 >= 10:
                    self.grafik_canvas.create_text(x0 + (5 * bar_w4) // 2 + 2, baseline - h3 - 7, text=f"{third_subject_values[i]:.1f}", fill=self._fg, font=("Segoe UI", 6))
                if sayi_goster and h4 >= 10:
                    self.grafik_canvas.create_text(x0 + (7 * bar_w4) // 2 + 3, baseline - h4 - 7, text=f"{fourth_subject_values[i]:.1f}", fill=self._fg, font=("Segoe UI", 6))
                lbl = (trend[i].get("ad") or trend[i].get("tarih", ""))[:6]
                self.grafik_canvas.create_text(x0 + 2 * bar_w4, baseline + 10, text=lbl, fill=self._fg, font=("Segoe UI", 7))
        else:
            ham_var = bool(self.aktif_ogrenci and self.aktif_ogrenci in self.denemeler and self.denemeler[self.aktif_ogrenci])
            msg = (
                "Bu dönem / tür için net verisi yok."
                if ham_var
                else "Net trendi için deneme ekleyin."
            )
            self.grafik_canvas.create_text(pad_left + chart_w // 2, y0_4 - (chart_h - 42) // 2, text=msg, fill=self._label_muted, font=self._font_label)

    def _ogrenci_ayt_alani(self, ogrenci_adi: Optional[str] = None) -> str:
        """Öğrencinin AYT alanını döndürür; yoksa isim/denemeden tahmin eder."""
        ad = ogrenci_adi or self.aktif_ogrenci or ""
        ad_l = ad.lower()
        # İsimde alan etiketi varsa en yüksek öncelik (örn: "(EA)", " - EA ", "(Sözel)")
        if re.search(r"\bea\b", ad_l):
            return "EA"
        if re.search(r"s[öo]zel", ad_l):
            return "Sözel"
        if re.search(r"say[ıi]sal", ad_l):
            return "Sayısal"

        alan = (self.ogrenci_alanlari.get(ad) or "").strip()
        if alan in ("Sayısal", "EA", "Sözel"):
            return alan
        alan_l = alan.lower()
        if alan_l in ("sayisal", "sayısal"):
            return "Sayısal"
        if alan_l in ("sozel", "sözel"):
            return "Sözel"
        if alan_l in ("ea", "eşit ağırlık", "esit agirlik", "eşit agirlik"):
            return "EA"

        # 2) Varsa son AYT denemedeki alandan tahmin et
        liste = self.denemeler.get(ad, []) if isinstance(self.denemeler, dict) else []
        for d in reversed(liste):
            if (d.get("tur") or "TYT") != "AYT":
                continue
            a = (d.get("ayt_alan") or "").strip().lower()
            if a in ("ea", "eşit ağırlık", "esit agirlik", "eşit agirlik"):
                return "EA"
            if a in ("sozel", "sözel"):
                return "Sözel"
            if a in ("sayisal", "sayısal"):
                return "Sayısal"
        return "Sayısal"

    def ogrenci_alani_degisti(self, event=None):
        """Aktif öğrenci için alan seçimini kaydeder."""
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            return
        if self.ogrenci_alan_combo is None:
            return
        alan = (self.ogrenci_alan_combo.get() or "").strip()
        if alan not in ("Sayısal", "EA", "Sözel"):
            alan = "Sayısal"
        self.ogrenci_alanlari[self.aktif_ogrenci] = alan
        self.kaydet_diske()
        try:
            self._grafikleri_ciz()
        except Exception:
            pass

    def _deneme_toplam_net(self, d: dict, ogrenci_ayt_alani: Optional[str] = None) -> float:
        """Deneme türüne ve (AYT ise) alana göre toplam neti hesaplar."""
        tur = (d.get("tur") or "TYT").strip()
        mat = float(d.get("mat", 0) or 0)
        fen = float(d.get("fen", 0) or 0)
        turkce = float(d.get("turkce", 0) or 0)
        sosyal = float(d.get("sosyal", 0) or 0)
        sos1 = float(d.get("sos1", 0) or 0)
        sos2 = float(d.get("sos2", 0) or 0)
        if tur == "AYT":
            alan = (ogrenci_ayt_alani or d.get("ayt_alan") or "Sayısal").strip().lower()
            if alan in ("ea", "eşit ağırlık", "esit agirlik", "eşit agirlik"):
                return mat + sos1
            if alan in ("sozel", "sözel"):
                return sos1 + sos2
            return mat + fen
        return turkce + mat + fen + sosyal

    @staticmethod
    def _deneme_tarih_parse(d: dict):
        """Deneme kaydındaki tarihi date nesnesine çevirir; olmazsa None."""
        t = (d.get("tarih") or "").strip()[:10]
        if len(t) < 10:
            return None
        try:
            return datetime.strptime(t, "%Y-%m-%d").date()
        except ValueError:
            return None

    def _deneme_listesi_filtrele(self, tum_liste: list) -> list:
        """Zaman (son 30 gün / 3 ay) ve tür (TYT/AYT) filtresini uygular."""
        bugun = date.today()
        if self.deneme_zaman_filtre == "Son 30 gün":
            esik = bugun - timedelta(days=30)
        elif self.deneme_zaman_filtre == "Son 3 ay":
            esik = bugun - timedelta(days=90)
        else:
            esik = None
        out = []
        for d in tum_liste:
            if esik is not None:
                td = self._deneme_tarih_parse(d)
                if td is None or td < esik:
                    continue
            if self.deneme_tur_filtre in ("TYT", "AYT"):
                if (d.get("tur") or "TYT") != self.deneme_tur_filtre:
                    continue
            out.append(d)
        return out

    def _deneme_zaman_filtre_degisti(self, event=None):
        """Dönem filtresi değiştiğinde grafikleri yeniler."""
        if self.deneme_zaman_combo is not None:
            secim = self.deneme_zaman_combo.get().strip()
            self.deneme_zaman_filtre = secim if secim in ("Tümü", "Son 30 gün", "Son 3 ay") else "Tümü"
        self._grafikleri_ciz()

    def _deneme_tur_filtre_degisti(self, event=None):
        """Deneme tür filtresi (Tümü/TYT/AYT) değiştiğinde grafikleri yeniler."""
        if self.deneme_tur_combo is not None:
            secim = self.deneme_tur_combo.get().strip()
            self.deneme_tur_filtre = secim if secim in ("Tümü", "TYT", "AYT") else "Tümü"
        self._grafikleri_ciz()

    def gunluk_istatistik(self, program: dict):
        gunluk_toplam = {}
        gunluk_yapilan = {}

        for gun in gunler:
            gun_prog = program.get(gun, {})
            gunluk_toplam[gun] = len(gun_prog)

            yapilan = 0
            for saat, entry in gun_prog.items():
                _, done = self.parse_entry(entry)
                if done:
                    yapilan += 1
            gunluk_yapilan[gun] = yapilan

        haftalik_toplam = sum(gunluk_toplam.values())
        haftalik_yapilan = sum(gunluk_yapilan.values())

        if gunluk_toplam:
            en_yogun_gun = max(gunluk_toplam, key=lambda g: gunluk_toplam[g])
            en_bos_gun = min(gunluk_toplam, key=lambda g: gunluk_toplam[g])
        else:
            en_yogun_gun = en_bos_gun = None

        return (
            gunluk_toplam,
            haftalik_toplam,
            en_yogun_gun,
            en_bos_gun,
            gunluk_yapilan,
            haftalik_yapilan,
        )

    def aktif_ogrenci_notunu_guncelle(self):
        if self.aktif_ogrenci is not None and self.ogrenci_not_text is not None:
            metin = self.ogrenci_not_text.get("1.0", tk.END).strip()
            self.ogrenci_notlari[self.aktif_ogrenci] = metin

    # -------------- ÇİZELGE -----------------

    def guncelle_tablo(self):
        if self.tablo_tree is None:
            return

        for item in self.tablo_tree.get_children():
            self.tablo_tree.delete(item)

        program = self.aktif_program()
        if program is None:
            return

        for saat in saatler:
            satir = [saat]
            for gun in gunler:
                metin = ""
                gun_prog = program.get(gun, {})
                if saat in gun_prog:
                    text, done = self.parse_entry(gun_prog[saat])
                    metin = ("✓ " if done else "") + text
                satir.append(metin)
            self.tablo_tree.insert("", "end", values=satir)

    def listeyi_guncelle(self):
        self.listbox.delete(0, tk.END)
        self.list_items_map.clear()

        program = self.aktif_program()
        if program is None:
            if self.liste_sayac_label is not None:
                self.liste_sayac_label.configure(text="Gösterilen: 0 / 0")
            return

        filtre = ""
        if self.liste_filtre_var is not None:
            filtre = (self.liste_filtre_var.get() or "").strip().lower()
        sadece_yapilmayan = bool(self.sadece_yapilmayan_var.get()) if self.sadece_yapilmayan_var is not None else False
        toplam_kayit = 0
        gosterilen_kayit = 0

        for gun in gunler:
            gun_prog = program.get(gun, {})
            if gun_prog:
                gun_satirlari = []
                for saat in sorted(gun_prog.keys(), key=lambda s: TIME_TO_ROW.get(self._normalize_saat_label(s), 999)):
                    saat_norm = self._normalize_saat_label(saat)
                    entry = gun_prog[saat]
                    text, done = self.parse_entry(entry)
                    toplam_kayit += 1
                    if sadece_yapilmayan and done:
                        continue
                    haystack = f"{gun} {saat_norm} {text}".lower()
                    if filtre and filtre not in haystack:
                        continue
                    durum = "✓" if done else " "
                    gun_satirlari.append((f"[{durum}] {saat_norm} -> {text}", (gun, saat)))

                if gun_satirlari:
                    self.listbox.insert(tk.END, f"--- {gun} ---")
                    self.list_items_map.append(None)
                    for satir, key in gun_satirlari:
                        self.listbox.insert(tk.END, satir)
                        self.list_items_map.append(key)
                        gosterilen_kayit += 1

        if self.liste_sayac_label is not None:
            self.liste_sayac_label.configure(text=f"Gösterilen: {gosterilen_kayit} / {toplam_kayit}")

        self.guncelle_tablo()
        try:
            self._grafikleri_ciz()
        except Exception:
            pass

    def liste_filtresi_degisti(self, event=None):
        """Kayıt listesi filtresi değiştiğinde listeyi anlık yeniler."""
        self.listeyi_guncelle()

    def liste_filtresini_temizle(self):
        """Kayıt listesi filtrelerini temizler."""
        if self.liste_filtre_var is not None:
            self.liste_filtre_var.set("")
        if self.sadece_yapilmayan_var is not None:
            self.sadece_yapilmayan_var.set(False)
        self.listeyi_guncelle()

    # -------------- ÖĞRENCİ -----------------

    def ogrenci_ekle_veya_sec(self):
        ad = self.ogrenci_entry.get().strip()
        if not ad:
            messagebox.showwarning("Uyarı", "Öğrenci adı soyadı boş olamaz.")
            return

        if self.aktif_ogrenci is not None:
            self.aktif_ogrenci_notunu_guncelle()

        if ad not in self.ogrenciler:
            # Yeni öğrenci için başlangıçta sadece Hafta 1 oluştur
            self.ogrenciler[ad] = {
                "Hafta 1": {g: {} for g in gunler}
            }
            self.ogrenci_notlari[ad] = ""
        secilen_alan = "Sayısal"
        if self.ogrenci_alan_combo is not None:
            secilen_alan = self.ogrenci_alan_combo.get().strip() or "Sayısal"
        self.ogrenci_alanlari[ad] = secilen_alan

        self.aktif_ogrenci = ad
        self.ogrenci_combo["values"] = list(self.ogrenciler.keys())
        self.ogrenci_combo.set(ad)
        self._ogrenci_tercihini_kaydet()

        ogr_data = self.ogrenciler[ad]
        haftalar = list(ogr_data.keys())
        self.aktif_hafta = "Hafta 1" if "Hafta 1" in haftalar else haftalar[0]
        if self.hafta_combo is not None:
            self.hafta_combo["values"] = haftalar
            self.hafta_combo.set(self.aktif_hafta)

        self.listeyi_guncelle()
        self._ogrenci_tercihini_uygula()
        self._onerileri_yenile()

        if self.ogrenci_not_text is not None:
            self.ogrenci_not_text.delete("1.0", tk.END)
            self.ogrenci_not_text.insert("1.0", self.ogrenci_notlari.get(self.aktif_ogrenci, ""))

        self.kaydet_diske()
        self.update_button_states()
        messagebox.showinfo("Bilgi", f"Aktif öğrenci: {self.aktif_ogrenci}")

    def ogrenci_degisti(self, event=None):
        if self.aktif_ogrenci is not None:
            self.aktif_ogrenci_notunu_guncelle()

        ad = self.ogrenci_combo.get()
        if ad in self.ogrenciler:
            self.aktif_ogrenci = ad

            ogr_data = self.ogrenciler[ad]
            haftalar = list(ogr_data.keys())
            if self.hafta_combo is not None:
                self.hafta_combo["values"] = haftalar

            if not self.aktif_hafta or self.aktif_hafta not in haftalar:
                self.aktif_hafta = haftalar[0] if haftalar else "Hafta 1"
            if self.hafta_combo is not None:
                self.hafta_combo.set(self.aktif_hafta)

            self.listeyi_guncelle()
            self._ogrenci_tercihini_uygula()
            self._onerileri_yenile()
            if self.ogrenci_not_text is not None:
                self.ogrenci_not_text.delete("1.0", tk.END)
                self.ogrenci_not_text.insert("1.0", self.ogrenci_notlari.get(self.aktif_ogrenci, ""))

            self.guncelle_ders_combo()
            if self.ogrenci_alan_combo is not None:
                self.ogrenci_alan_combo.set(self._ogrenci_ayt_alani(ad))
            self.kaydet_diske()
            self.update_button_states()

    # -------------- KAYIT EKLE / SİL / DÜZENLE / YAPILDI -----------------

    def ekle_kayit(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        gun = self.gun_combo.get()
        metin = self.metin_entry.get().strip()

        if not gun or not metin:
            messagebox.showwarning("Eksik Bilgi", "Gün ve metin boş olamaz.")
            return

        gun_prog = program[gun]

        # --- DÜZENLEME MODU (tek saat) ---
        if self.edit_mode and self.edit_prev_key is not None:
            saat = self.saat_combo.get()
            if not saat:
                messagebox.showwarning("Eksik Bilgi", "Saat seçmelisin.")
                return

            mevcut_var = gun_prog.get(saat) is not None
            if mevcut_var and not (self.edit_prev_key == (gun, saat)):
                devam = messagebox.askyesno(
                    "Üzerine Yazılsın mı?",
                    f"{gun} - {saat} saatinde zaten bir kayıt var.\n"
                    f"Üzerine yazmak istiyor musun?"
                )
                if not devam:
                    return

            done_flag = False
            eski_gun, eski_saat = self.edit_prev_key
            eski_entry = program.get(eski_gun, {}).get(eski_saat)
            if eski_entry is not None:
                _, done_flag = self.parse_entry(eski_entry)
            # Eski kaydı sil
            if eski_saat in program.get(eski_gun, {}):
                del program[eski_gun][eski_saat]

            gun_prog[saat] = {"text": metin, "done": done_flag}

        # --- NORMAL MOD (birden fazla saat eklenebilir) ---
        else:
            hedef_saatler = []

            # Çoklu listbox'tan seçilen saatler
            if self.saat_multi_listbox is not None:
                secili_indeksler = self.saat_multi_listbox.curselection()
                hedef_saatler = [self.saat_multi_listbox.get(i) for i in secili_indeksler]

            # Çoklu seçim yoksa combobox'taki tek saat kullan
            if not hedef_saatler:
                saat = self.saat_combo.get()
                if not saat:
                    messagebox.showwarning("Eksik Bilgi", "En az bir saat seçmelisin.")
                    return
                hedef_saatler = [saat]

            # Var olan kayıtlar için uyarı
            dolu_saatler = [s for s in hedef_saatler if s in gun_prog]
            if dolu_saatler:
                devam = messagebox.askyesno(
                    "Üzerine Yazılsın mı?",
                    f"{gun} gününde şu saatlerde zaten kayıt var:\n"
                    f"{', '.join(dolu_saatler)}\n\n"
                    "Bu saatlerin ÜZERİNE yazmak istiyor musun?"
                )
                if not devam:
                    return

            for s in hedef_saatler:
                eski = gun_prog.get(s)
                done_flag = False
                if isinstance(eski, dict):
                    _, done_flag = self.parse_entry(eski)
                gun_prog[s] = {"text": metin, "done": done_flag}

        # Düzenleme modunu sıfırla
        self.edit_mode = False
        self.edit_prev_key = None

        self.listeyi_guncelle()
        self._onerileri_yenile()
        self.metin_entry.delete(0, tk.END)
        self.sablon_combo.set(SABLON_SECIMLERI[0])
        self._otomatik_metin_guncelle(force=True)
        if self.saat_multi_listbox is not None:
            self.saat_multi_listbox.selection_clear(0, tk.END)

        self.kaydet_diske()

    def secili_kaydi_bul(self):
        try:
            idx = self.listbox.curselection()[0]
        except IndexError:
            return None

        if idx < 0 or idx >= len(self.list_items_map):
            return None

        return self.list_items_map[idx]

    def sil_kayit(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        key = self.secili_kaydi_bul()
        if key is None:
            messagebox.showwarning("Uyarı", "Silmek için listeden bir ders satırı seçmelisin.")
            return

        gun, saat = key

        if messagebox.askyesno("Onay", f"{gun} - {saat} saatindeki kaydı silmek istiyor musun?"):
            if saat in program.get(gun, {}):
                del program[gun][saat]

            self.listeyi_guncelle()
            self._onerileri_yenile()

            if self.edit_mode and self.edit_prev_key == key:
                self.edit_mode = False
                self.edit_prev_key = None
                self.metin_entry.delete(0, tk.END)

            self.kaydet_diske()

    def duzenle_kayit(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        key = self.secili_kaydi_bul()
        if key is None:
            messagebox.showwarning("Uyarı", "Düzenlemek için listeden bir ders satırı seçmelisin.")
            return

        gun, saat = key
        entry = program[gun][saat]
        metin, _ = self.parse_entry(entry)

        self.gun_combo.set(gun)
        self.saat_combo.set(saat)
        if self.saat_multi_listbox is not None:
            self.saat_multi_listbox.selection_clear(0, tk.END)
            try:
                index = saatler.index(saat)
                self.saat_multi_listbox.selection_set(index)
            except ValueError:
                pass

        self.metin_entry.delete(0, tk.END)
        self.metin_entry.insert(0, metin)
        self.sablon_combo.set(SABLON_SECIMLERI[0])
        if self.kaynak_combo is not None:
            self.kaynak_combo.set("(Yayın seç)")

        self.edit_mode = True
        self.edit_prev_key = (gun, saat)

    def tamamlandi_degistir(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        key = self.secili_kaydi_bul()
        if key is None:
            messagebox.showwarning("Uyarı", "Durum değiştirmek için listeden bir ders satırı seçmelisin.")
            return

        gun, saat = key
        entry = program.get(gun, {}).get(saat)
        if entry is None:
            return

        text, done = self.parse_entry(entry)
        program[gun][saat] = {"text": text, "done": not done}

        self.listeyi_guncelle()
        self._onerileri_yenile()
        self.kaydet_diske()

    def cizelgeyi_temizle(self):
        program = self.aktif_program()
        if program is None:
            return

        bos_mu = all(len(program[g]) == 0 for g in gunler)
        if bos_mu:
            messagebox.showinfo("Bilgi", "Çizelge zaten boş.")
            return

        devam = messagebox.askyesno(
            "Onay",
            "Bu haftanın tüm programını sıfırlamak istiyor musun?\n"
            "(Bu işlem Excel dosyasını ETKİLEMEZ, sadece ekrandaki programı temizler.)"
        )
        if not devam:
            return

        for g in gunler:
            program[g] = {}
        self.edit_mode = False
        self.edit_prev_key = None
        self.metin_entry.delete(0, tk.END)
        self.sablon_combo.set(SABLON_SECIMLERI[0])
        if self.kaynak_combo is not None:
            self.kaynak_combo.set("(Yayın seç)")
        if self.saat_multi_listbox is not None:
            self.saat_multi_listbox.selection_clear(0, tk.END)
        self.listeyi_guncelle()
        self._onerileri_yenile()
        self.kaydet_diske()

    def tum_gunlere_ayni_saat_metin(self):
        program = self.aktif_program()
        if program is None:
            return

        # Bu fonksiyon tek bir saat için çalışıyor; combobox'taki saati kullanıyoruz.
        saat = self.saat_combo.get()
        metin = self.metin_entry.get().strip()

        if not saat:
            messagebox.showwarning("Uyarı", "Önce bir saat seçmelisin.")
            return

        if not metin:
            messagebox.showwarning("Uyarı", "Yazılacak metin boş olamaz.")
            return

        win = tk.Toplevel(self.root)
        win.title("Gün Seçimi")
        win.resizable(False, False)

        tk.Label(
            win,
            text=f"Bu metni hangi günlere ekleyelim?\n\nSaat: {saat}\nMetin: {metin}",
            justify="left"
        ).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

        gun_vars = {}

        for i, g in enumerate(gunler, start=1):
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(win, text=g, variable=var)
            cb.grid(row=i, column=0, columnspan=2, sticky="w", padx=15)
            gun_vars[g] = var

        def uygula():
            secilenler = [g for g, v in gun_vars.items() if v.get()]

            if not secilenler:
                messagebox.showwarning("Uyarı", "En az bir gün seçmelisin.")
                return

            devam2 = messagebox.askyesno(
                "Onay",
                "Aşağıdaki bilgi, SEÇİLEN günler için aynı saate yazılacak:\n\n"
                f"Saat : {saat}\n"
                f"Metin: {metin}\n\n"
                f"Seçilen günler: {', '.join(secilenler)}\n\n"
                "Var olan kayıtların üzerine YAZILACAKTIR.\n\n"
                "Devam etmek istiyor musun?"
            )
            if not devam2:
                return

            for gun in secilenler:
                program[gun][saat] = {"text": metin, "done": False}

            self.listeyi_guncelle()
            self.kaydet_diske()

            messagebox.showinfo(
                "Bilgi",
                f"{', '.join(secilenler)} günlerinde {saat} saatine metin yazıldı."
            )
            win.destroy()

        def iptal():
            win.destroy()

        tk.Button(win, text="Uygula", width=12, command=uygula).grid(
            row=len(gunler) + 1, column=0, padx=10, pady=10, sticky="e"
        )
        tk.Button(win, text="İptal", width=8, command=iptal).grid(
            row=len(gunler) + 1, column=1, padx=5, pady=10, sticky="w"
        )

    # -------------- HAFTA YÖNETİMİ -----------------

    def hafta_degisti(self, event=None):
        secilen = self.hafta_combo.get()
        if not secilen:
            return

        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci seçmelisin.")
            return

        ogr_data = self.ogrenciler[self.aktif_ogrenci]
        self.aktif_hafta = secilen

        if secilen not in ogr_data:
            ogr_data[secilen] = {g: {} for g in gunler}

        self.listeyi_guncelle()
        self.kaydet_diske()

    def yeni_hafta_olustur(self):
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci seçmelisin.")
            return

        ogr_data = self.ogrenciler[self.aktif_ogrenci]
        mevcut_haftalar = list(ogr_data.keys())

        numaralar = []
        for h in mevcut_haftalar:
            if h.lower().startswith("hafta"):
                parcali = h.split()
                if len(parcali) >= 2 and parcali[1].isdigit():
                    numaralar.append(int(parcali[1]))
        sonraki = (max(numaralar) + 1) if numaralar else 1
        varsayilan_isim = f"Hafta {sonraki}"

        isim = simpledialog.askstring(
            "Yeni Hafta",
            "Yeni hafta adı:",
            initialvalue=varsayilan_isim,
            parent=self.root
        )
        if not isim:
            return

        if isim in ogr_data:
            messagebox.showwarning("Uyarı", "Bu isimde bir hafta zaten var.")
            return

        ogr_data[isim] = {g: {} for g in gunler}
        self.aktif_hafta = isim

        if self.hafta_combo is not None:
            self.hafta_combo["values"] = list(ogr_data.keys())
            self.hafta_combo.set(isim)

        self.listeyi_guncelle()
        self.kaydet_diske()

    def hafta_kopyala(self):
        """Seçili haftanın programını başka bir haftaya (mevcut veya yeni) kopyalar."""
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci seçmelisin.")
            return
        ogr_data = self.ogrenciler[self.aktif_ogrenci]
        if not self.aktif_hafta or self.aktif_hafta not in ogr_data:
            messagebox.showwarning("Uyarı", "Önce kopyalanacak haftayı seçmelisin.")
            return

        mevcut_haftalar = list(ogr_data.keys())
        diger_haftalar = [h for h in mevcut_haftalar if h != self.aktif_hafta]

        win = tk.Toplevel(self.root)
        win.title("Bu haftayı kopyala")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="Hedef hafta (mevcut veya yeni):", font=self._font_label).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        hedef_combo = ttk.Combobox(win, values=diger_haftalar, state="readonly", width=22)
        hedef_combo.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self._lock_combobox_wheel(hedef_combo)
        if diger_haftalar:
            hedef_combo.set(diger_haftalar[0])
        tk.Label(win, text="Veya yeni hafta adı:", font=self._font_label).grid(row=2, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")
        yeni_entry = tk.Entry(win, width=24, font=self._font_label)
        yeni_entry.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        def kopyala():
            hedef = (yeni_entry.get() or "").strip() or (hedef_combo.get() if hedef_combo["values"] else "")
            if not hedef:
                messagebox.showwarning("Uyarı", "Hedef hafta seçin veya yeni hafta adı yazın.", parent=win)
                return
            if hedef == self.aktif_hafta:
                messagebox.showwarning("Uyarı", "Hedef, şu anki haftadan farklı olmalı.", parent=win)
                return
            kaynak = ogr_data[self.aktif_hafta]
            ogr_data[hedef] = copy.deepcopy(kaynak)
            if self.hafta_combo is not None:
                self.hafta_combo["values"] = list(ogr_data.keys())
                self.hafta_combo.set(hedef)
            self.aktif_hafta = hedef
            self.listeyi_guncelle()
            self.kaydet_diske()
            win.destroy()
            messagebox.showinfo("Tamam", f"'{self.aktif_hafta}' haftası kopyalandı.", parent=self.root)

        tk.Button(win, text="Kopyala", command=kopyala, **self._btn_opts).grid(row=4, column=0, padx=10, pady=10, sticky="e")
        tk.Button(win, text="İptal", command=win.destroy, **self._btn_opts).grid(row=4, column=1, padx=5, pady=10, sticky="w")
        win.columnconfigure(0, weight=1)

    def deneme_ekle_penceresi(self):
        """Deneme kaydı ekler (tarih, tür, puan, ders netleri)."""
        if self.aktif_ogrenci is None or self.aktif_ogrenci not in self.ogrenciler:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci seçmelisin.")
            return
        from datetime import date
        win = tk.Toplevel(self.root)
        win.title("Deneme ekle")
        win.transient(self.root)
        tk.Label(win, text="Tarih (YYYY-MM-DD):", font=self._font_label).grid(row=0, column=0, padx=10, pady=8, sticky="e")
        tarih_entry = tk.Entry(win, width=14, font=self._font_label)
        tarih_entry.insert(0, date.today().isoformat())
        tarih_entry.grid(row=0, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Ad (örn. 1. Deneme):", font=self._font_label).grid(row=1, column=0, padx=10, pady=8, sticky="e")
        ad_entry = tk.Entry(win, width=22, font=self._font_label)
        ad_entry.grid(row=1, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Puan (0-100 veya net):", font=self._font_label).grid(row=2, column=0, padx=10, pady=8, sticky="e")
        puan_entry = tk.Entry(win, width=10, font=self._font_label)
        puan_entry.grid(row=2, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Tür:", font=self._font_label).grid(row=3, column=0, padx=10, pady=8, sticky="e")
        tur_combo = ttk.Combobox(win, values=["TYT", "AYT"], state="readonly", width=10)
        tur_combo.set("TYT")
        tur_combo.grid(row=3, column=1, padx=10, pady=8, sticky="w")
        self._lock_combobox_wheel(tur_combo)
        tk.Label(win, text="AYT alanı:", font=self._font_label).grid(row=4, column=0, padx=10, pady=8, sticky="e")
        ayt_alan_combo = ttk.Combobox(win, values=["Sayısal", "EA", "Sözel"], state="readonly", width=10)
        varsayilan_alan = self._ogrenci_ayt_alani()
        ayt_alan_combo.set(varsayilan_alan if varsayilan_alan in ("Sayısal", "EA", "Sözel") else "Sayısal")
        ayt_alan_combo.grid(row=4, column=1, padx=10, pady=8, sticky="w")
        self._lock_combobox_wheel(ayt_alan_combo)
        tk.Label(win, text="Türkçe net (TYT):", font=self._font_label).grid(row=5, column=0, padx=10, pady=8, sticky="e")
        tr_entry = tk.Entry(win, width=10, font=self._font_label)
        tr_entry.insert(0, "0")
        tr_entry.grid(row=5, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Matematik net:", font=self._font_label).grid(row=6, column=0, padx=10, pady=8, sticky="e")
        mat_entry = tk.Entry(win, width=10, font=self._font_label)
        mat_entry.insert(0, "0")
        mat_entry.grid(row=6, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Fen net:", font=self._font_label).grid(row=7, column=0, padx=10, pady=8, sticky="e")
        fen_entry = tk.Entry(win, width=10, font=self._font_label)
        fen_entry.insert(0, "0")
        fen_entry.grid(row=7, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Sosyal net (TYT):", font=self._font_label).grid(row=8, column=0, padx=10, pady=8, sticky="e")
        sosyal_entry = tk.Entry(win, width=10, font=self._font_label)
        sosyal_entry.insert(0, "0")
        sosyal_entry.grid(row=8, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Edebiyat-Sos1 net (AYT):", font=self._font_label).grid(row=9, column=0, padx=10, pady=8, sticky="e")
        sos1_entry = tk.Entry(win, width=10, font=self._font_label)
        sos1_entry.insert(0, "0")
        sos1_entry.grid(row=9, column=1, padx=10, pady=8, sticky="w")
        tk.Label(win, text="Sosyal-2 net (AYT):", font=self._font_label).grid(row=10, column=0, padx=10, pady=8, sticky="e")
        sos2_entry = tk.Entry(win, width=10, font=self._font_label)
        sos2_entry.insert(0, "0")
        sos2_entry.grid(row=10, column=1, padx=10, pady=8, sticky="w")

        def tur_degisti_local(event=None):
            is_ayt = tur_combo.get() == "AYT"
            tyt_state = "disabled" if is_ayt else "normal"
            ayt_state = "normal" if is_ayt else "disabled"
            ayt_alan_combo.set(self._ogrenci_ayt_alani())
            ayt_alan_combo.configure(state="disabled")
            tr_entry.configure(state=tyt_state)
            sosyal_entry.configure(state=tyt_state)
            sos1_entry.configure(state=ayt_state)
            sos2_entry.configure(state=ayt_state)

        tur_combo.bind("<<ComboboxSelected>>", tur_degisti_local)
        tur_degisti_local()

        def ekle():
            tarih = tarih_entry.get().strip()
            ad = ad_entry.get().strip() or "Deneme"
            try:
                puan = int(puan_entry.get().strip())
            except ValueError:
                messagebox.showwarning("Uyarı", "Puan sayı olmalı.", parent=win)
                return
            try:
                mat = float((mat_entry.get() or "0").replace(",", "."))
                fen = float((fen_entry.get() or "0").replace(",", "."))
                turkce = float((tr_entry.get() or "0").replace(",", "."))
                sosyal = float((sosyal_entry.get() or "0").replace(",", "."))
                sos1 = float((sos1_entry.get() or "0").replace(",", "."))
                sos2 = float((sos2_entry.get() or "0").replace(",", "."))
            except ValueError:
                messagebox.showwarning("Uyarı", "Net alanları sayı olmalı.", parent=win)
                return
            tur = tur_combo.get() or "TYT"
            ayt_alan = self._ogrenci_ayt_alani()
            if tur == "TYT":
                sos1 = 0.0
                sos2 = 0.0
                ayt_alan = ""
            else:
                turkce = 0.0
                sosyal = 0.0
            if self.aktif_ogrenci not in self.denemeler:
                self.denemeler[self.aktif_ogrenci] = []
            self.denemeler[self.aktif_ogrenci].append({
                "tarih": tarih,
                "ad": ad,
                "puan": puan,
                "tur": tur,
                "ayt_alan": ayt_alan,
                "turkce": turkce,
                "mat": mat,
                "fen": fen,
                "sosyal": sosyal,
                "sos1": sos1,
                "sos2": sos2,
            })
            self.denemeler[self.aktif_ogrenci].sort(key=lambda x: x["tarih"])
            self.kaydet_diske()
            self._grafikleri_ciz()
            win.destroy()
            messagebox.showinfo("Tamam", "Deneme kaydedildi.", parent=self.root)

        tk.Button(win, text="Ekle", command=ekle, **self._btn_opts).grid(row=11, column=0, padx=10, pady=10, sticky="e")
        tk.Button(win, text="İptal", command=win.destroy, **self._btn_opts).grid(row=11, column=1, padx=5, pady=10, sticky="w")

    # -------------- EXCEL -----------------

    def dosya_sec(self):
        program = self.aktif_program()
        if program is None:
            return

        path = filedialog.askopenfilename(
            title="Koçluk Excel Dosyasını Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xlsm *.xltx *.xltm"), ("Tümü", "*.*")]
        )
        if not path:
            return

        self.excel_dosya_yolu = path
        self.excel_label.config(text=f"Seçili dosya: {self.excel_dosya_yolu}")
        self.update_button_states()

        for g in gunler:
            program[g] = {}

        try:
            loaded = self._excel_service.load_program_from_file(path)
            for gun in gunler:
                program[gun] = loaded.get(gun, {})
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası açılamadı:\n{e}")
            return

        self.listeyi_guncelle()
        self.kaydet_diske()
        messagebox.showinfo("Bilgi", f"Dosya okundu. Mevcut program aktarıldı.\nÖğrenci: {self.aktif_ogrenci}")

    def excel_yaz(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        if self.excel_dosya_yolu is None:
            messagebox.showwarning("Uyarı", "Önce bir Excel dosyası seçmelisin.")
            return

        try:
            self._excel_service.save_program_to_file(self.excel_dosya_yolu, program)
        except PermissionError:
            messagebox.showerror(
                "Hata",
                "Dosya kaydedilemedi.\nBüyük ihtimalle Excel'de açık.\n"
                "Lütfen Excel dosyasını kapatıp tekrar deneyin."
            )
            return
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme sırasında hata oluştu:\n{e}")
            return

        messagebox.showinfo(
            "Başarılı",
            f"Çizelge kaydedildi.\nÖğrenci: {self.aktif_ogrenci}\nDosya: {self.excel_dosya_yolu}"
        )

    # -------------- PDF -----------------

    def pdf_aktar(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        if not any(len(program[g]) > 0 for g in gunler):
            messagebox.showwarning(
                "Uyarı", "PDF oluşturmak için önce programa en az bir kayıt eklemelisin."
            )
            return

        default_name = "program.pdf"
        if self.aktif_ogrenci:
            default_name = self.aktif_ogrenci.replace(" ", "_")
            if self.aktif_hafta:
                default_name += f"_{self.aktif_hafta.replace(' ', '_')}"
            default_name += "_program.pdf"

        path = filedialog.asksaveasfilename(
            title="PDF olarak kaydet",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF Dosyası", "*.pdf")]
        )
        if not path:
            return

        try:
            self._pdf_service.build_pdf(
                path, program,
                ogrenci_adi=self.aktif_ogrenci,
                hafta_adi=self.aktif_hafta,
            )
        except Exception as e:
            messagebox.showerror("Hata", f"PDF oluşturulurken hata oluştu:\n{e}")
            return

        messagebox.showinfo("Başarılı", f"PDF oluşturuldu.\nÖğrenci: {self.aktif_ogrenci}\nDosya: {path}")

    # -------------- ÖZET / İSTATİSTİK -----------------

    def program_ozeti(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        (
            gunluk_toplam,
            haftalik_toplam,
            en_yogun_gun,
            en_bos_gun,
            gunluk_yapilan,
            haftalik_yapilan,
        ) = self.gunluk_istatistik(program)

        oran = int(round(100 * haftalik_yapilan / haftalik_toplam)) if haftalik_toplam else 0

        satirlar = [
            f"Haftalık planlanan çalışma : {haftalik_toplam} saat",
            f"Haftalık tamamlanan çalışma: {haftalik_yapilan} saat",
            f"Genel tamamlama oranı     : %{oran}",
            "",
        ]

        for gun in gunler:
            gun_prog = program.get(gun, {})
            dolu_saatler = sorted(gun_prog.keys(), key=lambda s: TIME_TO_ROW[s])
            bos_saatler = [s for s in saatler if s not in gun_prog]
            toplam_saat = gunluk_toplam.get(gun, 0)
            yapilan_saat = gunluk_yapilan.get(gun, 0)
            oran_gun = int(round(100 * yapilan_saat / toplam_saat)) if toplam_saat else 0

            yapilan_listesi = []
            yapilmayan_listesi = []
            for saat in dolu_saatler:
                entry = gun_prog[saat]
                _, done = self.parse_entry(entry)
                if done:
                    yapilan_listesi.append(saat)
                else:
                    yapilmayan_listesi.append(saat)

            satirlar.append(f"{gun}:")
            satirlar.append(
                f"  Planlanan: {toplam_saat} saat, Tamamlanan: {yapilan_saat} saat (%{oran_gun})"
            )
            satirlar.append(
                "  Tamamlanan saatler : " + (", ".join(yapilan_listesi) if yapilan_listesi else "Yok")
            )
            satirlar.append(
                "  Eksik kalan saatler: " + (", ".join(yapilmayan_listesi) if yapilmayan_listesi else "Yok")
            )
            satirlar.append(
                "  Boş saatler        : " + (", ".join(bos_saatler) if bos_saatler else "Yok")
            )
            satirlar.append("")

        if en_yogun_gun is not None:
            satirlar.append(
                f"En yoğun gün (planlanan) : {en_yogun_gun} ({gunluk_toplam[en_yogun_gun]} saat)"
            )
            satirlar.append(
                f"En boş gün (planlanan)   : {en_bos_gun} ({gunluk_toplam[en_bos_gun]} saat)"
            )

        mesaj = "\n".join(satirlar)
        messagebox.showinfo("Program Özeti", mesaj)

    def istatistik_penceresi(self, event=None):
        program = self.aktif_program()
        if program is None:
            return

        (
            gunluk_toplam,
            haftalik_toplam,
            en_yogun_gun,
            en_bos_gun,
            gunluk_yapilan,
            haftalik_yapilan,
        ) = self.gunluk_istatistik(program)

        win = tk.Toplevel(self.root)
        win.title(f"İstatistikler - {self.aktif_ogrenci if self.aktif_ogrenci else ''}")
        win.resizable(False, False)

        baslik_lbl = tk.Label(
            win,
            text=f"Haftalık Çalışma İstatistikleri - {self.aktif_ogrenci} ({self.aktif_hafta})",
            font=("Arial", 12, "bold")
        )
        baslik_lbl.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        kolonlar = ("gun", "planlanan", "yapilan")
        tree = ttk.Treeview(win, columns=kolonlar, show="headings", height=len(gunler))
        tree.heading("gun", text="Gün")
        tree.heading("planlanan", text="Planlanan (saat)")
        tree.heading("yapilan", text="Tamamlanan (saat)")
        tree.column("gun", width=150, anchor="center")
        tree.column("planlanan", width=130, anchor="center")
        tree.column("yapilan", width=150, anchor="center")

        for gun in gunler:
            tree.insert(
                "",
                "end",
                values=(
                    gun,
                    gunluk_toplam.get(gun, 0),
                    gunluk_yapilan.get(gun, 0),
                ),
            )

        tree.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

        oran = int(round(100 * haftalik_yapilan / haftalik_toplam)) if haftalik_toplam else 0
        ozet_text = (
            f"Haftalık planlanan çalışma : {haftalik_toplam} saat"
            f"\nHaftalık tamamlanan çalışma: {haftalik_yapilan} saat"
            f"\nGenel tamamlama oranı      : %{oran}"
        )
        if en_yogun_gun is not None:
            ozet_text += (
                f"\nEn yoğun gün (planlanan) : {en_yogun_gun} ({gunluk_toplam[en_yogun_gun]} saat)"
            )
        if en_bos_gun is not None:
            ozet_text += (
                f"\nEn boş gün (planlanan)   : {en_bos_gun} ({gunluk_toplam[en_bos_gun]} saat)"
            )

        ozet_lbl = tk.Label(win, text=ozet_text, justify="left")
        ozet_lbl.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

    # -------------- ŞABLON / NOT / KONU / YAYIN -----------------

    def _ogrenci_tercihini_kaydet(self):
        """Aktif öğrenci için form seçimlerini hatırlar."""
        if self.aktif_ogrenci is None:
            return
        self.ogrenci_tercihleri[self.aktif_ogrenci] = {
            "gun": self.gun_combo.get() if self.gun_combo is not None else "",
            "saat": self.saat_combo.get() if self.saat_combo is not None else "",
            "sinav": self.sinav_combo.get() if self.sinav_combo is not None else "",
            "ders": self.ders_combo.get() if self.ders_combo is not None else "",
            "konu": self.konu_combo.get() if self.konu_combo is not None else "",
            "kaynak": self.kaynak_combo.get() if self.kaynak_combo is not None else "",
        }

    def _ogrenci_tercihini_uygula(self):
        """Aktif öğrenciye ait son seçimleri forma geri uygular."""
        if self.aktif_ogrenci is None:
            return
        saved_preferences = self.ogrenci_tercihleri.get(self.aktif_ogrenci, {})
        if not saved_preferences:
            return
        if self.gun_combo is not None and saved_preferences.get("gun") in gunler:
            self.gun_combo.set(saved_preferences.get("gun"))
        if self.saat_combo is not None and saved_preferences.get("saat") in saatler:
            self.saat_combo.set(saved_preferences.get("saat"))
        if self.sinav_combo is not None and saved_preferences.get("sinav"):
            self.sinav_combo.set(saved_preferences.get("sinav"))
        self.guncelle_ders_combo()
        if self.ders_combo is not None and saved_preferences.get("ders") in (self.ders_combo["values"] or []):
            self.ders_combo.set(saved_preferences.get("ders"))
        self.guncelle_konu_combo()
        if self.konu_combo is not None and saved_preferences.get("konu") in (self.konu_combo["values"] or []):
            self.konu_combo.set(saved_preferences.get("konu"))
        if self.kaynak_combo is not None and saved_preferences.get("kaynak"):
            kaynak_values = list(self.kaynak_combo["values"] or [])
            if saved_preferences.get("kaynak") in kaynak_values:
                self.kaynak_combo.set(saved_preferences.get("kaynak"))
        self._otomatik_metin_guncelle(force=True)

    def _otomatik_metin_parcasi(self) -> str:
        """Seçimlerden otomatik metin parçası üretir."""
        tur = (self.sinav_combo.get() if self.sinav_combo is not None else "").strip()
        ders = (self.ders_combo.get() if self.ders_combo is not None else "").strip()
        konu = (self.konu_combo.get() if self.konu_combo is not None else "").strip()
        if not tur or not ders or not konu:
            return ""
        auto_text = f"{tur} {ders} - {konu}"
        kaynak = (self.kaynak_combo.get() if self.kaynak_combo is not None else "").strip()
        if kaynak and kaynak != "(Yayın seç)":
            auto_text += f" | Yayın: {kaynak}"
        return auto_text

    def _otomatik_metin_guncelle(self, force: bool = False):
        """Gerekirse metin kutusunu seçimlere göre günceller."""
        if self.metin_entry is None:
            return
        auto_text = self._otomatik_metin_parcasi()
        if not auto_text:
            return
        mevcut = self.metin_entry.get().strip()
        if force or (not mevcut) or (mevcut == self._last_auto_text):
            self.metin_entry.delete(0, tk.END)
            self.metin_entry.insert(0, auto_text)
            self._last_auto_text = auto_text

    def _onerileri_yenile(self):
        """Aktif öğrenci için en sık kullanılan metinleri öneri kutusuna yazar."""
        if self.oneri_combo is None:
            return
        suggestions = []
        if self.aktif_ogrenci and self.aktif_ogrenci in self.ogrenciler:
            usage_counts = {}
            student_weeks = self.ogrenciler.get(self.aktif_ogrenci, {})
            for week_data in student_weeks.values():
                for day_name in gunler:
                    for entry in (week_data.get(day_name, {}) or {}).values():
                        text, _ = self.parse_entry(entry)
                        text = text.strip()
                        if not text:
                            continue
                        usage_counts[text] = usage_counts.get(text, 0) + 1
            suggestions = [k for k, _ in sorted(usage_counts.items(), key=lambda item: (-item[1], item[0]))[:8]]
        self.oneri_combo["values"] = suggestions
        if suggestions:
            self.oneri_combo.set(suggestions[0])
        else:
            self.oneri_combo.set("")

    def oneri_uygula(self):
        """Seçili öneriyi metin alanına uygular."""
        if self.metin_entry is None or self.oneri_combo is None:
            return
        selected_text = (self.oneri_combo.get() or "").strip()
        if not selected_text:
            return
        self.metin_entry.delete(0, tk.END)
        self.metin_entry.insert(0, selected_text)
        self._last_auto_text = selected_text

    def _varsayilan_otoplan_onerileri(self) -> list[str]:
        """Öğrenci alanına göre hızlı plan öneri metinleri üretir."""
        alan = self._ogrenci_ayt_alani()
        if alan == "EA":
            subject_pairs = [
                ("AYT", "Matematik"),
                ("TYT", "Turkce"),
                ("TYT", "Sosyal"),
                ("TYT", "Matematik"),
            ]
        elif alan == "Sözel":
            subject_pairs = [
                ("TYT", "Turkce"),
                ("TYT", "Sosyal"),
                ("TYT", "Cografya"),
                ("TYT", "Matematik"),
            ]
        else:  # Sayısal
            subject_pairs = [
                ("AYT", "Matematik"),
                ("AYT", "Fizik"),
                ("AYT", "Kimya"),
                ("TYT", "Matematik"),
            ]
        default_suggestions: list[str] = []
        for tur, ders in subject_pairs:
            konular = KONU_VERISI.get(tur, {}).get(ders, [])
            if konular:
                default_suggestions.append(f"{tur} {ders} - {konular[0]}")
        if not default_suggestions:
            default_suggestions = ["TYT Matematik - Temel Kavramlar", "TYT Turkce - Paragraf", "Deneme analizi + yanlışlar"]
        return default_suggestions

    def otomatik_plan_ekle(self):
        """Seçili güne, minimum input ile otomatik 3 kayıt ekler."""
        program = self.aktif_program()
        if program is None:
            return
        gun = (self.gun_combo.get() if self.gun_combo is not None else "").strip()
        if not gun or gun not in gunler:
            messagebox.showwarning("Uyarı", "Önce bir gün seçmelisin.")
            return

        day_program = program.get(gun, {})
        target_hours = [hour_label for hour_label in ("09:00", "16:00", "21:00") if hour_label in saatler]
        if len(target_hours) < 3:
            target_hours = list(saatler[:3])

        selected_suggestions = []
        if self.oneri_combo is not None:
            selected_suggestions = [value for value in (self.oneri_combo["values"] or []) if str(value).strip()]
        candidate_suggestions = list(selected_suggestions) + self._varsayilan_otoplan_onerileri()
        unique_suggestions = []
        for suggestion in candidate_suggestions:
            suggestion = str(suggestion).strip()
            if suggestion and suggestion not in unique_suggestions:
                unique_suggestions.append(suggestion)
        planned_texts = unique_suggestions[:3]
        while len(planned_texts) < 3:
            planned_texts.append("Deneme analizi + yanlışlar")

        occupied_hours = [hour_label for hour_label in target_hours if hour_label in day_program]
        if occupied_hours:
            approved = messagebox.askyesno(
                "Onay",
                f"{gun} gününde şu saatler dolu: {', '.join(occupied_hours)}\nÜzerine yazılsın mı?"
            )
            if not approved:
                return

        for i, hour_label in enumerate(target_hours[:3]):
            existing_entry = day_program.get(hour_label)
            done_flag = False
            if isinstance(existing_entry, dict):
                _, done_flag = self.parse_entry(existing_entry)
            day_program[hour_label] = {"text": planned_texts[i], "done": done_flag}

        self.edit_mode = False
        self.edit_prev_key = None
        self.listeyi_guncelle()
        self._onerileri_yenile()
        self.kaydet_diske()
        messagebox.showinfo("Hazır", f"{gun} için 3 otomatik kayıt eklendi.")

    def gun_secimi_degisti(self, event=None):
        self._ogrenci_tercihini_kaydet()
        self.kaydet_diske()

    def saat_secimi_degisti(self, event=None):
        self._ogrenci_tercihini_kaydet()
        self.kaydet_diske()

    def sablon_secildi(self, event=None):
        ad = self.sablon_combo.get()
        if not ad or ad == SABLON_SECIMLERI[0]:
            return

        sablon_metin = HAZIR_METINLER.get(ad, "")
        if not sablon_metin:
            return

        mevcut = self.metin_entry.get().strip()
        yeni = (mevcut + " | " + sablon_metin) if mevcut else sablon_metin

        self.metin_entry.delete(0, tk.END)
        self.metin_entry.insert(0, yeni)
        self._last_auto_text = yeni

    def kaynak_secildi(self, event=None):
        if self.kaynak_combo is None:
            return

        secim = self.kaynak_combo.get()
        if not secim or secim == "(Yayın seç)":
            return
        self._ogrenci_tercihini_kaydet()
        self._otomatik_metin_guncelle(force=True)
        self.kaydet_diske()

    def notu_kaydet(self):
        if self.aktif_ogrenci is None:
            messagebox.showwarning("Uyarı", "Önce bir öğrenci seçmelisin.")
            return

        self.aktif_ogrenci_notunu_guncelle()
        self.kaydet_diske()
        messagebox.showinfo("Bilgi", "Öğrenci notu kaydedildi.")

    # ----- TYT / AYT - DERS - KONU SİSTEMİ -----

    def guncelle_ders_combo(self):
        tur = self.sinav_combo.get()
        if not tur:
            self.ders_combo["values"] = []
            self.ders_combo.set("")
            return
        dersler = sorted(KONU_VERISI.get(tur, {}).keys())
        self.ders_combo["values"] = dersler
        if dersler:
            aktif_t = self.ogrenci_tercihleri.get(self.aktif_ogrenci or "", {})
            tercih_ders = aktif_t.get("ders", "")
            self.ders_combo.set(tercih_ders if tercih_ders in dersler else dersler[0])
        else:
            self.ders_combo.set("")
        self.guncelle_konu_combo()

    def guncelle_konu_combo(self):
        tur = self.sinav_combo.get()
        ders = self.ders_combo.get()
        if not tur or not ders:
            self.konu_combo["values"] = []
            self.konu_combo.set("")
            return
        konular = KONU_VERISI.get(tur, {}).get(ders, [])
        self.konu_combo["values"] = konular
        if konular:
            aktif_t = self.ogrenci_tercihleri.get(self.aktif_ogrenci or "", {})
            tercih_konu = aktif_t.get("konu", "")
            self.konu_combo.set(tercih_konu if tercih_konu in konular else konular[0])
            self.konu_secildi()
        else:
            self.konu_combo.set("")

    def sinav_degisti(self, event=None):
        self.guncelle_ders_combo()
        self._ogrenci_tercihini_kaydet()
        self._otomatik_metin_guncelle(force=True)
        self.kaydet_diske()

    def ders_degisti(self, event=None):
        self.guncelle_konu_combo()
        self._ogrenci_tercihini_kaydet()
        self._otomatik_metin_guncelle(force=True)
        self.kaydet_diske()

    def konu_metnini_ekle(self):
        tur = self.sinav_combo.get().strip()
        ders = self.ders_combo.get().strip()
        konu = self.konu_combo.get().strip()

        if not tur or not ders or not konu:
            messagebox.showwarning("Uyarı", "Seviye, ders ve konu seçmelisin.")
            return

        parca = f"{tur} {ders} - {konu}"

        mevcut = self.metin_entry.get().strip()
        yeni = (mevcut + " | " + parca) if mevcut else parca

        self.metin_entry.delete(0, tk.END)
        self.metin_entry.insert(0, yeni)

    def konu_secildi(self, event=None):
        tur = self.sinav_combo.get().strip()
        ders = self.ders_combo.get().strip()
        konu = self.konu_combo.get().strip()

        if not tur or not ders or not konu:
            return

        self._ogrenci_tercihini_kaydet()
        self._otomatik_metin_guncelle(force=True)
        self.kaydet_diske()

    def konu_yonetim_penceresi(self):
        win = tk.Toplevel(self.root)
        win.title("Konu Yönetimi (TYT / AYT)")
        win.resizable(False, False)

        tk.Label(win, text="Sınav Türü:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tur_combo = ttk.Combobox(win, values=["TYT", "AYT"], state="readonly", width=10)
        tur_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tur_combo.current(0)
        self._lock_combobox_wheel(tur_combo)

        tk.Label(win, text="Dersler:").grid(row=1, column=0, padx=5, pady=5, sticky="ne")
        ders_listbox = tk.Listbox(win, width=25, height=8)
        ders_listbox.grid(row=1, column=1, rowspan=3, padx=5, pady=5, sticky="w")

        tk.Label(win, text="Yeni Ders:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        yeni_ders_entry = tk.Entry(win, width=20)
        yeni_ders_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        ders_ekle_btn = tk.Button(win, text="Ders Ekle", width=12)
        ders_ekle_btn.grid(row=1, column=4, padx=5, pady=5)

        tk.Label(win, text="Konular:").grid(row=4, column=0, padx=5, pady=5, sticky="ne")
        konu_listbox = tk.Listbox(win, width=40, height=8)
        konu_listbox.grid(row=4, column=1, columnspan=2, rowspan=3, padx=5, pady=5, sticky="w")

        tk.Label(win, text="Yeni Konu:").grid(row=4, column=3, padx=5, pady=5, sticky="e")
        yeni_konu_entry = tk.Entry(win, width=25)
        yeni_konu_entry.grid(row=4, column=4, padx=5, pady=5, sticky="w")

        konu_ekle_btn = tk.Button(win, text="Konu Ekle", width=12)
        konu_ekle_btn.grid(row=5, column=4, padx=5, pady=5)

        konu_sil_btn = tk.Button(win, text="Seçili Konuyu Sil", width=16)
        konu_sil_btn.grid(row=6, column=4, padx=5, pady=5)

        def ders_listesini_yenile():
            ders_listbox.delete(0, tk.END)
            tur = tur_combo.get()
            dersler = sorted(KONU_VERISI.get(tur, {}).keys())
            for d in dersler:
                ders_listbox.insert(tk.END, d)
            if dersler:
                ders_listbox.selection_set(0)
            konulari_yenile()

        def konulari_yenile():
            konu_listbox.delete(0, tk.END)
            tur = tur_combo.get()
            try:
                idx = ders_listbox.curselection()[0]
                ders = ders_listbox.get(idx)
            except Exception:
                return
            konular = KONU_VERISI.get(tur, {}).get(ders, [])
            for k in konular:
                konu_listbox.insert(tk.END, k)

        def tur_degisti_local(event=None):
            ders_listesini_yenile()

        def ders_listbox_degisti(event=None):
            konulari_yenile()

        def ders_ekle():
            tur = tur_combo.get()
            ad = yeni_ders_entry.get().strip()
            if not ad:
                messagebox.showwarning("Uyarı", "Ders adı boş olamaz.")
                return
            if tur not in KONU_VERISI:
                KONU_VERISI[tur] = {}
            if ad not in KONU_VERISI[tur]:
                KONU_VERISI[tur][ad] = []
                self.kaydet_diske()
                ders_listesini_yenile()
                self.guncelle_ders_combo()
            yeni_ders_entry.delete(0, tk.END)

        def konu_ekle():
            tur = tur_combo.get()
            try:
                idx = ders_listbox.curselection()[0]
                ders = ders_listbox.get(idx)
            except Exception:
                messagebox.showwarning("Uyarı", "Önce bir ders seçmelisin.")
                return

            ad = yeni_konu_entry.get().strip()
            if not ad:
                messagebox.showwarning("Uyarı", "Konu adı boş olamaz.")
                return

            if tur not in KONU_VERISI:
                KONU_VERISI[tur] = {}
            if ders not in KONU_VERISI[tur]:
                KONU_VERISI[tur][ders] = []

            if ad not in KONU_VERISI[tur][ders]:
                KONU_VERISI[tur][ders].append(ad)
                self.kaydet_diske()
                konulari_yenile()
                self.guncelle_konu_combo()

            yeni_konu_entry.delete(0, tk.END)

        def konu_sil():
            tur = tur_combo.get()
            try:
                d_idx = ders_listbox.curselection()[0]
                ders = ders_listbox.get(d_idx)
            except Exception:
                messagebox.showwarning("Uyarı", "Önce bir ders seçmelisin.")
                return
            try:
                k_idx = konu_listbox.curselection()[0]
                konu = konu_listbox.get(k_idx)
            except Exception:
                messagebox.showwarning("Uyarı", "Silmek için bir konu seçmelisin.")
                return

            if messagebox.askyesno("Onay", f"{ders} dersinden '{konu}' konusunu silmek istiyor musun?"):
                if tur in KONU_VERISI and ders in KONU_VERISI[tur]:
                    if konu in KONU_VERISI[tur][ders]:
                        KONU_VERISI[tur][ders].remove(konu)
                        self.kaydet_diske()
                        konulari_yenile()
                        self.guncelle_konu_combo()

        tur_combo.bind("<<ComboboxSelected>>", tur_degisti_local)
        ders_listbox.bind("<<ListboxSelect>>", ders_listbox_degisti)
        ders_ekle_btn.configure(command=ders_ekle)
        konu_ekle_btn.configure(command=konu_ekle)
        konu_sil_btn.configure(command=konu_sil)

        ders_listesini_yenile()

    # -------------- GUI -----------------

    def on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_mousewheel(self, event):
        w = self.root.winfo_containing(event.x_root, event.y_root)
        if w is not None:
            cls = w.winfo_class()
            # Combo/Listbox üzerinde wheel ile seçimin değişmesini engelle
            if cls in ("TCombobox", "Combobox", "Listbox", "Text", "Treeview"):
                return "break"
        delta = int(-1 * (event.delta / 120)) if event.delta else 0
        if delta:
            self.canvas.yview_scroll(delta, "units")
            return "break"
        return None

    def _wheel_sadece_kaydir(self, event):
        """Combobox/dropdown üstündeyken wheel hiçbir etki yapmasın."""
        return "break"

    def _lock_combobox_wheel(self, cb):
        """Tek bir combobox için wheel olayını tamamen kapat."""
        if cb is None:
            return
        cb.bind("<MouseWheel>", self._wheel_sadece_kaydir)
        cb.bind("<Button-4>", self._wheel_sadece_kaydir)
        cb.bind("<Button-5>", self._wheel_sadece_kaydir)

    def _setup_ui_style(self):
        """Tema ve font ayarları: okunaklı, sade arayüz. Tema config'den okunur."""
        theme_name = get_theme(_SCRIPT_DIR)
        if theme_name not in THEMES:
            theme_name = THEME_KEYS[0]
        self._current_theme = theme_name
        t = THEMES[theme_name]
        self._bg = t["bg"]
        self._fg = t["fg"]
        self._entry_bg = t["entry_bg"]
        self._accent = t["accent"]
        self._label_muted = t["label_muted"]
        self._btn_bg = t["btn_bg"]
        self._btn_active = t["btn_active"]

        self.root.configure(bg=self._bg)
        self._font_label = ("Segoe UI", 10)
        self._font_heading = ("Segoe UI", 10, "bold")
        self._pad = {"padx": 8, "pady": 6}
        self._pad_section = {"padx": 10, "pady": 10}
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background=self._bg)
        style.configure("TLabelframe", background=self._bg, font=self._font_heading)
        style.configure("TLabelframe.Label", background=self._bg, font=self._font_heading, foreground=self._fg)
        style.configure("TButton", font=self._font_label, padding=(12, 6))
        style.configure("TCombobox", font=self._font_label)
        style.configure("Treeview", font=("Consolas", 9), rowheight=22)
        style.configure("Treeview.Heading", font=self._font_heading)
        self._btn_opts = {"font": self._font_label, "cursor": "hand2", "bg": self._btn_bg, "activebackground": self._btn_active, "relief": "flat", "bd": 0, "padx": 12, "pady": 6}

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding=8)
        main_frame.pack(fill="both", expand=True)

        paned = ttk.PanedWindow(main_frame, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Sol panel: mevcut form (kaydırılabilir)
        left_panel = ttk.Frame(paned)
        self.canvas = tk.Canvas(left_panel, bg=self._bg, highlightthickness=0)
        v_scroll = ttk.Scrollbar(left_panel, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        paned.add(left_panel, weight=1)

        # Sağ panel: çalışma / haftalık / deneme istatistik grafikleri
        right_panel = ttk.Frame(paned, width=400)
        right_panel.pack_propagate(False)
        self._grafik_sec = ttk.LabelFrame(right_panel, text=" Çalışma & deneme istatistikleri ", padding=8)
        self._grafik_sec.pack(fill="both", expand=True)
        btn_frame = ttk.Frame(self._grafik_sec)
        btn_frame.pack(fill="x", pady=(0, 4))
        tk.Button(btn_frame, text="Deneme ekle", command=self.deneme_ekle_penceresi, **self._btn_opts).pack(side="left")
        self.deneme_zaman_combo = ttk.Combobox(
            btn_frame, values=["Tümü", "Son 30 gün", "Son 3 ay"], state="readonly", width=14
        )
        self.deneme_zaman_combo.set(self.deneme_zaman_filtre)
        self.deneme_zaman_combo.pack(side="right", padx=(4, 0))
        self.deneme_zaman_combo.bind("<<ComboboxSelected>>", self._deneme_zaman_filtre_degisti)
        self.deneme_tur_combo = ttk.Combobox(btn_frame, values=["Tümü", "TYT", "AYT"], state="readonly", width=10)
        self.deneme_tur_combo.set(self.deneme_tur_filtre)
        self.deneme_tur_combo.pack(side="right")
        self.deneme_tur_combo.bind("<<ComboboxSelected>>", self._deneme_tur_filtre_degisti)
        self.grafik_canvas = tk.Canvas(self._grafik_sec, bg=self._bg, highlightthickness=0)
        self.grafik_canvas.pack(fill="both", expand=True)
        self.grafik_canvas.bind("<Configure>", lambda e: self._grafikleri_ciz())
        paned.add(right_panel, weight=0)

        self.content_frame = ttk.Frame(self.canvas, padding=4)
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        self.content_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.content_frame.bind("<MouseWheel>", self.on_mousewheel)
        self.grafik_canvas.bind("<MouseWheel>", self.on_mousewheel)
        # Combobox değerlerinin wheel ile istemsiz değişmesini global engelle
        self.root.bind_class("TCombobox", "<MouseWheel>", self._wheel_sadece_kaydir)
        self.root.bind_class("TCombobox", "<Button-4>", self._wheel_sadece_kaydir)
        self.root.bind_class("TCombobox", "<Button-5>", self._wheel_sadece_kaydir)
        # Combobox açılır listesindeki listbox için de aynı davranış
        self.root.bind_class("Listbox", "<MouseWheel>", self._wheel_sadece_kaydir)
        self.root.bind_class("Listbox", "<Button-4>", self._wheel_sadece_kaydir)
        self.root.bind_class("Listbox", "<Button-5>", self._wheel_sadece_kaydir)
        # ttk Combobox popdown listesi için özel bindtag (Windows/Tk)
        self.root.bind_class("ComboboxListbox", "<MouseWheel>", self._wheel_sadece_kaydir)
        self.root.bind_class("ComboboxListbox", "<Button-4>", self._wheel_sadece_kaydir)
        self.root.bind_class("ComboboxListbox", "<Button-5>", self._wheel_sadece_kaydir)

        p, ps = self._pad, self._pad_section
        lbl_font = self._font_label
        entry_opts = {"font": lbl_font, "bg": self._entry_bg}

        # ---- Bölüm: Öğrenci & Hafta ----
        sec_ogr = ttk.LabelFrame(self.content_frame, text=" Öğrenci & Hafta ", padding=8)
        sec_ogr.grid(row=0, column=0, sticky="ew", **ps)

        tk.Label(sec_ogr, text="Öğrenci Ad Soyad:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=0, column=0, **p, sticky="e")
        self.ogrenci_entry = tk.Entry(sec_ogr, width=26, **entry_opts)
        self.ogrenci_entry.grid(row=0, column=1, **p, sticky="w")
        ogrenci_btn = tk.Button(sec_ogr, text="Öğrenci Ekle / Seç", command=self.ogrenci_ekle_veya_sec, **self._btn_opts)
        ogrenci_btn.grid(row=0, column=2, **p, sticky="w")

        tk.Label(sec_ogr, text="Öğrenci alanı:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=1, column=0, **p, sticky="e")
        self.ogrenci_alan_combo = ttk.Combobox(sec_ogr, values=["Sayısal", "EA", "Sözel"], state="readonly", width=12)
        self.ogrenci_alan_combo.grid(row=1, column=1, **p, sticky="w")
        self.ogrenci_alan_combo.set("Sayısal")
        self.ogrenci_alan_combo.bind("<<ComboboxSelected>>", self.ogrenci_alani_degisti)

        tk.Label(sec_ogr, text="Aktif Öğrenci:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=2, column=0, **p, sticky="e")
        self.ogrenci_combo = ttk.Combobox(sec_ogr, values=[], state="readonly", width=24)
        self.ogrenci_combo.grid(row=2, column=1, **p, sticky="w")
        self.ogrenci_combo.bind("<<ComboboxSelected>>", self.ogrenci_degisti)
        dosya_btn = tk.Button(sec_ogr, text="Excel Dosyası Seç", command=self.dosya_sec, **self._btn_opts)
        dosya_btn.grid(row=2, column=2, **p, sticky="w")

        tk.Label(sec_ogr, text="Hafta:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=3, column=0, **p, sticky="e")
        self.hafta_combo = ttk.Combobox(sec_ogr, values=[], state="readonly", width=14)
        self.hafta_combo.grid(row=3, column=1, **p, sticky="w")
        self.hafta_combo.bind("<<ComboboxSelected>>", self.hafta_degisti)
        yeni_hafta_btn = tk.Button(sec_ogr, text="Yeni Hafta", command=self.yeni_hafta_olustur, **self._btn_opts)
        yeni_hafta_btn.grid(row=3, column=2, **p, sticky="w")
        hafta_kopyala_btn = tk.Button(sec_ogr, text="Bu haftayı kopyala", command=self.hafta_kopyala, **self._btn_opts)
        hafta_kopyala_btn.grid(row=3, column=3, **p, sticky="w")

        self.excel_label = tk.Label(sec_ogr, text="Seçili dosya: (yok)", font=lbl_font, fg=self._label_muted, bg=self._bg)
        self.excel_label.grid(row=4, column=0, columnspan=3, **p, sticky="w")

        tk.Label(sec_ogr, text="Panel rengi:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=5, column=0, **p, sticky="e")
        self.theme_combo = ttk.Combobox(sec_ogr, values=THEME_KEYS, state="readonly", width=18)
        self.theme_combo.grid(row=5, column=1, **p, sticky="w")
        self.theme_combo.set(self._current_theme)
        self.theme_combo.bind("<<ComboboxSelected>>", self._tema_degisti)
        tk.Button(sec_ogr, text="Ayarlar", command=self.ayarlar_penceresi, **self._btn_opts).grid(row=6, column=0, **p, sticky="w")
        tk.Button(sec_ogr, text="Dışa Aktar", command=self.disa_aktar_json, **self._btn_opts).grid(row=6, column=1, **p, sticky="w")
        tk.Button(sec_ogr, text="İçe Aktar", command=self.ice_aktar_json, **self._btn_opts).grid(row=6, column=2, **p, sticky="w")
        tk.Button(sec_ogr, text="Yedekten Dön", command=self.yedekten_don, **self._btn_opts).grid(row=6, column=3, **p, sticky="w")
        tk.Button(sec_ogr, text="Veriyi Sıfırla", command=self.veriyi_sifirla, **self._btn_opts).grid(row=7, column=0, columnspan=2, **p, sticky="w")
        self._sec_frames = [sec_ogr]

        # ---- Bölüm: Çizelge girişi ----
        sec_giris = ttk.LabelFrame(self.content_frame, text=" Çizelge girişi ", padding=8)
        sec_giris.grid(row=1, column=0, sticky="ew", **ps)

        tk.Label(sec_giris, text="Gün:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=0, column=0, **p, sticky="e")
        self.gun_combo = ttk.Combobox(sec_giris, values=gunler, state="readonly")
        self.gun_combo.grid(row=0, column=1, **p, sticky="w")
        self.gun_combo.current(0)
        self.gun_combo.bind("<<ComboboxSelected>>", self.gun_secimi_degisti)

        tk.Label(sec_giris, text="Saat:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=1, column=0, **p, sticky="e")
        self.saat_combo = ttk.Combobox(sec_giris, values=saatler, state="readonly")
        self.saat_combo.grid(row=1, column=1, **p, sticky="w")
        self.saat_combo.current(0)
        self.saat_combo.bind("<<ComboboxSelected>>", self.saat_secimi_degisti)

        tk.Label(sec_giris, text="Çoklu Saat:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=1, column=2, **p, sticky="e")
        saat_multi_wrap = ttk.Frame(sec_giris)
        saat_multi_wrap.grid(row=1, column=3, **p, sticky="w")
        self.saat_multi_listbox = tk.Listbox(
            saat_multi_wrap, selectmode="extended", height=6, exportselection=False, width=12,
            font=lbl_font, bg=self._entry_bg
        )
        saat_multi_scroll = ttk.Scrollbar(saat_multi_wrap, orient="vertical", command=self.saat_multi_listbox.yview)
        self.saat_multi_listbox.configure(yscrollcommand=saat_multi_scroll.set)
        self.saat_multi_listbox.grid(row=0, column=0, sticky="nsew")
        saat_multi_scroll.grid(row=0, column=1, sticky="ns")
        saat_multi_wrap.columnconfigure(0, weight=1)
        for s in saatler:
            self.saat_multi_listbox.insert(tk.END, s)

        tk.Label(sec_giris, text="Hazır şablon:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=2, column=0, **p, sticky="e")
        self.sablon_combo = ttk.Combobox(sec_giris, values=SABLON_SECIMLERI, state="readonly", width=32)
        self.sablon_combo.grid(row=2, column=1, columnspan=3, **p, sticky="w")
        self.sablon_combo.current(0)
        self.sablon_combo.bind("<<ComboboxSelected>>", self.sablon_secildi)

        tk.Label(sec_giris, text="İçerik / Not:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=3, column=0, **p, sticky="e")
        self.metin_entry = tk.Entry(sec_giris, width=42, **entry_opts)
        self.metin_entry.grid(row=3, column=1, columnspan=3, **p, sticky="w")

        tk.Label(sec_giris, text="Hızlı öneri:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=4, column=0, **p, sticky="e")
        self.oneri_combo = ttk.Combobox(sec_giris, values=[], state="readonly", width=42)
        self.oneri_combo.grid(row=4, column=1, columnspan=2, **p, sticky="w")
        self.oneri_combo.bind("<<ComboboxSelected>>", lambda e: self.oneri_uygula())
        self.oneri_btn = tk.Button(sec_giris, text="Öneriyi Uygula", command=self.oneri_uygula, **self._btn_opts)
        self.oneri_btn.grid(row=4, column=3, **p, sticky="w")

        tk.Label(sec_giris, text="Yayın:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=5, column=0, **p, sticky="e")
        self.kaynak_combo = ttk.Combobox(
            sec_giris, values=["(Yayın seç)"] + KAYNAK_LISTESI, state="readonly", width=42
        )
        self.kaynak_combo.grid(row=5, column=1, columnspan=3, **p, sticky="w")
        self.kaynak_combo.current(0)
        self.kaynak_combo.bind("<<ComboboxSelected>>", self.kaynak_secildi)

        tk.Label(sec_giris, text="Seviye:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=6, column=0, **p, sticky="e")
        self.sinav_combo = ttk.Combobox(
            sec_giris,
            values=["TYT", "AYT", "5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf",
                    "9. Sınıf", "10. Sınıf", "11. Sınıf", "12. Sınıf"],
            state="readonly", width=12
        )
        self.sinav_combo.grid(row=6, column=1, **p, sticky="w")
        self.sinav_combo.bind("<<ComboboxSelected>>", self.sinav_degisti)
        self.sinav_combo.set("TYT")
        konu_yonet_btn = tk.Button(sec_giris, text="Konu Yönetimi", command=self.konu_yonetim_penceresi, **self._btn_opts)
        konu_yonet_btn.grid(row=6, column=2, columnspan=2, **p, sticky="w")

        tk.Label(sec_giris, text="Ders:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=7, column=0, **p, sticky="e")
        self.ders_combo = ttk.Combobox(sec_giris, values=[], state="readonly", width=20)
        self.ders_combo.grid(row=7, column=1, **p, sticky="w")
        self.ders_combo.bind("<<ComboboxSelected>>", self.ders_degisti)

        tk.Label(sec_giris, text="Konu:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=8, column=0, **p, sticky="e")
        self.konu_combo = ttk.Combobox(sec_giris, values=[], state="readonly", width=28)
        self.konu_combo.grid(row=8, column=1, columnspan=3, **p, sticky="w")
        self.konu_combo.bind("<<ComboboxSelected>>", self.konu_secildi)

        self.ekle_btn = tk.Button(sec_giris, text="Ekle / Kaydet", command=self.ekle_kayit, **self._btn_opts)
        self.ekle_btn.grid(row=9, column=0, columnspan=4, **p)

        self.dokuz_btn = tk.Button(
            sec_giris, text="Bu saati seçili günlere uygula", command=self.tum_gunlere_ayni_saat_metin, **self._btn_opts
        )
        self.dokuz_btn.grid(row=10, column=0, columnspan=4, **p)
        self.otoplan_btn = tk.Button(
            sec_giris, text="Seçili güne otomatik 3 kayıt", command=self.otomatik_plan_ekle, **self._btn_opts
        )
        self.otoplan_btn.grid(row=11, column=0, columnspan=4, **p)

        # ---- Bölüm: Kayıt listesi & Dışa aktar ----
        sec_liste = ttk.LabelFrame(self.content_frame, text=" Kayıt listesi & Dışa aktar ", padding=8)
        sec_liste.grid(row=2, column=0, sticky="ew", **ps)

        self.liste_filtre_var = tk.StringVar(value="")
        self.sadece_yapilmayan_var = tk.BooleanVar(value=False)
        tk.Label(sec_liste, text="Listede ara:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=0, column=0, **p, sticky="e")
        self.liste_filtre_entry = tk.Entry(sec_liste, width=28, textvariable=self.liste_filtre_var, **entry_opts)
        self.liste_filtre_entry.grid(row=0, column=1, **p, sticky="w")
        self.liste_filtre_entry.bind("<KeyRelease>", self.liste_filtresi_degisti)
        ttk.Checkbutton(
            sec_liste,
            text="Sadece yapılmayanlar",
            variable=self.sadece_yapilmayan_var,
            command=self.liste_filtresi_degisti,
        ).grid(row=0, column=2, **p, sticky="w")
        tk.Button(sec_liste, text="Filtreyi temizle", command=self.liste_filtresini_temizle, **self._btn_opts).grid(row=0, column=3, **p, sticky="w")
        self.liste_sayac_label = tk.Label(sec_liste, text="Gösterilen: 0 / 0", font=lbl_font, bg=self._bg, fg=self._label_muted)
        self.liste_sayac_label.grid(row=1, column=0, columnspan=4, **p, sticky="w")

        liste_wrap = ttk.Frame(sec_liste)
        liste_wrap.grid(row=2, column=0, columnspan=4, **p, sticky="ew")
        self.listbox = tk.Listbox(
            liste_wrap,
            width=72,
            height=8,
            selectmode="browse",  # tekli seçim
            font=lbl_font,
            bg=self._entry_bg,
            selectbackground=self._accent,
        )
        liste_scroll = ttk.Scrollbar(liste_wrap, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=liste_scroll.set)
        self.listbox.grid(row=0, column=0, sticky="ew")
        liste_scroll.grid(row=0, column=1, sticky="ns")
        liste_wrap.columnconfigure(0, weight=1)
        self.listbox.bind("<Double-Button-1>", self.duzenle_kayit)

        self.sil_btn = tk.Button(sec_liste, text="Seçileni Sil", command=self.sil_kayit, **self._btn_opts)
        self.sil_btn.grid(row=3, column=0, **p, sticky="e")
        self.duzenle_btn = tk.Button(sec_liste, text="Seçileni Düzenle", command=self.duzenle_kayit, **self._btn_opts)
        self.duzenle_btn.grid(row=3, column=1, **p, sticky="w")
        self.yapildi_btn = tk.Button(sec_liste, text="Yapıldı / Yapılmadı", command=self.tamamlandi_degistir, **self._btn_opts)
        self.yapildi_btn.grid(row=3, column=2, **p, sticky="w")
        self.temizle_btn = tk.Button(sec_liste, text="Çizelgeyi Temizle", command=self.cizelgeyi_temizle, **self._btn_opts)
        self.temizle_btn.grid(row=3, column=3, **p, sticky="w")

        self.excel_btn = tk.Button(sec_liste, text="Excel'e Yaz", command=self.excel_yaz, **self._btn_opts)
        self.excel_btn.grid(row=4, column=0, **p, sticky="e")
        self.pdf_btn = tk.Button(sec_liste, text="PDF'ye Aktar", command=self.pdf_aktar, **self._btn_opts)
        self.pdf_btn.grid(row=4, column=1, **p, sticky="w")
        self.ozet_btn = tk.Button(sec_liste, text="Boş Saatler + Günlük Toplam", command=self.program_ozeti, **self._btn_opts)
        self.ozet_btn.grid(row=4, column=2, **p, sticky="w")
        self.istatistik_btn = tk.Button(sec_liste, text="İstatistik Ekranı", command=self.istatistik_penceresi, **self._btn_opts)
        self.istatistik_btn.grid(row=4, column=3, **p, sticky="w")

        # ---- Bölüm: Haftalık tablo ----
        sec_tablo = ttk.LabelFrame(self.content_frame, text=" Haftalık tablo özeti ", padding=8)
        sec_tablo.grid(row=3, column=0, sticky="ew", **ps)

        kolonlar = ["Saat"] + gunler
        self.tablo_tree = ttk.Treeview(sec_tablo, columns=kolonlar, show="headings", height=len(saatler) + 1)
        for col in kolonlar:
            self.tablo_tree.heading(col, text=col)
            if col == "Saat":
                self.tablo_tree.column(col, width=70, anchor="center")
            else:
                self.tablo_tree.column(col, width=140, anchor="center")
        self.tablo_tree.grid(row=0, column=0, **p, sticky="ew")

        # ---- Bölüm: Öğrenci notları ----
        sec_not = ttk.LabelFrame(self.content_frame, text=" Öğrenci notları ", padding=8)
        sec_not.grid(row=4, column=0, sticky="ew", **ps)

        tk.Label(sec_not, text="Notlar:", font=lbl_font, bg=self._bg, fg=self._fg).grid(row=0, column=0, **p, sticky="nw")
        self.ogrenci_not_text = tk.Text(sec_not, width=62, height=4, font=lbl_font, bg=self._entry_bg, wrap="word")
        self.ogrenci_not_text.grid(row=0, column=1, **p, sticky="w")
        self.not_kaydet_btn = tk.Button(sec_not, text="Notu Kaydet", command=self.notu_kaydet, **self._btn_opts)
        self.not_kaydet_btn.grid(row=0, column=2, **p, sticky="nw")

        self._sec_frames.extend([sec_giris, sec_liste, sec_tablo, sec_not])
        self.content_frame.columnconfigure(0, weight=1)

        # Tüm combobox'larda wheel ile değer/panel hareketini engelle
        for cb in [
            self.ogrenci_combo, self.ogrenci_alan_combo, self.hafta_combo, self.theme_combo, self.gun_combo, self.saat_combo,
            self.sablon_combo, self.kaynak_combo, self.oneri_combo, self.sinav_combo, self.ders_combo, self.konu_combo,
            self.deneme_tur_combo, self.deneme_zaman_combo,
        ]:
            self._lock_combobox_wheel(cb)

    def ayarlar_penceresi(self):
        """Font yolu, Excel sayfa adı ve tema ayarlarını düzenle."""
        st = load_settings(_SCRIPT_DIR)
        win = tk.Toplevel(self.root)
        win.title("Ayarlar")
        win.transient(self.root)

        tk.Label(win, text="PDF font dosyası:", font=self._font_label).grid(row=0, column=0, padx=10, pady=8, sticky="e")
        font_entry = tk.Entry(win, width=45, font=self._font_label)
        font_entry.insert(0, st.get("pdf_font_path", ""))
        font_entry.grid(row=0, column=1, padx=10, pady=8, sticky="w")

        tk.Label(win, text="Excel sayfa adı:", font=self._font_label).grid(row=1, column=0, padx=10, pady=8, sticky="e")
        sheet_entry = tk.Entry(win, width=30, font=self._font_label)
        sheet_entry.insert(0, st.get("excel_sheet_name", ""))
        sheet_entry.grid(row=1, column=1, padx=10, pady=8, sticky="w")

        tk.Label(win, text="Panel rengi (tema):", font=self._font_label).grid(row=2, column=0, padx=10, pady=8, sticky="e")
        tema_combo = ttk.Combobox(win, values=THEME_KEYS, state="readonly", width=22)
        tema_combo.set(st.get("theme", THEME_KEYS[0]))
        tema_combo.grid(row=2, column=1, padx=10, pady=8, sticky="w")
        self._lock_combobox_wheel(tema_combo)

        def kaydet():
            new_st = {
                "pdf_font_path": font_entry.get().strip() or st.get("pdf_font_path", ""),
                "excel_sheet_name": sheet_entry.get().strip() or st.get("excel_sheet_name", ""),
                "theme": tema_combo.get() if tema_combo.get() in THEMES else THEME_KEYS[0],
            }
            save_settings(_SCRIPT_DIR, new_st)
            self._excel_service.sheet_name = new_st["excel_sheet_name"]
            self._pdf_service = PdfService(font_path=new_st["pdf_font_path"] or None)
            if new_st["theme"] != self._current_theme:
                self.theme_combo.set(new_st["theme"])
                self._apply_theme(new_st["theme"])
            win.destroy()
            messagebox.showinfo("Ayarlar", "Ayarlar kaydedildi.", parent=self.root)

        tk.Button(win, text="Kaydet", command=kaydet, **self._btn_opts).grid(row=3, column=0, padx=10, pady=10, sticky="e")
        tk.Button(win, text="İptal", command=win.destroy, **self._btn_opts).grid(row=3, column=1, padx=5, pady=10, sticky="w")

    def _tema_degisti(self, event=None):
        """Panel rengi seçildi; ayarı kaydet ve arayüzü güncelle."""
        sel = self.theme_combo.get()
        if not sel or sel == self._current_theme or sel not in THEMES:
            return
        try:
            st = load_settings(_SCRIPT_DIR)
            st["theme"] = sel
            save_settings(_SCRIPT_DIR, st)
        except Exception:
            pass
        self._apply_theme(sel)

    def _apply_theme(self, theme_name: str):
        """Tema renklerini uygula (root, canvas, style ve tüm ilgili widget'lar)."""
        if theme_name not in THEMES:
            return
        self._current_theme = theme_name
        t = THEMES[theme_name]
        self._bg = t["bg"]
        self._fg = t["fg"]
        self._entry_bg = t["entry_bg"]
        self._accent = t["accent"]
        self._label_muted = t["label_muted"]
        self._btn_bg = t["btn_bg"]
        self._btn_active = t["btn_active"]
        self._btn_opts = {"font": self._font_label, "cursor": "hand2", "bg": self._btn_bg, "activebackground": self._btn_active, "relief": "flat", "bd": 0, "padx": 12, "pady": 6}

        self.root.configure(bg=self._bg)
        self.canvas.configure(bg=self._bg)
        try:
            self.grafik_canvas.configure(bg=self._bg)
            self._grafikleri_ciz()
        except Exception:
            pass
        style = ttk.Style()
        style.configure("TFrame", background=self._bg)
        style.configure("TLabelframe", background=self._bg)
        style.configure("TLabelframe.Label", background=self._bg, foreground=self._fg)

        self.excel_label.configure(bg=self._bg, fg=self._label_muted)
        self.ogrenci_entry.configure(bg=self._entry_bg, fg=self._fg)
        self.metin_entry.configure(bg=self._entry_bg, fg=self._fg)
        self.saat_multi_listbox.configure(bg=self._entry_bg, fg=self._fg, selectbackground=self._accent)
        self.listbox.configure(bg=self._entry_bg, fg=self._fg, selectbackground=self._accent)
        self.ogrenci_not_text.configure(bg=self._entry_bg, fg=self._fg, insertbackground=self._fg, selectbackground=self._accent)

        for btn in (self.ekle_btn, self.dokuz_btn, self.sil_btn, self.duzenle_btn, self.yapildi_btn, self.temizle_btn,
                    self.excel_btn, self.pdf_btn, self.ozet_btn, self.istatistik_btn, self.not_kaydet_btn, self.oneri_btn, self.otoplan_btn):
            if btn is not None:
                btn.configure(bg=self._btn_bg, fg=self._fg, activebackground=self._btn_active, activeforeground=self._fg)

        def _recurse(w):
            try:
                c = w.winfo_class()
                if c == "Label":
                    w.configure(bg=self._bg, fg=self._fg)
                elif c == "Button":
                    w.configure(bg=self._btn_bg, fg=self._fg, activebackground=self._btn_active, activeforeground=self._fg)
                elif c == "Entry":
                    w.configure(bg=self._entry_bg, fg=self._fg)
                elif c == "Listbox":
                    w.configure(bg=self._entry_bg, fg=self._fg, selectbackground=self._accent)
                elif c == "Text":
                    w.configure(bg=self._entry_bg, fg=self._fg, insertbackground=self._fg, selectbackground=self._accent)
            except Exception:
                pass
            for ch in w.winfo_children():
                _recurse(ch)

        for sec in self._sec_frames:
            _recurse(sec)

    def bind_shortcuts(self):
        self.root.bind("<Return>", self.ekle_kayit)
        self.root.bind("<Delete>", self.sil_kayit)
        self.root.bind("<Control-s>", self.excel_yaz)
        self.root.bind("<Control-Shift-S>", lambda e: self.kaydet_diske())  # JSON kaydet
        self.root.bind("<Control-e>", lambda e: self.disa_aktar_json())
        self.root.bind("<Control-Shift-I>", lambda e: self.ice_aktar_json())
        self.root.bind("<Control-p>", self.pdf_aktar)
        self.root.bind("<Control-o>", self.program_ozeti)
        self.root.bind("<Control-i>", self.istatistik_penceresi)
        self.root.bind("<Control-f>", self._liste_filtresine_odaklan)
        self.root.bind("<Control-Shift-A>", lambda e: self.otomatik_plan_ekle())
        self.root.bind("<Escape>", self._iptal_duzenleme)

    def _iptal_duzenleme(self, event=None):
        """Düzenleme modundan çık, metin kutusunu temizle."""
        if self.edit_mode:
            self.edit_mode = False
            self.edit_prev_key = None
            self.metin_entry.delete(0, tk.END)
            self.sablon_combo.set(SABLON_SECIMLERI[0])
            if self.kaynak_combo is not None:
                self.kaynak_combo.set("(Yayın seç)")

    def _liste_filtresine_odaklan(self, event=None):
        """Ctrl+F ile kayıt listesi arama kutusuna odaklan."""
        if self.liste_filtre_entry is not None:
            self.liste_filtre_entry.focus_set()
            self.liste_filtre_entry.select_range(0, tk.END)
        return "break"

    # -------------- BUTON STATE YÖNETİMİ -----------------

    def update_button_states(self):
        has_student = self.aktif_ogrenci is not None and self.aktif_ogrenci in self.ogrenciler
        has_excel = self.excel_dosya_yolu is not None

        state = "normal" if has_student else "disabled"

        for btn in [
            self.ekle_btn,
            self.dokuz_btn,
            self.sil_btn,
            self.duzenle_btn,
            self.temizle_btn,
            self.pdf_btn,
            self.ozet_btn,
            self.istatistik_btn,
            self.not_kaydet_btn,
            self.yapildi_btn,
            self.oneri_btn,
            self.otoplan_btn,
        ]:
            if btn is not None:
                btn.config(state=state)

        if self.excel_btn is not None:
            if has_student and has_excel:
                self.excel_btn.config(state="normal")
            else:
                self.excel_btn.config(state="disabled")

    # -------------- KAPAT -----------------

    def uygulamayi_kapat(self):
        if self.aktif_ogrenci is not None:
            self.aktif_ogrenci_notunu_guncelle()
            self.kaydet_diske()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CoachingApp(root)
    root.mainloop()
