# Excel Organiser – Refactor Özeti

Uygulama modüler yapıya taşındı; mevcut davranış ve veri formatı korunuyor.

## Yeni yapı

```
excel_organiser/
├── main.py              # Giriş noktası (DPI + CoachingApp)
├── excel_organizer.py   # Ana UI (konu verisi artık data/ modülünde)
├── constants.py         # Gün/saat, şablonlar, yayın listesi, dosya adları
├── data/
│   ├── __init__.py
│   └── konu_verisi.py   # TYT/AYT ve sınıf bazlı konu verisi (varsayılan)
├── config/
│   ├── __init__.py
│   └── settings.py     # Ayarlar: PDF font yolu, Excel sayfa adı (config/ayarlar.json)
├── storage/
│   ├── __init__.py
│   └── json_storage.py # JSON yükleme/kaydetme, program normalize
└── services/
    ├── __init__.py
    ├── excel_service.py # Excel grid okuma/yazma (sayfa adı ayarlardan)
    └── pdf_service.py   # PDF dışa aktarma (font yolu ayarlardan)
```

## Ne değişti?

- **constants.py**: `DAY_TO_COL`, `TIME_TO_ROW`, `GUNLER`, `SAATLER`, `HAZIR_METINLER`, `KAYNAK_LISTESI`, `SAYFA_ADI`, `SINAV_SECIMLERI` burada.
- **storage.JsonStorage**: `load()` / `save()`, eski program yapısını "Hafta 1" ve `{text, done}` formatına normalize ediyor.
- **ExcelService**: `load_program_from_file()`, `save_program_to_file()` – sabit grid (satır 3–17, sütun 2–9).
- **PdfService**: `build_pdf()` – başlık, özet tablo, günlük özet, çizelge tablosu.
- **excel_organizer.py**: DPI ayarı, import’lar, `_yukle_ve_uygula()`, `kaydet_diske()`, `dosya_sec()`, `excel_yaz()`, `pdf_aktar()` yeni modülleri kullanacak şekilde güncellendi. TYT/AYT konu verisi ve tüm UI aynı dosyada kalmaya devam ediyor.

## Çalıştırma

```bash
cd c:\Users\mert\PycharmProjects\excel_organiser
python main.py
# veya: python excel_organizer.py
```

Mevcut `program_kayitlari.json` aynen kullanılır; konu verisi JSON’da yoksa dosyadaki `KONU_VERISI` varsayılan olarak yazılır.

## Sonraki adımlar (isteğe bağlı)

- Konu verisini `data/konu_verisi_default.py` gibi ayrı bir modüle taşımak.
- UI’ı panellere bölmek (öğrenci, hafta, çizelge, konu yönetimi).
- Ayarlar (ör. PDF font yolu, sayfa adı) için `config/settings.json` kullanmak.
