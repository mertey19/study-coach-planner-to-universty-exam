# Windows’ta .exe yapma ve diğer bilgisayarlarda çalıştırma

## Önemli notlar

- **.exe’yi Windows’ta üretin** → başka Windows bilgisayarlarda çalışır.
- **Mac/Linux’ta üretilen .exe olmaz**; her işletim sistemi için ayrı paket gerekir.
- Hedef PC’de **Python kurulu olması gerekmez** (tek dosya exe ile).
- `program_kayitlari.json`, `config/`, yedekler **exe’nin yanındaki klasörde** oluşur (bu projede buna göre ayarlandı).

---

## 1) Geliştirme ortamı

```powershell
cd C:\Users\...\excel_organiser
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
python main.py
```

---

## 2) Tek .exe (kolay dağıtım)

Proje kökünde:

```powershell
.\.venv\Scripts\activate
pyinstaller --noconfirm --clean --onefile --windowed --name "KoclukCizelgesi" main.py
```

Çıktı: `dist\KoclukCizelgesi.exe`

- **`--windowed`**: Konsol penceresi açılmaz (sadece arayüz).
- İlk çalıştırmada bazı antivirüsler uyarı verebilir; yaygın PyInstaller durumu.

### ReportLab / font

PDF için `config\ayarlar.json` veya Ayarlar’dan font yolu verilebilir. Varsayılan `C:/Windows/Fonts/arial.ttf` hedef PC’de yoksa PDF Türkçe bozulabilir; gerekirse font yolunu o PC’ye göre ayarlayın.

---

## 3) Klasör halinde dağıtım (bazen daha az sorun)

Tek dosya yerine klasör:

```powershell
pyinstaller --noconfirm --clean --onedir --windowed --name "KoclukCizelgesi" main.py
```

Çıktı: `dist\KoclukCizelgesi\` — tüm klasörü zipleyip kopyalayın.

---

## 4) Başka PC’de kullanım

1. `KoclukCizelgesi.exe` dosyasını (veya `dist\KoclukCizelgesi` klasörünü) istediğiniz yere koyun.
2. İlk çalıştırmada aynı klasörde `program_kayitlari.json` ve `config` oluşur.
3. Excel dosyası yolu kullanıcıya özeldir; her PC’de dosyayı tekrar seçmek gerekebilir.

---

## Sorun giderme

| Sorun | Öneri |
|--------|--------|
| Exe açılıp kapanıyor | Konsollu deneyin: `pyinstaller ...` satırında `--windowed` kaldırıp tekrar derleyin; hata mesajını görürsünüz. |
| Modül bulunamadı | `pyinstaller` komutuna `--hidden-import=openpyxl` `--hidden-import=reportlab` ekleyin. |
| Veri kayboldu | Veriler exe’nin bulunduğu klasörde; exe’yi taşıdıysanız json’u da taşıyın. |

---

## Geliştirme fikirleri (kısa)

- Deneme kayıtlarına **ders bazlı AYT alanları** (Sayısal/Sözel/EA) ve grafikte ayrı renkler.
- **İçe/dışa aktar**: tüm `program_kayitlari.json` yedekleme sihirbazı.
- **Otomatik güncelleme** veya en azından sürüm numarası + “Yenilikler” penceresi.
- **pytest** ile `JsonStorage`, `ExcelService` birim testleri.

Detaylı liste için `GELISTIRME_ONERILERI.md` dosyasına bakın.
