# Geliştirme Önerileri — Haftalık Koçluk Çizelgesi

Uygulamayı geliştirmek istersen aşağıdaki fikirler işini kolaylaştırabilir. Öncelik ve zorluk hissine göre seçebilirsin.

---

## 1. Kullanıcı deneyimi (UX)

- **Hafta kopyalama:** Bir haftanın çizelgesini “Hafta 2’ye kopyala” gibi başka bir haftaya tek tıkla kopyalama. Özellikle benzer haftaları doldururken zaman kazandırır.
- **Şablon setleri:** “TYT odaklı hafta”, “Sınav haftası” gibi hazır şablon setleri; seçilince tüm gün/saatlere belirli etkinlikler otomatik dağıtılır.
- **Sürükle-bırak:** Tabloda veya listede satır/saat sırasını sürükle-bırak ile değiştirme.
- **Toplu işlemler:** “Bu haftanın tüm ‘Yapıldı’ işaretlerini kaldır”, “Seçili günleri temizle” gibi toplu aksiyonlar.
- **Klavye odaklı kullanım:** Tab ile alanlar arasında akıcı geçiş, Enter ile “Ekle”, kısayol bilgisini bir “Yardım” veya status bar’da gösterme.

---

## 2. Önizleme ve raporlama

- **PDF önizleme:** “PDF’ye Aktar” öncesi küçük bir pencerede veya sekmede metin tabanlı veya basit bir önizleme.
- **Excel önizleme:** Seçili haftanın Excel’e nasıl yazılacağını (hangi hücrelere ne gidecek) kısa bir özetle gösterme.
- **Haftalık/aylık rapor:** Haftalık tamamlanma oranı, en yoğun günler, ders dağılımı gibi basit grafikler veya tablolar (matplotlib veya sadece metin raporu).

---

## 3. Veri ve güvenlik

- **Otomatik yedekleme:** Belirli aralıklarla (örn. 5 dakikada bir) `program_kayitlari.json` için otomatik yedek (örn. `program_kayitlari_backup_20250314.json`).
- **Export/Import:** Tüm veriyi tek JSON/Excel dosyasına dışa aktarma ve başka bir bilgisayarda içe aktarma (öğrenci + hafta + notlar).
- **Sürüm uyumluluğu:** JSON formatı değişirse eski kayıtları okuyup yeni formata çeviren bir “migrate” adımı.

---

## 4. Ayarlar ve kişiselleştirme

- **Ayarlar penceresi:** Font yolu, varsayılan sayfa adı, tema (açık/koyu), dil gibi ayarların tek bir pencereden yapılması (şu anki `config/ayarlar.json` ile uyumlu).
- **Tema seçimi:** Açık/koyu tema veya birkaç renk paleti (şu an tek tema var).
- **Çoklu dil:** Arayüz metinlerini Türkçe/İngilizce seçebilme (sabitleri dosyaya taşıyıp dil dosyasıyla değiştirme).

---

## 5. Teknik iyileştirmeler

- **Unit testler:** `storage`, `excel_service`, `pdf_service` ve çizelge mantığı (örn. `aktif_program`, `parse_entry`) için pytest ile testler; yeni özellik eklerken geri dönüşümleri önler.
- **Loglama:** Hata ve önemli aksiyonların `logging` ile dosyaya yazılması; kullanıcı “bir şey olmadı” dediğinde loglara bakılabilir.
- **Type hints:** Tüm modüllerde tutarlı type hint kullanımı (özellikle Python 3.8 uyumu için `Optional`, `Union`, `Tuple`).
- **Konfigürasyon:** Tüm sabitlerin (font, sayfa adı, renkler) `config` veya `constants` üzerinden tek yerden gelmesi; kod içinde dağınık “sihirli sayı” kalmaması.

---

## 6. Yeni özellik fikirleri

- **Hatırlatıcılar:** Belirli bir gün/saat için “Bu aktiviteyi yap” tarzı basit hatırlatıcı (sistem bildirimi veya uygulama içi).
- **Hedef takibi:** “Bu hafta en az X saat matematik” gibi hedef tanımlayıp gerçekleşenle karşılaştırma.
- **Takvim görünümü:** Haftalık tabloyu takvim benzeri bir görünümle gösterme (opsiyonel görünüm).
- **Çoklu öğrenci karşılaştırma:** İki öğrencinin aynı haftasını yan yana görme (karşılaştırma penceresi).

---

Özetle: Önce **hafta kopyalama**, **otomatik yedekleme** ve **ayarlar penceresi** ile başlamak hem kullanıcıya hem de bakımına fayda sağlar. İstersen bir sonraki adımda bu maddelerden birini birlikte tasarlayıp koda dökebiliriz.
