# Bitki Numuneleme İstasyonları Platformu 🌿

Arazi numuneleme istasyonlarını kaydetmek, interaktif haritada göstermek ve
tarihsel olarak izlemek için geliştirilmiş R Shiny uygulaması.

## Klasör yapısı

```
bitki-haritasi/
├── app.R
├── manifest.json            # Connect Cloud için (writeManifest ile üretilir)
└── data/
    ├── istasyonlar.xlsx      # Başlangıç veri seti (7 sütun)
    └── history/              # Sürüm anlık görüntüleri (.xlsx) — otomatik oluşur
```

## Veri şeması (8 sütun)

`ID · ID2 · Tarih · Tür · Bilimsel · Türkçe · Enlem (DD - N) · Boylam (DD - E)`

- **ID**: sıralı kayıt numarası (otomatik artar). **ID2**: istasyon etiketi (örn. BN-1 veya düz numara).

- Tarih hem yıl (2024) hem tam tarih (2026-05-22) olabilir; yeni kayıtlarda o günün tarihi otomatik yazılır.
- Koordinatlar ondalık derece (WGS84); ondalık ayraç `.` veya `,` kabul edilir.

## Özellikler

- **Tür filtresi** ve **açılır/kapanır tür lejantı** (haritayı kaplamaması için başlığa tıklayınca kapanır).
- **Konum:** GPS düğmesi ile mevcut konum bulunur; haritaya dokunarak da nokta seçilir. Enlem/boylam panelde görünür.
- **Konuma kayıt ekleme:** haritadaki **+** düğmesi veya yan paneldeki düğme ile. Enlem/boylam ve tarih otomatik; ID otomatik ilerler. Tür adı yazıldıkça öneri gelir, **Türkçe↔Bilimsel çift yönlü** otomatik dolar; yeni tür manuel eklenebilir.
- **Silme:** Veri Tablosu'nda satır seçip "Seçili kaydı sil".
- **Sürüm geçmişi:** her ekleme/silme/yükleme ayrı bir Excel olarak `data/history/` altına yazılır; **son 20 sürüm** saklanır. İstenen sürüme dönülebilir veya indirilebilir; "Tüm geçmiş" tek çalışma kitabı olarak indirilebilir.
- **Toplu Excel yükleme / şablon indirme / dışa aktarma.**

## Gerekli R paketleri

```r
install.packages(c(
  "rsconnect","shiny","bslib","bsicons","leaflet","leaflet.extras",
  "htmlwidgets","readxl","writexl","DT","dplyr","shinyWidgets"))
```

## Posit Connect Cloud'a yükleme  (https://connect.posit.cloud/cm-yesilkanat)

1. Uygulama klasöründe `rsconnect::writeManifest()` çalıştırıp `manifest.json` üretin.
2. Klasörü **public** bir GitHub deposuna gönderin (`app.R`, `manifest.json`, `data/`).
3. Connect Cloud → **Publish → Shiny** → depoyu seçin → primary dosya `app.R` → **Deploy**.

> **Kalıcılık notu:** Sürüm Excelleri çalışan örnekte ve sunucu yeniden başlayana kadar
> `data/history/` altında saklanır. Connect Cloud ücretsiz katmanında dosya sistemi
> yeniden başlatmada sıfırlanabilir; kalıcı yedek için uygulamadaki **"Tüm geçmiş (.xlsx)"**
> indirme düğmesini kullanın.

---
PhysIQ Dynamic · R Shiny + leaflet
