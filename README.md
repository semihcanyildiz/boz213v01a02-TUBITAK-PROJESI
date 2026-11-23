# boz213v01a02 — Dijital Sunularda WCAG 2.2 Otomatik Denetimi ve ÜYZ Destekli İyileştirme Aracı

[![Tanıtım ve Demo Videosu](https://img.youtube.com/vi/DrqJyF6NnsY/0.jpg)](https://youtu.be/DrqJyF6NnsY)

> 🎥 **Tanıtım ve Demo Videosunu İzlemek İçin Görsele veya Bağlantıya Tıklayın:** [https://youtu.be/DrqJyF6NnsY](https://youtu.be/DrqJyF6NnsY)

---

## 📌 Proje Hakkında

Bu depo, **TÜBİTAK 2209-A Üniversite Öğrencileri Araştırma Projeleri Destekleme Programı (2026 Yılı 1. Dönem)** kapsamında hazırlanan **"Dijital Sunularda Seçili WCAG 2.2 Başarı Ölçütlerinin Otomatik Denetimi ve Üretken Yapay Zekâ Destekli İyileştirilmesi: Bir Tasarım ve Geliştirme Araştırması"** başlıklı araştırma projesinin dokümantasyonunu ve yazılım geliştirme sürecini içermektedir.

Yükseköğretimde ve açık öğrenme ortamlarında yaygın olarak kullanılan `.pptx` formatındaki ders sunularının erişilebilirlik standartlarına uygun olmaması, başta görme engelli ve az gören öğrenciler olmak üzere özel gereksinimli bireyler için ciddi bilgiye erişim engelleri oluşturmaktadır. Mevcut sunu yazılımlarının yerleşik denetleyicileri yalnızca teknik eksiklikleri (örneğin alt metin kutusunun boş olmasını) tespit edebilmekte; görsel-içerik tutarlılığını anlamsal düzeyde denetleyememekte ve bağlama duyarlı düzeltme önerileri sunamamaktadır.

Bu proje ile; sunu dosyalarını **WCAG 2.2** standartları doğrultusunda hem kural tabanlı hem de anlamsal olarak denetleyen, tespit edilen ihlaller için **Üretken Yapay Zekâ (ÜYZ)** destekli düzeltme önerileri üreten ve **İnsan-Döngüde (Human-in-the-Loop)** yaklaşımıyla kullanıcı onayının ardından dosyayı otomatik iyileştiren Python tabanlı açık kaynaklı bir yazılım geliştirilmektedir.

---

## 🎯 Denetlenen WCAG 2.2 Başarı Ölçütleri

Sistem, `.pptx` (OOXML) yapısını ayrıştırarak aşağıdaki dört temel ölçüt üzerinden otomatik denetim gerçekleştirmektedir:

1. **1.1.1 Metin Dışı İçerik (Alt Metin — A Düzeyi):** Görsellerin alternatif metin (`descr`) alanlarının boş olup olmadığını, dosya adı (`resim1.jpg` vb.) içerip içermediğini kontrol eder; çok modlu ÜYZ modeli ile görselin ders kazanımına uygunluğunu analiz ederek pedagojik alt metin önerir.
2. **2.4.2 / 2.4.6 Slayt Başlığı (A / AA Düzeyi):** Her slaytta başlık yer tutucusunun (`TITLE` / `CENTER_TITLE`) varlığını, benzersizliğini ve slayt gövde metniyle anlamsal tutarlılığını denetler.
3. **3.1.1 Sayfa / Belge Dili (A Düzeyi):** Metinlerin doğal dilini (`langdetect`) tespit ederek XML yapısındaki `<a:rPr lang="...">` etiketiyle eşleşip eşleşmediğini (ekran okuyucuların doğru telaffuz yapabilmesi için) doğrular.
4. **1.4.3 Kontrast [Minimum] (AA Düzeyi):** Metin ve arka plan renkleri arasındaki bağıl parlaklık kontrast oranını hesaplar (normal metinlerde en az `4.5:1`, büyük metinlerde en az `3:1`) ve ihlal durumunda uygun renk kodları (HEX/RGB) önerir.

---

## 🛠️ Teknoloji Yığını ve Mimari

Yazılım, üç katmanlı modüler bir mimari ile geliştirilmektedir:

* **Programlama Dili:** Python
* **OOXML Ayrıştırma & Görüntü İşleme:** `python-pptx`, `lxml`, `Pillow`, `langdetect`
* **Yapay Zekâ Entegrasyonu:** Çok modlu Üretken Yapay Zekâ API'si (Yapılandırılmış istem şablonları)
* **Grafiksel Kullanıcı Arayüzü (GUI):** `CustomTkinter` (Kullanıcının önerileri *Onayla*, *Düzenle ve Onayla* veya *Reddet* seçenekleriyle yönetebildiği arayüz)

---

## 🔬 Araştırma Yöntemi ve Doğrulama Süreci

Çalışma, **Tasarım ve Geliştirme Araştırması (DDR - Tip 1)** yöntemiyle yürütülmektedir:

* **1. Aşama Test (Sentetik Veri Seti):** Bilinen hata dağılımına (*Ground Truth*) sahip **20 adet sentetik sunu** (200 slayt) üzerinde yerleşik denetleyici ile geliştirilen sistemin Kesinlik (*Precision*), Duyarlılık (*Recall*) ve *F1-Skoru* performansları karşılaştırılır.
* **2. Aşama Test (Gerçek Ortam Doğrulaması):** **Ankara Üniversitesi Açık Ders Malzemeleri (ADM)** arşivinden seçilen **30 adet gerçek ders sunusu** üzerinde denetim başarımı ve ÜYZ önerilerinin kullanıcı onay oranları ölçülür.

---
Not: Bu depodaki kodlar, TÜBİTAK 2209-A projesi kapsamında geliştirilen ilk çalışma prototipini (v0.1) yansıtmaktadır; araştırma takvimi boyunca güncellenecektir.

👤 Proje Ekibi ve Kurum
Proje Yürütücüsü: Semih Can Yıldız

Kurum: Ankara Üniversitesi — Eğitim Bilimleri Fakültesi, Bilgisayar ve Öğretim Teknolojileri Eğitimi (BÖTE) Bölümü

Destek Programı: TÜBİTAK 2209-A (2026 / 1. Dönem)

📄 Lisans ve Etik İlkeler
Bu proje TÜBİTAK Üretken Yapay Zekâ Rehberi (2025) sorumlu yapay zekâ ve şeffaflık ilkelerine uygun olarak geliştirilmektedir. Sistem yalnızca sentetik ve kamuya açık ders materyalleri üzerinde test edilmekte olup hiçbir kişisel veri (KVKK kapsamındaki veri) dış servislere aktarılmamaktadır.
## 📂 Depo Yapısı

```text
boz213v01a02/
├── docs/               # Araştırma önerisi özeti ve teknik dokümantasyon
├── prototype/          # Mevcut çalışma prototipi ve örnek test dosyaları (v0.1)
└── README.md           # Proje tanıtım ve kullanım kılavuzu
