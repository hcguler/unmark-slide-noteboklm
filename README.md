# unmark-slide-notebooklm 🧹

Üretken yapay zeka araçları (NotebookLM vb.) tarafından oluşturulan PowerPoint (`.pptx`) sunumlarındaki filigranları ve logoları, sunumun orijinal yapısını ve tasarımını bozmadan otomatik olarak temizleyen bir Python komut satırı aracıdır.

## 🚀 Özellikler

- **Hassas Tespit:** Nesne isimlerine değil, doğrudan `(x, y)` koordinatlarına bakarak sadece sağ alt köşeye yerleştirilmiş resim formatındaki logoları hedefler.
- **Derinlemesine Temizlik:** Sadece normal slaytları değil; Slide Layouts (Slayt Düzenleri) ve Slide Master (Asıl Slayt) yapılarını da tarayarak derine gömülmüş logoları da temizler.
- **Güvenli Silme:** `python-pptx` kütüphanesini kullanarak nesneleri doğrudan OpenXML ağacından güvenle çıkarır, dosyanın bozulmasını (corruption) engeller.

## 📋 Gereksinimler

- Python 3.8 veya üzeri
- `python-pptx` kütüphanesi

## 🛠️ Kurulum

1. Repoyu bilgisayarınıza klonlayın:
   ```bash
   git clone [https://github.com/KULLANICI_ADINIZ/pptx-watermark-cleaner.git](https://github.com/KULLANICI_ADINIZ/pptx-watermark-cleaner.git)
   cd pptx-watermark-cleaner
