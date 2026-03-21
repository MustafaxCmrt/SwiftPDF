# SwiftPDF v1.0.0

Word, Excel ve PowerPoint dosyalarını tek tıkla PDF'e dönüştüren masaüstü uygulaması.

## Gereksinimler

- **Python 3.10+**
- **LibreOffice** — dönüşüm motoru olarak kullanılır, sisteme kurulu olmalıdır
  - macOS: [https://www.libreoffice.org/download](https://www.libreoffice.org/download) adresinden `.dmg` indirip `/Applications` klasörüne kurun
  - Windows: Aynı adresten `.msi` indirip kurun (varsayılan konum yeterli)

## Kurulum

```bash
git clone <repo-url>
cd SwiftPDF
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Çalıştırma

```bash
source .venv/bin/activate        # Windows: .venv\Scripts\activate
python main.py
```

Uygulama açılışta LibreOffice'in kurulu olup olmadığını kontrol eder. Kurulu değilse indirme sayfasına yönlendiren bir uyarı gösterir.

## Desteklenen formatlar

| Giriş | Çıkış |
|-------|-------|
| `.docx`, `.doc` (Word) | PDF |
| `.xlsx`, `.xls` (Excel) | PDF |
| `.pptx`, `.ppt` (PowerPoint) | PDF |

## Özellikler

- Sürükle & bırak ile dosya ekleme
- Toplu dönüştürme (birden fazla dosya aynı anda)
- Çıktı klasörü seçimi
- Dosya bazında ilerleme takibi (Bekliyor → Dönüştürülüyor → Hazır / Hata)
- Zaman damgalı olay günlüğü
- Koyu / Açık tema
- Windows ve macOS desteği
