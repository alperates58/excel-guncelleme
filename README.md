# Excel SQL Connection & IP Bulk Updater 🚀

**Excel SQL Connect Pro**, 100-200+ adet Excel (`.xlsx`, `.xlsm`, `.xlsb`) dosyasındaki **SQL Server IP adreslerini, OLEDB/ODBC Veri Bağlantı cümlelerini, Power Query formüllerini ve VBA Makro kodlarını** toplu olarak taramak, yedeklemek ve tek tıkla güncellemek için geliştirilmiş web tabanlı bir masaüstü otomasyon aracıdır.

---

## 🌟 Öne Çıkan Özellikler

- **🔍 Kapsamlı Tarama Motoru**:
  - **Power Query Sorguları**: Tüm `Sql.Databases("IP", ...)` M formüllerini tespit eder.
  - **Veri Bağlantıları**: OLEDB, ODBC, Mashup ve Data Model bağlantı cümlelerini (`ConnectionString`), bağlantı adlarını (`Connection Name`) ve SQL komut metinlerini (`CommandText`) inceler.
  - **VBA Makro Kodları**: `.xlsm` ve `.xlsb` dosyalarındaki makro modüllerindeki IP adreslerini tarar.
- **📁 Görsel Web Klasör Seçici (Gözat)**: Masaüstü, İndirilenler, sürücüler (`C:\`, `D:\`) ve alt klasörler arasında görsel gezinti sunar.
- **📊 Canlı İlerleme Çubuğu (Real-time Progress)**: 150+ dosyalık klasörler taranırken ve güncellenirken anlık dosya sayısı, işlenen dosya adı ve yüzde oranını gösterir.
- **🛡️ Otomatik Güvenli Yedekleme**: Değişiklik yapılmadan önce tüm Excel dosyalarını Windows Geçici Klasörüne (`%TEMP%\excel_backups`) güvenle kopyalar.
- **🚀 Portable & Standalone**: Kurulum veya Node.js gerektirmez, herhangi bir Windows bilgisayarda `.bat` dosyasına çift tıklanarak çalışır.

---

## 📁 Proje Yapısı

```
excel-guncelleme/
├── BAŞLAT_Excel_Updater.bat  # Masaüstü hızlı başlatıcı (Çift tıkla çalıştır)
├── server.ps1                # PowerShell REST API sunucusu (Port 3005)
├── engine/
│   └── excel_engine.ps1      # Excel COM interop otomasyon motoru
├── public/
│   ├── index.html            # Ultra modern dark-mode arayüz
│   ├── style.css             # Glassmorphism stil dosyası
│   └── app.js                # Frontend istemci mantığı & canlı polling
└── .gitignore                # Excel ve yedek dosyalarının yüklenmesini engeller
```

---

## 🚀 Çalıştırma Talimatı

1. Projeyi indirin veya klonlayın.
2. Klasör içindeki **`BAŞLAT_Excel_Updater.bat`** dosyasına çift tıklayın.
3. Uygulama otomatik olarak varsayılan tarayıcınızda `http://localhost:3005/` adresinde açılacaktır.

---

## 🔒 Güvenlik & Gizlilik

- Bu depo yalnızca uygulama kaynak kodlarını içerir.
- **Excel dosyaları, şirket verileri ve yedek klasörleri kesinlikle bu depoya yüklenmez** (`.gitignore` korumalıdır).
