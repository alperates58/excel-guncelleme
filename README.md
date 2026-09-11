# Excel SQL Connect Pro 🚀
### Toplu Excel SQL Server IP, Bağlantı Dizesi ve VBA Makro Güncelleyici

[![Excel Automation](https://img.shields.io/badge/Excel-COM%20Interop-green.svg)](https://microsoft.com/excel)
[![PowerShell REST API](https://img.shields.io/badge/Backend-PowerShell%205.1+-blue.svg)](https://microsoft.com/powershell)
[![License](https://img.shields.io/badge/License-MIT-purple.svg)](LICENSE)

**Excel SQL Connect Pro**, şirketinizdeki SQL Server IP adresi veya sunucu adı değiştiğinde yüzlerce (100 - 200+) Excel raporunu tek tek elle düzenleme derdini ortadan kaldıran **taşınabilir (portable), ultra-hızlı ve web tabanlı bir masaüstü otomasyon aracıdır**.

---

## 🎯 Çözülen Problem

SQL Server IP adresi değiştiğinde (Örn: `192.168.2.15` ➔ `10.0.0.100` veya `192.168.2.8` ➔ `192.168.2.50`):
- Excel'deki **Power Query M sorguları** bozulur ve veri çekemez.
- **OLEDB ve ODBC veri bağlantı cümleleri** eski IP'ye bağlanmaya çalışır.
- **VBA Makro kodları** içerisindeki ADODB / DAO bağlantı stringleri hata verir.

Bu araç ile belirttiğiniz klasördeki tüm Excel dosyaları sırayla taranır, eski IP'ler tespit edilir ve **tüm sorgular, veri bağlantıları ve makrolar tek tıkla toplu olarak güncellenir**.

---

## ✨ Öne Çıkan Özellikler

- 🔍 **Kapsamlı 3'ü 1 Arada Tarama Engine**:
  - **Power Query M Sorguları**: `Sql.Databases("192.168.2.15", ...)` tüm formülleri inceler.
  - **Veri Bağlantıları**: OLEDB, ODBC, Mashup ve Data Model bağlantı cümlelerini (`ConnectionString`), bağlantı adlarını (`Connection Name`) ve SQL komut metinlerini (`CommandText`) inceler.
  - **VBA Makro Kodları**: `.xlsm` ve `.xlsb` dosyalarındaki makro modüllerini satır satır tarar.
- 📁 **Görsel Web Klasör Seçici (Gözat)**:
  - Masaüstü, İndirilenler, sürücüler (`C:\`, `D:\`) ve alt klasörler arasında görsel gezinti sunar.
- 📊 **Canlı İlerleme Çubuğu (Real-time Progress Streaming)**:
  - 150+ dosyalık klasörler taranırken ve güncellenirken anlık dosya sayısı (`66 / 157`), işlenen dosya adı ve yüzde oranını (`%42`) ekranda canlı olarak gösterir.
- 🛡️ **Otomatik Güvenli Yedekleme**:
  - Güncelleme işleminden önce dosyalarınızın tamamını otomatik olarak Windows Geçici Klasörüne (`%TEMP%\excel_backups\backup_YYYYMMDD_HHMMSS`) yedekler.
- ⚡ **Taşınabilir & Sıfır Bağımlılık (Portable)**:
  - `Node.js`, `npm` veya üçüncü taraf kurulum gerektirmez. Windows üzerinde yerleşik PowerShell ve Excel Interop ile çalışır. `.bat` dosyasına çift tıklamanız yeterlidir!

---

## 💻 Ekran Görüntüsü ve Kullanım Akışı

```
 +-----------------------------------------------------------------------+
 | Excel SQL Connect Pro                                    ● Sunucu Aktif|
 +-----------------------------------------------------------------------+
 | 📁 Çalışma Klasörü        | 📊 Sayaçlar                               |
 | [ C:\Sirket\Exceller   ]  |   157 Dosya | 45 Sorgu | 12 Bağlantı    |
 | [ 📁 Görsel Klasör Seç ]  |                                           |
 | [ 🔍 Klasörü Tara      ]  | 📋 Tarama & Önizleme Sonuçları Tablosu    |
 |                           |   Dosya Adı | Tür | Sorgu | IP'ler | Detay|
 | ⚙️ IP Değişim Kuralları   |   --------------------------------------- |
 | [ 192.168.2.15 ] ➔ [ 10.0.0.100 ]                                    |
 |                           | 📈 Canlı İlerleme Çubuğu (%42)             |
 | [ ⚡ Toplu Güncelle ]     | 💻 Canlı İşlem Günlüğü (Console)           |
 +-----------------------------------------------------------------------+
```

---

## 🚀 Hızlı Başlangıç (Kullanım Kılavuzu)

1. Projeyi klonlayın veya zip olarak indirin:
   ```bash
   git clone https://github.com/alperates58/excel-guncelleme.git
   ```
2. Klasör içindeki **`BAŞLAT_Excel_Updater.bat`** dosyasına çift tıklayın.
3. Otomatik olarak açılan tarayıcı ekranında (`http://127.0.0.1:3005`):
   - **`Görsel Klasör Seçici (Gözat...)`** butonuna tıklayarak Excel dosyalarınızın bulunduğu klasörü seçin.
   - **`Klasörü Tara`** butonuna basın.
   - Tespit edilen eski IP'nin yanına yeni IP adresinizi yazın.
   - **`Toplu Güncellemeyi Başlat`** butonuna basarak tüm dosyalarınızı saniyeler içinde güncelleyin!

---

## 📁 Proje Klasör Yapısı

```
excel-guncelleme/
├── BAŞLAT_Excel_Updater.bat  # Masaüstü çift tıkla başlatıcı scripti
├── server.ps1                # PowerShell REST API sunucusu (Port 3005)
├── engine/
│   └── excel_engine.ps1      # Excel COM Interop otomasyon & tarama motoru
├── public/
│   ├── index.html            # Ultra modern dark-mode HTML5 arayüz
│   ├── style.css             # Glassmorphism stil kütüphanesi
│   └── app.js                # Frontend istemci mantığı & canlı polling
├── README.md                 # Detaylı Türkçe dokümantasyon
└── .gitignore                # Excel ve yedek dosyalarının gizlilik koruması
```

---

## 🔒 Gizlilik ve Güvenlik İlkeleri

- Bu açık kaynak depo **yalnızca uygulamanın kaynak kodlarını** içermektedir.
- Şirketinize veya şahsınıza ait **Excel dosyaları, veri bağlantıları ve yedek klasörleri kesinlikle bu depoya yüklenmez** (`.gitignore` korumalıdır).

---

## 📜 Lisans

Bu proje [MIT Lisansı](LICENSE) altında piyasaya sürülmüştür. Özgürce kullanabilir ve geliştirebilirsiniz.
