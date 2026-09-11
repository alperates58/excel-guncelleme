# EXCEL SQL CONNECT PRO — TEKNİK MİMARİ, GÜVENİLİRLİK VE QA ANALİZ RAPORU

**Denetçi Rolü:** Kıdemli Yazılım Mimarı / Production Reliability Engineer (SRE) / Lead QA Architect  
**Tarih:** 11 Eylül 2026  
**İncelenen Çalışma Alanı:** `c:\Users\alper.ates.LIDER\Desktop\excel-guncelleme`  
**Durum:** Kesinlikle hiçbir kod değiştirilmedi, dosya silinmedi veya refactor edilmedi. Sistem mevcut haliyle dondurularak statik ve dinamik analize tabi tutuldu.

---

## 1. PROJENİN NE YAPTIĞINI ÇIKAR

### Temel Amaç
Bu proje; kurumsal ağlarda SQL Server IP adresi, sunucu adı veya bağlantı parametreleri değiştiğinde (örneğin şirket sunucusunun `192.168.2.15` adresinden `10.0.0.100` adresine taşınması), şirket bünyesindeki onlarca veya yüzlerce Excel rapor dosyasında (`.xlsx`, `.xlsm`, `.xlsb`) yer alan:
1. **Power Query M Formüllerini** (`Sql.Database("192.168.2.15", ...)`),
2. **Veri Bağlantı Dizelerini & SQL Komutlarını** (OLEDB / ODBC / Data Model Connection String ve CommandText),
3. **Çalışma Sayfası Sorgu Tablolarını** (`Worksheet.QueryTables.Connection`),
4. **VBA Makro Kodlarını** (ADODB / DAO / OLEDB bağlantısı içeren `.xlsm` ve `.xlsb` makro modülleri)

Microsoft Excel COM Interop otomasyonu kullanarak toplu olarak arayan, tespit eden ve yeni değerlerle değiştiren **yerel web tabanlı bir masaüstü yardımcı aracıdır**.

> **Önemli Ayrım:** Bu uygulama Excel hücrelerindeki operasyonel veri satırlarını (stok miktarı, sipariş kaydı vb.) güncellemez. Uygulamanın doğrudan hedefi **Excel'in veri çekme altyapısını oluşturan bağlantı metaverilerini, Power Query M kodlarını ve VBA kaynak kodlarını** güncellemektir. Ancak bu bağlantılar güncellenip dosya kaydedildiğinde Excel'in hesaplama motoru, Pivot önbellekleri ve veri tabloları doğrudan etkilenmektedir.

### Kullanıcı Akışı
1. Kullanıcı `BAŞLAT_Excel_Updater.bat` dosyasına çift tıklar.
2. Batch dosyası PowerShell 5.1 motoruyla `server.ps1` scriptini `-Port 3005` parametresiyle başlatır.
3. `server.ps1`, yerel HTTP dinleyicisini (`System.Net.HttpListener`) ayağa kaldırır ve kullanıcının varsayılan tarayıcısında `http://127.0.0.1:3005/` adresini açar.
4. Web arayüzünde kullanıcı, Excel dosyalarının bulunduğu klasör yolunu girer veya "Görsel Klasör Seçici (Gözat...)" modalından bir dizin seçer.
5. Kullanıcı **"Klasörü Tara"** butonuna basar.
6. Sunucu, dizindeki tüm `.xlsx`, `.xlsm` ve `.xlsb` dosyalarını Excel COM aracılığıyla salt-okunur açar; içlerindeki IPv4 adreslerini regex ile ayıklar ve özetler.
7. Kullanıcıya tespit edilen IP adresleri "Hızlı Seçim Çipleri" (Quick Chips) ve dosya bazlı sorgu/bağlantı/makro sayıları tablosu olarak sunulur.
8. Kullanıcı eski IP/metin ve yeni IP/metin kurallarını tanımlar; Power Query, Veri Bağlantıları, VBA Makroları ve Otomatik Yedekleme onay kutularını seçer.
9. İsteğe bağlı olarak **"Yedeğini Al"** butonuna basarak manuel yedek alır veya doğrudan **"Toplu Güncellemeyi Başlat"** butonuna tıklar.
10. Gelen standart tarayıcı onay (`confirm`) penceresine onay verdikten sonra güncelleme işlemi başlar. Dosyalar tek tek yazma modunda açılarak string değişimi uygulanır, `wb.Save()` ile üzerine yazılır ve sonuçlar arayüzdeki canlı konsola log olarak dökülür.

### Excel Dosyaları Sisteme Nasıl Giriyor?
Sisteme dosya yükleme (`file upload / multipart`) mantığı yoktur. İstemci arayüzü yalnızca dosya sistemindeki bir klasör yolunu (`directory` string) JSON gövdesiyle API'ye gönderir. PowerShell arka planda `Get-ChildItem -Path $DirectoryPath -File` komutu ile doğrudan yerel veya ağ sürücüsündeki dosyaları okur.

### Excel Üzerinde Gerçekleştirilen İşlemler
* **Okuma (Tarama):**
  * `wb.Queries`: Her sorgunun `q.Formula` metni okunur.
  * `wb.Connections`: Her bağlantının `conn.Name`, `OLEDBConnection.ConnectionString`, `OLEDBConnection.CommandText`, `ODBCConnection.ConnectionString`, `ODBCConnection.CommandText` metinleri okunur.
  * `ws.QueryTables`: Her çalışma sayfasındaki `qt.Connection` dizesi okunur.
  * `wb.VBProject.VBComponents`: Her modülün satırları (`CodeModule.Lines`) okunur.
* **Yazma (Güncelleme):**
  * `q.Formula = $newFormula`
  * `conn.Name = $newName` (Eski isim kural metnini içeriyorsa bağlantı adı da değiştirilir!)
  * `conn.OLEDBConnection.ConnectionString = $newCStr`
  * `conn.OLEDBConnection.CommandText = $newCmd`
  * `conn.ODBCConnection.ConnectionString = $newCStr`
  * `conn.ODBCConnection.CommandText = $newCmd`
  * `qt.Connection = $newQtConn`
  * `cm.ReplaceLine($i, $newLine)`
  * `$wb.Save()` (Dosya doğrudan üzerine kaydedilir!)

### Veri Kaynakları ve Hedefleri
* **Okunan Yer:** Dosya sistemindeki fiziksel Excel dosyalarının COM nesne ağacı (`Queries`, `Connections`, `VBProject`).
* **Yazılan Yer:** Doğrudan aynı dosyanın kendisi (In-place disk write).

### Kullanıcı Tarafından Girilen Ayarlar
* Çalışma Klasör Yolu (`folderPath`)
* Arama ve Değiştirme Kuralları: Liste halinde `[{ oldText: string, newText: string }]`
* Seçenekler:
  * `updateQueries` (boolean)
  * `updateConnections` (boolean)
  * `updateVba` (boolean)
  * `autoBackup` (boolean)

### Kalıcı Veriler ve Ayarlar Nerede Tutuluyor?
**HİÇBİR YERDE TUTULMUYOR.** Uygulamada veritabanı (SQLite vb.), konfigürasyon dosyası (`config.json`, `settings.ini`) veya tarayıcı `localStorage` mekanizması yoktur. Sayfa yenilendiğinde (F5) veya sunucu kapatıldığında tüm kurallar, klasör yolu ve işlem geçmişi sıfırlanır. Arayüz açıldığında HTML koduna gömülmüş (hard-coded) geliştirici klasör yolu ve varsayılan test IP'leri (`192.168.2.15` -> `10.0.0.100`) ekrana gelir.

### Uygulama Tipi, Framework ve Kütüphaneler
* **Mimari:** Yerel Web Tabanlı Masaüstü Aracı (Local Web Desktop Utility).
* **Backend:** Saf Windows PowerShell 5.1 scriptleri (`server.ps1`, `excel_engine.ps1`). Framework yoktur; doğrudan .NET Framework sınıfları (`System.Net.HttpListener`, `System.IO.FileStream`, `System.Runtime.InteropServices.Marshal`) ve COM Interop (`Excel.Application`) kullanılır.
* **Frontend:** Saf HTML5, CSS3, Vanilla JavaScript (ES6). UI kütüphanesi (React, Vue vb.) yoktur. Dış bağımlılık olarak CDN üzerinden FontAwesome 6.4.0 ve Google Fonts (Inter, Outfit, JetBrains Mono) çağrılmaktadır.
* **Çalışma / Paketleme:** Derlenmiş exe veya MSI installer yoktur. `BAŞLAT_Excel_Updater.bat` dosyası ile çalıştırılır.

### Özet Cümle
> **"Bu proje, SQL Server taşıma ve IP değişiklikleri sonrasında çok sayıdaki Excel raporunda bozulan Power Query sorgularını, OLEDB/ODBC veri bağlantılarını ve VBA makro kodlarını; harici bir runtime (Node/Python) gerektirmeksizin yerel bir PowerShell HTTP sunucusu ve Microsoft Excel COM Interop otomasyonu üzerinden web arayüzü ile toplu olarak arayıp doğrudan dosya üzerine güncelleyen taşınabilir bir masaüstü otomasyon aracıdır."**

---

## 2. DOSYA VE MİMARİ HARİTASI

```
excel-guncelleme/
│
├── BAŞLAT_Excel_Updater.bat   # [Giriş Noktası] Ortamı hazırlar, PowerShell sunucusunu başlatır
├── server.ps1                 # [API & Web Sunucusu] HttpListener döngüsü, statik dosya ve REST API sunar
│
├── engine/
│   └── excel_engine.ps1       # [İş Mantığı & COM] Excel COM işlemleri, tarama, yedekleme ve güncelleme
│
├── public/
│   ├── index.html             # [Arayüz İskeleti] Kontrol paneli, tablolar, sayaçlar ve modallar
│   ├── style.css              # [Arayüz Stili] Dark theme, glassmorphism CSS kütüphanesi
│   └── app.js                 # [İstemci Mantığı] API istekleri, polling, modal yönetimi, DOM render
│
├── inspect_excels.ps1         # [Test/Araç] OpenXML zip tabanlı customXml inceleme scripti
├── inspect_siparis.ps1        # [Test/Araç] COM tabanlı bağlantı dizesi inceleme scripti
├── scan_all_samples.ps1       # [Test/Araç] Dizin genelinde COM tarama scripti
├── test_com.ps1               # [Test/Araç] COM nesnesi ve VBA erişim deneme scripti
├── test_folder.ps1            # [Test/Araç] STA Shell.Application klasör seçici testi
├── test_picker_proc.ps1       # [Test/Araç] Harici WinForms FolderBrowserDialog süreci testi
│
├── PLAN v4.2.xlsm             # [Örnek/Gerçek Veri] 3.7 MB Makrolu kurumsal üretim planlama dosyası
├── REÇETE PATLATMA v4.xlsm    # [Örnek/Gerçek Veri] 3.2 MB Makrolu reçete patlatma dosyası
├── SİPARİŞ İHTİYAÇ v3.xlsx    # [Örnek/Gerçek Veri] 889 KB Sipariş ihtiyaç takip dosyası
├── ÖZET VE ÜRETİM DATA v1.xlsm# [Örnek/Gerçek Veri] 709 KB Üretim veri takip dosyası
└── .gitignore                 # Excel, yedek ve geçici dosyaların git takibini engelleyen kural seti
```

### Detaylı Bileşen Analiz Tablosu

| Dosya / Klasör | Görevi | Kim Tarafından Çağrılıyor? | Neyi Çağırıyor? | Kritik mi? | Temel Riskleri |
| :--- | :--- | :--- | :--- | :---: | :--- |
| **`BAŞLAT_Excel_Updater.bat`** | Giriş noktası (Entry point). Konsol penceresini açar ve `powershell.exe` ile sunucuyu tetikler. | Son kullanıcı (Çift tık) | `server.ps1` | **EVET** | PowerShell çalıştırma ilkesi (ExecutionPolicy), admin yetkisi eksikliği veya yol boşluklarında çalışma kesintisi. |
| **`server.ps1`** | HTTP REST API ve Web sunucusu. `System.Net.HttpListener` ile port dinler, routing yapar, statik dosyaları döndürür. | `BAŞLAT_Excel_Updater.bat` | `engine\excel_engine.ps1`, `public/*`, .NET Framework API'leri | **EVET** | **Tek iş parçacıklı (Single-threaded) senkron döngü.** Tarama/güncelleme sırasında sunucu bloke olur; `/api/progress` isteklerine yanıt veremez. Unhandled exception durumunda HTTP soketi asılı kalır. |
| **`engine/excel_engine.ps1`** | Çekirdek iş mantığı ve Excel COM motoru. Dosyaları açar, Regex ile IP arar, string replace yapar, dosyayı diske kaydeder. | `server.ps1` (Dot-source: `. $engineScript`) | `Excel.Application` COM nesnesi, dosya sistemi (`System.IO`) | **KRİTİK** | **Doğrudan in-place overwrite.** Atomik yazma yoktur; kısmi başarısızlıkta dosya bozulur. COM RCW sızıntısı (leak) nedeniyle arkada zombi `excel.exe` süreçleri kalır. Regex yerine düz string replace yapıldığı için alt-IP çarpışması (`192.168.1.1` -> `192.168.1.10` yaparken `192.168.1.15`'in `192.168.1.105` olması). |
| **`public/index.html`** | Web arayüzü görünümü. | Tarayıcı (istemci) | `style.css`, `app.js`, CDN font/ikonları | **ORTA** | CDN erişimi olmayan kapalı kurumsal ağlarda ikonlar ve yazı tipleri yüklenemez; arayüz görsel olarak bozulur. |
| **`public/app.js`** | İstemci kontrol mantığı. Buton tıklamaları, API istekleri, periyodik progress polling, dinamik tablo filtreleme. | `index.html` | `/api/*` uç noktaları | **YÜKSEK** | Sunucu senkron çalıştığı için polling kilitlenir; işlem sonrası hemen ikinci bir `runScan()` tetikleyerek kullanıcıyı gereksiz bekletir. Geri alma (rollback) düğmesi yoktur. |
| **`public/style.css`** | Glassmorphism dark-theme tasarım stilleri. | `index.html` | Google Fonts | **DÜŞÜK** | Performansa veya veri bütünlüğüne doğrudan riski yoktur. |
| **`inspect_*.ps1` / `test_*.ps1`** | Geliştirme aşamasında yazılmış geçici test ve inceleme scriptleri. | Yalnızca manuel çalıştırılır | Excel COM, .NET ZipArchive | **DÜŞÜK** | Üretimde kullanılmıyor ancak repo kökünde dağınık duruyor; yetkisiz çalıştırıldığında yan etki riski taşır. |
| **Kurumsal Excel Dosyaları (`.xlsm`, `.xlsx`)** | Şirkete ait gerçek üretim, planlama ve sipariş verilerini içeren dosyalar. | Manuel inceleme / Test | SQL Server veritabanları | **KRİTİK** | **Gerçek şirket verileri `.gitignore` kuralına rağmen yerel depoda bulunmaktadır!** Yanlış bir işlemde bu dosyalar geri dönülemez şekilde bozulabilir. |

### Mimari Bağımlılık Değerlendirmesi
* **Aşırı Bağımlılık (Tight Coupling):** `server.ps1`, `excel_engine.ps1` dosyasını `.` (dot-source) ile aynı scope içerisine dahil etmektedir. Her iki dosya da `$global:ProgressState` global değişkenine bağımlıdır.
* **COM ve İş Parçacığı Kilidi:** Web sunucusunun HTTP istek dinleme döngüsü ile Excel COM'un 15-20 saniye süren dosya açma-kaydetme işlemleri **aynı tek iş parçacığı üzerinde** dönmektedir. Arka plan iş parçacığı (`Runspace`, `Job` veya `Task`) kullanılmamıştır.

---

## 3. EXCEL İŞLEMLERİNİN DERİN ANALİZİ

### Dosya Açma Mekanizması
* **Varlık ve Filtre Kontrolü:** `Get-ExcelFiles` fonksiyonu `Get-ChildItem` ile klasörü tarar. Sadece `.xlsx`, `.xlsm` ve `.xlsb` uzantılarını alır; `~$` ile başlayan geçici kilit dosyalarını ve `backup_` ile başlayan yedekleri eler.
* **Eksik Doğrulamalar:**
  * Eski ikili format olan `.xls` (Excel 97-2003) tamamen yoksayılmaktadır. Şirkette eski `.xls` raporları varsa bunlar taranmaz ve güncellenmez.
  * Döngü içinde her dosya açılmadan önce `Test-Path $file.FullName` kontrolü yapılmaz; dosya listelendikten sonra silinirse veya taşınırsa unhandled exception oluşur.
* **Dosya Kilitli / Başka Kullanıcıda Açık İse Ne Oluyor?**
  * `Scan` fonksiyonunda dosya `ReadOnly = $true` açılarak sorun kısmen aşılır.
  * Ancak `Update` fonksiyonunda: `$excel.Workbooks.Open($file.FullName, 0, $false)` çağrılır. Eğer dosya ağdaki başka bir kullanıcı veya kullanıcının kendi masaüstündeki Excel tarafından açıksa, Excel COM `$excel.DisplayAlerts = $false` olduğundan dosyayı kullanıcıya sormadan **sessizce salt-okunur (read-only)** açar! Ardından `$wb.Save()` satırına gelindiğinde:
    `COMException: Workbook is read-only and cannot be saved.`
    hatası fırlatılır ve dosya kaydedilemeden sonraki dosyaya geçilir.
* **Ağ Paylaşımı, UNC Path ve Uzun Yollar:**
  * UNC yolları (`\\fileserver\raporlar`) `Get-ChildItem` ve Excel COM tarafından açılabilir. Ancak ağdaki gecikme (latency) veya mikro kesintilerde Excel COM kilitlenir (hang).
  * 260 karakterden uzun Windows yollarında (`MAX_PATH`) `System.IO.PathTooLongException` veya Excel COM açılış hatası meydana gelir. Uzun yol ön eki (`\\?\`) kullanılmamıştır.
* **Türkçe Karakter Durumu:**
  * PowerShell ve .NET UTF-8 karakterleri tanır; dosya adlarında `ğ, ü, ş, ı, ö, ç` olması COM açılışını bozmaz. Ancak `BAŞLAT_Excel_Updater.bat` dosyasında `chcp 65001` aktif edilmediği için Windows komut satırında Türkçe karakterli klasör isimleri bozulabilmektedir.

### Excel Dosyasını Okuma Mekanizması
* **Çalışma Sayfaları:** Kod çalışma sayfalarının isimlerine bağımlı değildir (`$wb.Worksheets` koleksiyonu taranır).
* **Power Query M Sorguları:** `$wb.Queries` koleksiyonundaki `$q.Formula` okunur. M dili içindeki `Sql.Database("IP", "DB")` ifadeleri düz metin olarak taranır.
* **Veri Bağlantıları (Connections):**
  * OLEDB için `$conn.OLEDBConnection.ConnectionString` ve `CommandText`
  * ODBC için `$conn.ODBCConnection.ConnectionString` ve `CommandText` okunur.
* **Hücre ve Tablo Düzeyi:**
  * Kod hücre bazında (`Range("A1")`) okuma yapmaz. Dolayısıyla merged cell, boş hücreler, tarih biçimleri, decimal nokta/virgül ayrımı gibi hücre düzeyi tuzaklar **okuma aşamasında** koda etki etmez.
  * Ancak çalışma sayfası üzerindeki `QueryTable` nesneleri (`$ws.QueryTables`) taranırken `$qt.Connection` okunur.

### Excel'e Yazma ve Nesne Bütünlüğünün Korunması

Excel COM kullanıldığında (`$wb.Save()`), dosya üçüncü taraf bir XML ayrıştırıcı (OpenXML / ClosedXML / EPPlus) ile değil bizzat Microsoft Excel'in kendi çekirdeği tarafından diske yazılır. Bu durumun avantajları ve çok kritik dezavantajları şunlardır:

| Excel Özelliği | Korunuyor mu? | Teknik Açıklama ve Risk Değerlendirmesi |
| :--- | :---: | :--- |
| **Hücre Değerleri & Formüller** | **RİSKLİ** | Hücrelere doğrudan dokunulmaz. **ANCAK:** Dosya açıldığında Excel'in hesaplama motoru (`Calculation = xlCalculationAutomatic`) devrededir. IP değiştirilip SQL bağlantısı koptuğunda veya henüz erişilemez olduğunda, formüller yeniden hesaplanıp `#REF!`, `#N/A` veya `#VALUE!` hatalarına dönebilir ve dosya bu hatalı değerlerle diske kaydedilebilir! |
| **Conditional Formatting** | Evet | Excel COM tarafından çalışma kitabı şeması bozulmadan saklanır. |
| **Hücre Stilleri, Renk, Font, Kenarlık** | Evet | Excel COM workbook nesnesini tam olarak saklar. |
| **Data Validation (Veri Doğrulama)** | Evet | Excel COM tarafından korunur. |
| **Named Ranges (Adlandırılmış Alanlar)**| Evet | Korunur. |
| **Pivot Tables & Data Model** | **ÇOK KRİTİK RİSK** | **Bağlantı adı değiştirilirse Pivot Table çöker!** `excel_engine.ps1` satır 389'da `$conn.Name = $newName` çalıştırılmaktadır. Eğer bir Pivot Tablo veya Veri Modeli eski bağlantı adına (`conn.Name`) referans veriyorsa, bağlantı adı değiştiğinde Pivot Tablo yetim kalır ve veri yenilemesi bozulur! |
| **Grafikler & Şekiller** | Evet | Korunur. |
| **Gizli Satır ve Sütunlar** | Evet | Korunur. |
| **VBA Makro Kodları (`.xlsm`/`.xlsb`)**| **ŞARTLI KORUNUR** | 1. **Güven Merkezi Engeli:** Excel ayarlarında *"VBA projesi nesne modeline erişime güven"* seçeneği kapalıysa `wb.VBProject` erişimi fırlatır ve makrolar güncellenemez.<br>2. **VBA Şifresi:** Makro projesi şifreli ise kodlara erişilemez ve hata verir.<br>3. Kod satır satır değiştirilir (`ReplaceLine`). |
| **External Links (Dış Bağlantılar)** | Riskli | Dosya açılışında `UpdateLinks = 0` verilmiştir (güncelleme engellenir), ancak kaydederken bağlar korunur. |
| **Power Query Kimlik Bilgileri** | **BOZULUR** | Excel'de Power Query kimlik bilgileri (kullanıcı adı/şifre) dosya içinde değil, kullanıcının yerel Windows kimlik deposunda (`Data Source Settings`) IP bazlı tutulur. `192.168.2.15` -> `10.0.0.100` yapıldığında Excel bu IP'yi yepyeni bir sunucu görür ve kullanıcı ilk "Yenile" dediğinde şifre sorar. Kullanıcı bunu "Makro/Sorgu bozuldu" sanabilir. |

---

## 4. DOSYANIN BOZULMA RİSKİ ANALİZİ (CRITICAL)

### Mevcut Durum: DOĞRUDAN IN-PLACE OVERWRITE

Mevcut `Update-ExcelDirectory` fonksiyonu incelendiğinde (`excel_engine.ps1`, satır 523):
```powershell
if ($fileModified) {
    $wb.Save()
    $fileLog.status = "Updated"
    $updatedFilesCount++
}
```
**Tespit:** Uygulama doğrudan orijinal dosyanın üzerine yazmaktadır (`$wb.Save()`).

### Felaket Senaryoları Karşısındaki Davranış
1. **Elektrik Kesintisi / Windows Kapanması:** `$wb.Save()` disk yazma G/Ç (I/O) işlemi yürütürken elektrik kesilirse veya bilgisayar kapanırsa dosya başlıkları (ZIP headers / OpenXML parts) yarıda kalır. **Sonuç: Excel dosyası tamamen bozulur (Corrupted File) ve açılamaz.**
2. **Uygulama veya COM Çökmesi (Crash):** Bellek yetersizliği veya Excel'in GPF (General Protection Fault) vermesi durumunda dosya kilitli veya yarı yazılmış kalır.
3. **Disk Dolması:** `$wb.Save()` sırasında diskte yeterli alan yoksa Excel dosyayı 0 KB veya eksik part ile bırakır.
4. **Ağ Paylaşımı Kopması:** Dosya `\\fileserver` üzerindeyse ve ağ paketi düşerse dosya kilidi askıda kalır ve dosya bozulur.

### Olması Gereken Güvenli Pipeline vs. Mevcut Durum

```
GÜVENLİ VE HATASIZ MİMARİ (HEDEF)        MEVCUT UYGULAMA (GÖZLEMLENEN)
┌─────────────────────────────────┐      ┌─────────────────────────────────┐
│        ORIGINAL.xlsx            │      │        ORIGINAL.xlsx            │
└────────────────┬────────────────┘      └────────────────┬────────────────┘
                 │                                        │
        [1. Güvenli Backup Al]                            │
                 │                                        │
┌────────────────▼────────────────┐                       │ [Doğrudan Aç & Bellekte Değiştir]
│ %USERPROFILE%\Backups\Dosya.bak │                       │
└────────────────┬────────────────┘                       │
                 │                                        │
        [2. TEMP Dosya Oluştur]                           │
                 │                                        │
┌────────────────▼────────────────┐                       │
│       TEMP_XYZ.xlsx             │                       │
└────────────────┬────────────────┘                       │
                 │                                        │
        [3. TEMP Üzerinde Değiştir]                       │
                 │                                        │
        [4. TEMP Dosyayı Doğrula]                         │
                 │                                        │
        [5. Atomik Değiştir / Swap]                       │
                 │                                        │
        (File.Replace / Atomic)                           ▼
                 │                       ┌─────────────────────────────────┐
                 ▼                       │    $wb.Save() [DİREKT OVERWRITE] │
┌─────────────────────────────────┐      │   * Elektrik/Ağ kesilirse       │
│  ORIGINAL.xlsx (GÜNCELLENDİ)    │      │     DOSYA GERİ DÖNÜŞSÜZ BOZULUR!│
└─────────────────────────────────┘      └─────────────────────────────────┘
```

> **CRITICAL BULGU:** Mevcut sistem atomik dosya değiştirme (`File.Replace`), geçici çalışma dosyası (`temp staging`) ve işlem sonrası dosya bütünlük doğrulaması (ZIP bütünlük testi) uygulamamaktadır. Güncelleme anındaki herhangi bir kesinti veya kilitlenme, **orijinal dosyanın kalıcı olarak bozulmasına (data corruption)** yol açar.

---

## 5. BACKUP / ROLLBACK MEKANİZMASI İNCELEMESİ

### Yedekleme Davranışı
* **Yedek Alma Yeri:** `Create-ExcelBackup` fonksiyonu (`excel_engine.ps1`, satır 255):
  `$backupParent = Join-Path $env:TEMP "excel_backups"`
  `$backupDir = Join-Path $backupParent "backup_$timestamp"`
* **Yedekleme Yöntemi:** `Copy-FileWithShare` fonksiyonu ile `[System.IO.FileShare]::ReadWrite` modunda dosya akışı okunup geçici klasöre kopyalanmaktadır.

### Ciddi Kusurlar ve Riskler
1. **`%TEMP%` Klasörü Güvenilmezdir:** Windows Geçici Klasörü (`AppData\Local\Temp`), Windows'un "Akıllı Depolama" (Storage Sense), CCleaner gibi temizlik araçları veya sistem yeniden başlatmaları tarafından otomatik olarak temizlenebilir. Kullanıcının haftalar sonra fark ettiği hatalı bir güncellemede yedekler çoktan silinmiş olabilir.
2. **Kullanıcı Yedek Yolunu Seçemez:** Yedeklerin harici diske, güvenli bir ağ klasörüne veya çalışma klasörünün yanına alınması seçeneği sunulmamıştır.
3. **Eski Yedekler Asla Temizlenmez:** `excel_backups` klasöründe hiçbir zaman rotasyon veya temizlik (retention policy) yapılmaz. 300 MB'lık 10 dosya içeren bir klasör 10 kez güncellendiğinde geçici dizinde 30 GB alan sessizce işgal edilir.
4. **ROLLBACK (GERİ YÜKLEME) MEKANİZMASI YOKTUR:**
   * Ne web arayüzünde ne de API uç noktalarında bir "Yedekten Geri Yükle" (Rollback / Restore) düğmesi veya fonksiyonu **BULUNMAMAKTADIR**.
   * Kullanıcı yanlış bir kural uyguladığında (örneğin yanlışlıkla tüm `192` bloklarını bozduğunda), yedek dosyalarını `%TEMP%\excel_backups` altından elle bulup tek tek kopyalamak zorundadır.
   * **Prensip İhlali:** "Excel güncellemesi hiçbir zaman geri döndürülemez olmamalıdır" şartı mevcut sistemde karşılanmamaktadır.

---

## 6. VERİ DOĞRULUĞU VE EŞLEŞME RİSKLERİ

### Eşleşme Mantığı Nasıl Çalışıyor?
Uygulama veritabanı benzeri bir Primary Key veya kayıt anahtarı ile çalışmaz. Tüm eşleşmeler metin arama ve değiştirme mantığına (`Contains` ve `Replace`) dayanır.

### Yanlış Hücre / Metin Güncelleme Riskleri

#### 1. Alt-IP Çarpışması (Partial String / IP Substring Collision) — KRİTİK
Taramada regex kullanılırken (`$ipRegex`), güncelleme aşamasında (`excel_engine.ps1` satır 357, 381, 398, 503) saf .NET string `Replace` kullanılmaktadır:
```powershell
if ($newFormula.Contains($rule.oldText)) {
    $newFormula = $newFormula.Replace($rule.oldText, $rule.newText)
}
```
* **Felaket Senaryosu:**
  * Kural: `192.168.1.1` ➔ `10.0.0.10`
  * Formül içindeki gerçek IP: `192.168.1.15`
  * `Contains("192.168.1.1")` doğru döner!
  * `Replace("192.168.1.1", "10.0.0.10")` çalışır.
  * **Oluşan Hatalı Değer:** `10.0.0.105`!
  * **Sonuç:** SQL bağlantısı sessizce var olmayan `10.0.0.105` IP'sine yönlendirilir ve rapor çöker!
  * Regex sınır ayracı (`\b` word boundary) kullanılmadığı için her alt-dize eşleşmesi bir veri bozulması kaynağıdır.

#### 2. Kural Sırası ve Kademeli Değişim (Cascading Overwrite) — YÜKSEK
Birden fazla kural girildiğinde kurallar sırayla döngüye girer:
* Kural 1: `192.168.2.10` ➔ `192.168.2.20`
* Kural 2: `192.168.2.20` ➔ `10.0.0.50`
* Dosyadaki `192.168.2.10`, Kural 1 ile `192.168.2.20` olur; hemen ardından Kural 2'ye girerek `10.0.0.50` olur. İki kural birbirini ezer; kullanıcı ilk kuralın sonucunu öngöremez.

#### 3. Büyük/Küçük Harf Duyarlılığı (Case-Sensitivity) — YÜKSEK
PowerShell 5.1'deki `.NET` `[string]::Contains()` metodu varsayılan olarak **büyük/küçük harfe duyarlıdır (case-sensitive)**:
* Kullanıcı kurala eski sunucu adı olarak `sql-server` yazarsa, connection string içindeki `Server=SQL-SERVER;` veya `Data Source=Sql-Server;` ifadesi eşleşmez!
* Güncelleme sessizce atlanır, dosyada değişiklik yapılmaz ancak kullanıcıya "İşlem Tamamlandı" denir.

#### 4. Bağlantı Adı Değişikliği Yan Etkisi — YÜKSEK
`excel_engine.ps1` satır 389'da bağlantı nesnesinin adı da değiştirilmektedir:
```powershell
$oldName = $conn.Name
if ($newName.Contains($rule.oldText)) { $conn.Name = $newName }
```
Eğer bağlantının adı `Baglanti_192.168.2.15` ise ve adı `Baglanti_10.0.0.100` yapılırsa; bu bağlantıyı kullanan makro kodları (`ActiveWorkbook.Connections("Baglanti_192.168.2.15").Refresh`) ve Pivot önbellekleri **RUN-TIME ERROR '9': Subscript out of range** hatasıyla çöker.

#### 5. VBA Satır İçi Yorum ve Değişken Değişimi — ORTA
VBA kodlarında satır analizi yapılmadan düz `Replace` uygulanır. Eğer bir kod satırında açıklama (`' 192.168.2.15 eski sunucudur`) veya bir IP string değişkeni varsa bağlam ayrımı yapılmaksızın ezilir.

---

## 7. TRANSACTION (İŞLEM) BENZERİ DAVRANIŞ VAR MI?

### Cevap: KESİNLİKLE YOKTUR (ALL OR NOTHING SAĞLANMIYOR)

Projede transaction mantığı bulunmamaktadır.
* **Senaryo:** 100 adet Excel dosyası güncellenirken 80 tanesi başarıyla güncellendi ve 81. dosyada Excel COM bir hata fırlattı (`Error: Cannot access file` veya `Out of memory`).
* **Mevcut Davranış:**
  1. İlk 80 dosya çoktan diske yazılmış ve orijinal halleri değiştirilmiştir.
  2. 81. dosyada hata alınır, `$fileLog.status = "Error"` olarak işaretlenir.
  3. Döngü devam eder ve 82. dosyadan 100. dosyaya kadar işlem yapmaya çalışır.
  4. Hata olsa dahi `Update-ExcelDirectory` fonksiyonu en sonda `@{ success = $true, ... }` döner!
  5. Web arayüzü yeşil renkte **"Güncelleme Başarıyla Tamamlandı!"** pop-up'ı gösterir!
  6. Kullanıcı işlem günlüğüne dikkatle bakmazsa 81. dosyanın bozuk kaldığından haberdar olamaz.
  7. Hata anında önceki 80 dosyayı eski haline döndürecek bir geri alma (rollback) tetiklenmez.

---

## 8. HATA YÖNETİMİ ANALİZİ

Tüm kod tabanındaki `try / catch` blokları tarandı. Ciddi mimari ve mantıksal kusurlar tespit edildi:

### 1. Sessizce Yutulan Hatalar (Swallowed Exceptions / Boş Catch Blokları)
* `excel_engine.ps1` Satır 133: Power Query taramasında `catch { }` — boş bırakılmış, hata fırlatırsa sessizce sonraki adıma geçiyor.
* `excel_engine.ps1` Satır 148: OLEDB/ODBC bağlantı okumasında `catch { }` — boş bırakılmış.
* `excel_engine.ps1` Satır 169: Connection taramasında `catch { }` — boş bırakılmış.
* `excel_engine.ps1` Satır 189: QueryTables taramasında `catch { }` — boş bırakılmış.
* `excel_engine.ps1` Satır 425: OLEDB güncelleme bloğunda `catch { }` — boş bırakılmış.
* `excel_engine.ps1` Satır 459: ODBC güncelleme bloğunda `catch { }` — boş bırakılmış.
* `excel_engine.ps1` Satır 489: QueryTables güncelleme bloğunda `catch { }` — boş bırakılmış.
* `server.ps1` Satır 37: Tarayıcı açma işleminde `catch { }` — boş bırakılmış.

### 2. HTTP İstemcisini Sonsuz Beklemede Bırakan Catch (Hanging Request)
`server.ps1` Satır 217:
```powershell
    } catch {
        Write-Host "Request handling error: $_" -ForegroundColor Red
    }
```
Eğer ana döngüde bir istek işlenirken beklenmedik bir exception fırlatılırsa (örneğin JSON serileştirme hatası veya bellek taşması), catch bloğu hatayı sadece konsola yazar; ancak `$response.OutputStream.Close()` çağrılmaz ve HTTP yanıtı gönderilmez! Tarayıcıdaki `fetch()` isteği sonsuza kadar askıda kalır (pending timeout).

### 3. Hata Durumunda Sahte "Başarılı" Mesajı (False Positive Success)
`excel_engine.ps1` Satır 543: `Update-ExcelDirectory` fonksiyonu, içindeki 10 dosya hata alsa dahi en dışta her zaman `success = $true` dönmektedir:
```powershell
return @{
    success = $true
    directory = $DirectoryPath
    totalFilesProcessed = $files.Count
    updatedFilesCount = $updatedFilesCount
    totalReplacements = $totalReplacements
    logs = $updateLog
}
```
`app.js` satır 620 ise `if (data.success)` kontrolü yaparak ekrana **"Güncelleme Başarıyla Tamamlandı!"** uyarısı basmaktadır. Kullanıcı kısmi hatalardan haberdar edilmemektedir.

---

## 9. LOGGING (GÜNLÜKLEME) SİSTEMİ İNCELEMESİ

### Mevcut Durum
* **Disk Logu Yoktur:** Uygulama diske hiçbir `.log` veya `.txt` dosyası yazmaz.
* **Geçici Bellek Logu:** Loglar sadece web sayfasındaki `<div id="logConsole">` içerisine HTML elementi olarak basılır.
* **Audit Trail (Denetim İzi) Sıfırdır:** Sayfa yenilendiğinde, sekme kapatıldığında veya sunucu durdurulduğunda tüm loglar geri getirilemez şekilde silinir.

### Bir Kullanıcı "Excel Yanlış Güncellendi" Dediğinde Ne Oluyor?
Geriye dönük hiçbir şey tespit edilemez. Kimin çalıştırdığı, saat kaçta çalıştığı, hangi dosyanın hangi satırındaki hangi IP'nin neye dönüştürüldüğü kalıcı olarak kaydedilmemiştir.

---

## 10. CONCURRENCY VE ÇOKLU KULLANIM RİSKLERİ

1. **İki Kullanıcı veya Sekme:** `server.ps1` yerel makinede tek iş parçacığıyla çalışır. Eğer yerel ağdaki başka bir bilgisayar IP üzerinden sunucuya bağlanır ve aynı anda tarama başlatırsa, ilk istek bitene kadar ikinci istek HTTP.sys kuyruğunda bekletilir.
2. **Aynı Excel Dosyasına Eşzamanlı Erişim:** Excel COM, dosya üzerinde işletim sistemi düzeyinde kilit mekanizmasını tam kontrol edemez. Bir dosya açıkken güncelleme çalıştırılırsa dosya ya bozulur ya da kaydedilemez.
3. **Stale Lock (Askıda Kalan Kilitler):** Tarama veya güncelleme sırasında bir hata meydana gelip `wb.Close($false)` çalışmazsa, dosya Excel COM nesnesi tarafından kilitli tutulur. Kullanıcı Windows Explorer'da dosyayı açmak istediğinde "Dosya başka bir program tarafından kullanılıyor" uyarısı alır.

---

## 11. UI/UX VE KULLANICI HATASI RİSKLERİ

1. **Önizleme (Dry-Run / Preview) Özelliği Yoktur:**
   Kullanıcı "Toplu Güncellemeyi Başlat" demeden önce sistemin hangi dosyalarda tam olarak kaç değişiklik yapacağını gösteren bir önizleme tablosu YOKTUR.
2. **Geri Dönüşü Zor Onay Diyaloğu:**
   Kullanıcıya sadece standart tarayıcı `confirm()` penceresi gösterilir. Bu pencerede etkilenecek dosya sayısı veya tahmini süre yer almaz.
3. **Gözat Butonu & Yol Karışıklığı:**
   Arayüz açıldığında geliştiricinin kendi masaüstü yolu (`c:\Users\alper.ates.LIDER\...`) input kutusunda varsayılan olarak yazılı gelmektedir. Dikkatsiz bir kullanıcı doğrudan "Tara" butonuna bastığında yanlış klasör taranır.

---

## 12. PERFORMANS ANALİZİ

### 1. Tek İş Parçacıklı HTTP Sunucusu Blokesi (Architectural Flaw)
`server.ps1` satır 76'daki `while ($listener.IsListening)` döngüsü **senkron** çalışır.
* Kullanıcı "Klasörü Tara" veya "Güncelle" butonuna bastığında, ana iş parçacığı 150 dosyayı tek tek açıp kapatmak üzere döngüye girer (bu işlem 2-5 dakika sürer).
* Bu 2-5 dakika boyunca `listener.GetContext()` çağrılamaz!
* Frontend her 400 ms'de bir `/api/progress` isteği atar; ancak sunucu iş parçacığı meşgul olduğu için bu isteklerin HİÇBİRİNE yanıt veremez!
* **Sonuç:** Canlı ilerleme çubuğu (progress bar) işlem sırasında donar veya hiç ilerlemez; işlem bittiğinde bir anda %100 olur. "Real-time streaming" iddiası mimari olarak çalışmamaktadır!

### 2. Büyük Excel Dosyalarında COM Açılış Yükü
100-200 dosyalık bir klasörde her dosya için Excel COM motoru tam olarak başlatılmakta, XML ağacı ayrıştırılmakta ve kapatılmaktadır. 50 MB'lık raporlarda işlem dosya başına 10-15 saniye sürer. 1.000 dosyada işlem saatler alacaktır.

---

## 13. BELLEK VE KAYNAK SIZINTISI (MEMORY / COM LEAK)

### .NET COM Interop "İki Nokta Kuralı" (Two-Dot Rule) İhlali
PowerShell'de COM nesneleriyle çalışırken ara nesneler serbest bırakılmazsa Runtime Callable Wrapper (RCW) bellekten atılamaz.
* Kodda:
  * `$wb.Queries`
  * `$conn.OLEDBConnection`
  * `$ws.QueryTables`
  * `$wb.VBProject.VBComponents`
  gibi onlarca COM alt nesnesi doğrudan zincirleme çağrılmakta ve `Marshal::ReleaseComObject` yapılmamaktadır.
* Sadece `$wb` için `ReleaseComObject` çağrılmaktadır; ancak alt nesnelerin referansı canlı kaldığı için Excel süreci hafızadan temizlenemez.
* **Sonuç:** 100 dosya tarandıktan sonra Görev Yöneticisi'nde (Task Manager) kapanmayan, 1-2 GB RAM tüketen zombi `EXCEL.EXE` süreçleri birikir.

---

## 14. GÜVENLİK ANALİZİ

1. **Uzaktan Komut Çalıştırma (RCE / Command Injection) — KRİTİK:**
   `server.ps1` satır 157:
   ```powershell
   Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"`" `"$targetDir`"" -WindowStyle Hidden
   ```
   Eğer istemciden `/api/open-folder` uç noktasına `C:\ & calc.exe &` gibi bir payload gönderilirse doğrudan sistem komutu çalıştırılır!
2. **Dizin Geçişi (Path Traversal):**
   `/api/list-folders` ve `/api/scan` gelen `directory` girdisini doğrulamadan dosya sistemine iletir. Kullanıcı `C:\Windows\System32` gibi kritik sistem klasörlerini taratabilir.
3. **CORS Açıklığı:**
   `Access-Control-Allow-Origin: *` başlığı tanımlıdır. Kullanıcı arka planda kötü niyetli bir web sitesini ziyaret ederse, o site kullanıcının yerelindeki `http://127.0.0.1:3005` sunucusuna istek atarak dosyaları taratabilir veya silebilir.
4. **VBA / Formula Injection:**
   Yeni değer olarak girilen metin filtrelenmeden doğrudan VBA koduna veya formül içine basılmaktadır. Zararlı VBA kodu (`Shell "cmd.exe"`) makrolu dosyalara kalıcı olarak enjekte edilebilir.

---

## 15. KONFİGÜRASYON VE TAŞINABİLİRLİK ANALİZİ

1. **Hard-coded Geliştirici Yolu:** `public/index.html` satır 50'de `c:\Users\alper.ates.LIDER\Desktop\excel-guncelleme` yolu hard-coded olarak yazılmıştır. Başka bir kullanıcının bilgisayarında varsayılan olarak var olmayan bir klasör açılır.
2. **Hard-coded IP'ler:** `192.168.2.15` ve `10.0.0.100` değerleri doğrudan kaynak koda gömülüdür.
3. **Port Tutarsızlığı:** `server.ps1` içinde `$Port = 3000` varsayılan parametre iken, `BAŞLAT_Excel_Updater.bat` dosyası sunucuyu `3005` portu ile başlatmaktadır.

---

## 16. KURULUM, DAĞITIM VE ÇALIŞMA ORTAMI (BUILD & ENVIRONMENT)

Sıfır bir Windows bilgisayarda çalıştırmak için gereksinimler:
* **İşletim Sistemi:** Windows 10 veya Windows 11.
* **Runtime:** Windows PowerShell 5.1 (Windows ile yerleşik gelir).
* **Office Bağımlılığı (ZORUNLU):** Bilgisayarda **fiziksel Microsoft Excel kurulu olmak ZORUNDADIR.** Excel yüklü olmayan bir sunucuda veya istemcide uygulama HİÇ ÇALIŞMAZ (`New-Object -ComObject Excel.Application` çöker).
* **Excel Güven Merkezi (ZORUNLU):** Makrolu dosyalardaki (`.xlsm`) VBA kodlarını güncelleyebilmek için Excel'de:
  `Seçenekler -> Güven Merkezi -> Güven Merkezi Ayarları -> Makro Ayarları -> "VBA projesi nesne modeli erişimine güven"`
  kutusunun işaretli olması şarttır. Aksi halde VBA güncellemesi sessizce başarısız olur.

---

## 17. TEST DURUMU ANALİZİ

* **Mevcut Testler:** Projede Pester, Unit Test, Entegrasyon Testi veya CI/CD testi **YOKTUR**.
* **Mevcut Scriptler:** `inspect_*.ps1` ve `test_*.ps1` dosyaları geçici manuel deneme amaçlı yazılmıştır, otomatik test niteliği taşımaz.
* **Olması Gereken Test Paketi:**
  * Farklı Excel formatları (`.xlsx`, `.xlsm`, `.xlsb`, şifreli dosyalar, bozuk dosyalar) için izole `fixtures/` test seti.
  * Regex sınır ayracı (`\b`) doğrulama testleri.
  * Rollback ve atomik dosya değiştirme testleri.

---

## 18. GERÇEK HAYAT EDGE CASE LİSTESİ (30 SENARYO)

| No | Senaryo | Risk | Şu Anki Davranış | Beklenen Güvenli Davranış | Seviye |
| :---: | :--- | :--- | :--- | :--- | :---: |
| **1** | Güncelleme sırasında bilgisayarın elektriği kesildi veya kapandı. | Dosyanın kalıcı olarak bozulması (0 KB veya bozuk ZIP). | Doğrudan orijinalin üzerine yazıldığı için dosya bozulur. | Geçici dosyada çalışılmalı; işlem bitmeden orijinale dokunulmamalı. | **BLOCKER** |
| **2** | Kural `192.168.1.1` -> `10.0.0.10`, dosyada `192.168.1.15` var. | Yanlış IP oluşması (`10.0.0.105`). | `.Replace()` ile alt dize ezilir, yanlış IP oluşur. | `\b` Word boundary regex ile sadece tam eşleşmeler değiştirilmeli. | **CRITICAL** |
| **3** | Excel dosyası ağda başka bir kullanıcı tarafından açık tutuluyor. | Değişikliğin kaydedilememesi veya üzerine yazma çatışması. | Excel COM sessizce salt-okunur açar; kaydederken hata fırlatır. | İşlem öncesi dosya kilit kontrolü yapılmalı; kilitliyse atlanıp uyarılmalı. | **CRITICAL** |
| **4** | Dosya salt-okunur (Read-only dosya sistemi özniteliği). | Dosyanın kaydedilememesi. | Save sırasında unhandled exception verir. | Salt-okunur özniteliği kontrol edilmeli, kullanıcı uyarılmalı. | **HIGH** |
| **5** | Dosya şifre korumalı (Workbook Password / Open Password). | COM sürecinin arka planda kilitlenmesi veya hata vermesi. | COM hatası verir, dosya atlanır; detay loglanmaz. | Şifreli olduğu tespit edilip "Şifreli Dosya" statüsüyle raporlanmalı. | **HIGH** |
| **6** | VBA Makro projesi şifreli (VBA Project Locked). | Makro kodunun güncellenememesi. | Hata verir, dosya "Updated" sanılabilir. | "VBA Şifreli - Atlandı" uyarısı verilmeli, dosya durumu Warning olmalı. | **HIGH** |
| **7** | Excel Trust Center'da "VBA erişimine güven" kapalı. | Hiçbir `.xlsm` dosyasının makrolarının güncellenememesi. | COMException (0x800A03EC) fırlatır, sessizce geçilir. | Uygulama başlangıcında registry'den Trust Center kontrolü yapılmalı. | **HIGH** |
| **8** | Güncellenecek klasörde 200 dosya var, 50. dosyada disk doldu. | 50 dosya güncellendi, 50 dosya kaldı; 50. dosya bozuldu. | İşlem yarıda kalır, rollback yapılmaz, sahte success döner. | Disk alanı önceden kontrol edilmeli; hata anında rollback sorulmalı. | **CRITICAL** |
| **9** | Eski sunucu adı küçük harfle yazılmış (`srv-sql`), kuralda büyük (`SRV-SQL`). | Sunucu adının güncellenememesi. | `.Contains()` büyük/küçük harf duyarlı olduğu için atlar. | Case-insensitive eşleşme seçeneği sunulmalı. | **HIGH** |
| **10** | Dosya yolu 260 karakterden uzun (MAX_PATH). | Dosyanın açılamaması. | `PathTooLongException` verir, işlem başarısız olur. | `\\?\` Long path prefix desteği eklenmeli. | **MEDIUM** |
| **11** | Dosya adı Türkçe karakterli (`İŞLETME_VERİLERİ_ŞUBAT.xlsx`). | Batch veya konsol karakter kodlama bozukluğu. | Batch dosyasında chcp eksik olduğu için konsolda bozuk görünür. | UTF-8 tam desteklenmeli (`chcp 65001`). | **LOW** |
| **12** | Klasörde eski `.xls` (Excel 97-2003) dosyası var. | Dosyanın fark edilmeden eski IP'de kalması. | Uzantı filtresi `.xls`'i almaz, tamamen yok sayar. | `.xls` tespit edilip modern formata dönüştürme uyarısı verilmeli. | **MEDIUM** |
| **13** | Dosya indirilenlerden alınmış ve "Mark of the Web" (Zone.Identifier) var. | Excel'in dosyayı Korumalı Görünümde (Protected View) açması. | COM nesnesi formülleri veya makroları düzenleyemez. | Dosyanın Zone.Identifier engeli kaldırılmalı (Unblock-File). | **HIGH** |
| **14** | Pivot Table bir bağlantıya bağlı ve bağlantı adı değiştirildi (`$conn.Name = $newName`). | Pivot tablonun yenilenememesi (bozulması). | Bağlantı adı ezilir, Pivot tablo yetim kalır. | Bağlantı adı (`conn.Name`) ASLA değiştirilmemeli; sadece dize güncellenmeli.| **CRITICAL** |
| **15** | Formül hücresi dış SQL verisine bağlı (`Automatic Calculation`). | SQL kapalıyken dosya açılıp kaydedilirse hücrelerin `#REF!` olması. | Formüller bozuk değerlerle kaydedilir. | Dosya açılırken `Calculation = xlManual` yapılmalı. | **CRITICAL** |
| **16** | Power Query kimlik bilgileri yerel Windows kasasında kayıtlı. | Kullanıcı Excel'i açtığında SQL bağlantı hatası alması. | IP değiştiği için Excel yeni kimlik sorar. | Kullanıcıya Power Query kimliklerinin sıfırlanacağı uyarısı verilmeli. | **MEDIUM** |
| **17** | 1. Kural `A -> B`, 2. Kural `B -> C`. | Kademeli kural birbirini ezer, veriler beklenmedik hale gelir. | Sırayla replace yapıldığı için A olan yerler C olur. | Kurallar tek geçişte (token-based / parallel replace) uygulanmalı. | **HIGH** |
| **18** | Bağlantı dizesinde parola var (`Password=123;`). | Parolanın yanlışlıkla loglara yazılması. | Konsol loglarında connection string metni dökülürse parola ifşa olur. | Hassas veriler (Password/Pwd) loglarda maskelenmeli (`***`). | **HIGH** |
| **19** | Dosya içinde Power Query sorgusu yok ama Data Connection var. | Yanlış tespit veya eksik güncelleme. | Arayüzde sorgu sayısı 0 görünür, kullanıcı şaşırabilir. | Rapor detayında nesne türleri açıkça ayrıştırılmalı. | **LOW** |
| **20** | Klasör boş veya içinde Excel dosyası yok. | Arayüzün kilitlenmesi veya çökmesi. | Boş liste döner, sayaçlar 0 olur (düzgün ele alınmış). | Bilgilendirici boş durum mesajı verilmeli (mevcut). | **LOW** |
| **21** | Ağ paylaşımında klasör seçildi (`\\192.168.1.50\Paylasim`). | Ağ kopmasında COM sürecinin sonsuz kilitlenmesi. | Script kilitlenir, zaman aşımı (timeout) yoktur. | COM işlemleri için zaman aşımı (timeout) mekanizması kurulmalı. | **HIGH** |
| **22** | Tarama sırasında web sayfası kapatıldı veya yenilendi. | Arka planda `excel.exe` sürecinin yetim kalması. | PowerShell ve Excel arka planda çalışmaya devam eder, kaynak sızar. | İptal (cancellation token / abort) mekanizması eklenmeli. | **MEDIUM** |
| **23** | İstemci `/api/open-folder` uç noktasına komut enjeksiyonu gönderdi. | Sunucuda yetkisiz komut çalışması (RCE). | `cmd.exe /c start` parametresi nedeniyle komut çalıştırılır. | Sadece doğrulanmış klasör yolları `explorer.exe` ile açılmalı. | **CRITICAL** |
| **24** | VBA kodunda satır devamı karakteri (`_`) kullanılmış. | Bağlantı dizesinin iki satıra bölünmesi ve eşleşmemesi. | Satır satır arandığı için parçalanmış IP tespit edilemez. | Modülün tüm metni üzerinde multiline regex çalıştırılmalı. | **MEDIUM** |
| **25** | Dosya geçici dosyadır (`~$Rapor.xlsx`). | Hata fırlatılması. | `Get-ExcelFiles` `~$` filtresi ile eler (düzgün ele alınmış). | Korunmalı. | **LOW** |
| **26** | Klasör yolu sonunda ters slash var veya yok (`C:\Test\` vs `C:\Test`). | Yol birleştirmede çift slash veya geçersiz yol hatası. | `Join-Path` kullanıldığı için PowerShell bunu tolere eder. | Korunmalı. | **LOW** |
| **27** | Port 3005 başka bir program tarafından kullanılıyor (Örn: Node servisi). | Sunucunun başlayamaması. | Batch dosyası hata kodu verir ve kapanır. | Port kullanımda ise alternatif porta geçilmeli (3006, 3007...). | **MEDIUM** |
| **28** | Kullanıcı 100.000 satırlık devasa bir Excel dosyasını güncelliyor. | Yetersiz bellek (Out of Memory) ve Excel çökmesi. | Tüm dosya belleğe alınır, çökebilir. | Bellek kullanımı izlenmeli. | **HIGH** |
| **29** | Kullanıcı "Klasörü Tara" dedikten sonra sunucu progress yanıtı vermiyor. | Kullanıcının uygulamanın donduğunu sanıp kapatması. | Tek iş parçacıklı döngü nedeniyle progress yanıtı gecikir. | Arka plan iş parçacığı (Runspace) ile progress ayrılmalı. | **HIGH** |
| **30** | `%TEMP%` klasöründeki yedekler Windows tarafından silindi. | Kullanıcının eski haline dönememesi. | Kullanıcı yedek klasörünü bulamaz. | Yedekler proje veya kullanıcı belgeleri altında tutulmalı. | **HIGH** |

---

## 19. BUG AVI (HATA KATALOĞU)

### BUG-001 [CRITICAL]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 357, 381, 398, 413, 433, 448, 475, 503 (`Update-ExcelDirectory`)
* **Severity:** **CRITICAL**
* **Problem:** IP değiştirme işleminde Regex Boundary (`\b`) yerine düz `.Replace()` kullanılması.
* **Nasıl Oluşur:** Kullanıcı `192.168.1.1` adresini `10.0.0.1` yapmak istediğinde, dosya içinde `192.168.1.100` varsa bu adres `10.0.0.100` haline gelir; `192.168.1.15` varsa `10.0.0.15` değil `10.0.0.15` şeklinde bozulur.
* **Sonuç:** SQL IP adresleri yanlış IP'lere dönüşür; raporlar sessizce bozulur ve veri çekemez.
* **Önerilen Çözüm:** Değiştirme işlemi `[regex]::Replace($str, "\b" + [regex]::Escape($old) + "\b", $new)` ile yapılmalıdır.

---

### BUG-002 [CRITICAL]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 389 (`Update-ExcelDirectory`)
* **Severity:** **CRITICAL**
* **Problem:** Veri bağlantısının adının (`conn.Name`) doğrudan değiştirilmesi.
* **Nasıl Oluşur:** `if ($newName -ne $oldName) { $conn.Name = $newName }` çalışarak bağlantının Excel içindeki ID/Name değerini değiştirir.
* **Sonuç:** Bu bağlantıyı adıyla referans alan Pivot Tablolar, Veri Modelleri (Data Model / PowerPivot) ve VBA makroları kırılır; Excel dosyası açıldığında "Bağlantı bulunamadı" hatası verir.
* **Önerilen Çözüm:** Bağlantı adı (`conn.Name`) kesinlikle değiştirilmemelidir. Yalnızca bağlantı dizesi (`ConnectionString`) ve komut metni (`CommandText`) güncellenmelidir.

---

### BUG-003 [BLOCKER]
* **Dosya:** `server.ps1`
* **Satır / Fonksiyon:** Satır 76-196 (`while ($listener.IsListening)`)
* **Severity:** **BLOCKER**
* **Problem:** Senkron, tek iş parçacıklı döngü mimarisi nedeniyle HTTP dinleyicisinin bloke olması.
* **Nasıl Oluşur:** `/api/scan` veya `/api/update` çağrıldığında ana iş parçacığı döngüye girer. Tarayıcıdan 400 ms'de bir gelen `/api/progress` istekleri HTTP.sys kuyruğunda kilitlenir ve işlem bitene kadar yanıtlanamaz.
* **Sonuç:** Arayüzdeki canlı ilerleme çubuğu işlem boyunca çalışmaz, donar. Kullanıcı uygulamanın çöktüğünü düşünür.
* **Önerilen Çözüm:** Tarama ve güncelleme işlemleri PowerShell `[runspace]` veya `Start-ThreadJob` ile arka plan iş parçacığına taşınmalı; ana iş parçacığı progress isteklerine anında yanıt vermelidir.

---

### BUG-004 [CRITICAL]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 523 (`Update-ExcelDirectory`)
* **Severity:** **CRITICAL**
* **Problem:** Atomik yazma olmaksızın doğrudan dosya üzerine kaydetme (`$wb.Save()`).
* **Nasıl Oluşur:** Güncelleme sırasında elektrik kesintisi, bellek yetersizliği, Excel çökmesi veya ağ kopması meydana geldiğinde.
* **Sonuç:** Orijinal Excel dosyası 0 byte veya bozuk ZIP partları ile kalır; kalıcı veri kaybı oluşur.
* **Önerilen Çözüm:** Değişiklik geçici bir `.tmp` dosyasında yapılmalı; dosya doğrulanmalı ve ardından `[System.IO.File]::Replace` ile atomik olarak orijinal dosya ile yer değiştirilmelidir.

---

### BUG-005 [HIGH]
* **Dosya:** `server.ps1`
* **Satır / Fonksiyon:** Satır 217-219 (`catch`)
* **Severity:** **HIGH**
* **Problem:** Hata yakalandığında HTTP soketinin kapatılmaması (Hanging Request).
* **Nasıl Oluşur:** İstek gövdesi işlenirken veya API çağrısında beklenmeyen bir exception oluştuğunda.
* **Sonuç:** `$response.OutputStream.Close()` çağrılmadığı için tarayıcıdaki istek sonsuza kadar `Pending` kalır ve arayüz kilitlenir.
* **Önerilen Çözüm:** `try / finally` bloğu içinde `$response.StatusCode = 500` atanarak yanıt akışı mutlaka kapatılmalıdır.

---

### BUG-006 [HIGH]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 543 (`Update-ExcelDirectory`)
* **Severity:** **HIGH**
* **Problem:** Dosya hatalarına rağmen her zaman `success = $true` dönülmesi (False Positive).
* **Nasıl Oluşur:** 10 dosyadan 9'u kilitli olduğu için güncellenemediğinde bile API `success: true` döner.
* **Sonuç:** Arayüz kullanıcıya yeşil kutuyla "Tüm Güncellemeler Başarıyla Tamamlandı" der; kullanıcı dosyaların güncellenmediğini fark edemez.
* **Önerilen Çözüm:** Hatalı dosya sayısı > 0 ise `success = $false` veya `status = "partial_success"` dönülmeli ve arayüzde kırmızı/sarı uyarı verilmelidir.

---

### BUG-007 [CRITICAL / GÜVENLİK]
* **Dosya:** `server.ps1`
* **Satır / Fonksiyon:** Satır 157 (`/api/open-folder`)
* **Severity:** **CRITICAL**
* **Problem:** `cmd.exe` üzerinden filtrelenmemiş komut çalıştırma (Arbitrary Command Injection).
* **Nasıl Oluşur:** `directory` parametresine `C:\ & calc.exe &` gönderildiğinde.
* **Sonuç:** Yerel makinede yetkisiz komut çalıştırılabilir.
* **Önerilen Çözüm:** `cmd.exe` aradan çıkarılmalı; doğrudan `[System.Diagnostics.Process]::Start("explorer.exe", $validatedPath)` kullanılmalıdır.

---

### BUG-008 [HIGH]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 110-220 ve 347-531 (COM Nesne Yaşam Döngüsü)
* **Severity:** **HIGH**
* **Problem:** Child COM nesnelerinin (`Queries`, `Connections`, `VBComponents`) `Marshal::ReleaseComObject` ile serbest bırakılmaması.
* **Nasıl Oluşur:** Yüzlerce dosya tarandıktan sonra.
* **Sonuç:** RCW referansları temizlenemediği için arka planda onlarca `EXCEL.EXE` süreci askıda kalır; bellek ve CPU sızıntısı oluşur.
* **Önerilen Çözüm:** Açılan her COM alt nesnesi için `ReleaseComObject` çağrılmalı ve döngü sonunda `[GC]::Collect()` disiplini uygulanmalıdır.

---

### BUG-009 [MEDIUM - POTENTIAL BUG]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 89 (`$ipRegex`)
* **Severity:** **MEDIUM**
* **Problem:** SQL Server sunucu adlarının (Hostname, FQDN veya Named Instance) taranmaması.
* **Nasıl Oluşur:** Dosyalardaki bağlantılar IP (`192.168.2.15`) yerine sunucu adı (`SRV-SQL01` veya `SRV-SQL01\MIKRODB`) içeriyorsa regex bunları yakalayamaz.
* **Sonuç:** İlgili dosyalar "IP Bulunamadı" olarak işaretlenir ve kullanıcı tarafından güncellenemez.
* **Önerilen Çözüm:** Regex deseni SQL sunucu adlarını ve instance adlarını (`Server=...;` veya `Data Source=...;`) kapsayacak şekilde genişletilmelidir.

---

### BUG-010 [HIGH]
* **Dosya:** `engine/excel_engine.ps1`
* **Satır / Fonksiyon:** Satır 347 (`Workbooks.Open`)
* **Severity:** **HIGH**
* **Problem:** Otomatik formül hesaplamasının kapatılmaması (`Calculation Mode`).
* **Nasıl Oluşur:** SQL Server'a erişilemeyen bir ortamda dosya COM ile açılıp kaydedildiğinde.
* **Sonuç:** Excel formülleri yeniden hesaplar, dış veri gelmediği için hücreler `#REF!`, `#N/A` veya `#VALUE!` olur ve dosya bu hatalı değerlerle diske yazılır.
* **Önerilen Çözüm:** Dosya açılmadan önce `$excel.Calculation = -4135` (`xlCalculationManual`) yapılmalı ve `$excel.CalculateBeforeSave = $false` ayarlanmalıdır.

---

## 20. PRODUCTION READINESS DEĞERLENDİRMESİ

| Değerlendirme Kriteri | Puan (100 Üzerinden) | Gerekçe / Durum Özeti |
| :--- | :---: | :--- |
| **Veri Güvenliği** | 15 / 100 | In-place overwrite yapılıyor; atomik yazma yok; kısmi başarısızlıkta geri alma yok. |
| **Excel Bütünlüğü** | 30 / 100 | COM kullanılıyor ancak bağlantı adı değiştirme (`conn.Name`) ve otomatik hesaplama riskleri mevcut. |
| **Error Handling** | 20 / 100 | Yutulan (swallowed) boş catch blokları var; hata anında kullanıcıya sahte "Başarılı" dönülüyor. |
| **Logging & Audit** | 10 / 100 | Diske log yazılmıyor; sayfa kapatılınca tüm işlem geçmişi kayboluyor. |
| **Backup / Rollback** | 25 / 100 | Yedek `%TEMP%` altında tutuluyor (güvenilmez); arayüzde veya kodda geri yükleme (rollback) yok. |
| **Kullanıcı Hatalarına Dayanıklılık** | 20 / 100 | Önizleme (dry-run) yok; onay penceresi yetersiz; hard-coded yollar kafa karıştırıcı. |
| **Concurrency & Kilitleme** | 15 / 100 | Tek iş parçacıklı sunucu; dosya kilit kontrolü yok; stale lock riski yüksek. |
| **Performans & Kaynak Yönetimi** | 30 / 100 | COM nesneleri sızıyor (leak); progress polling HTTP sunucusunu kilitliyor. |
| **Güvenlik** | 25 / 100 | `/api/open-folder` üzerinde RCE riski; CORS `*`; path traversal kontrolü yok. |
| **Sürdürülebilirlik & Kod Kalitesi** | 40 / 100 | Kod az ve anlaşılır ancak ayrıştırılmamış (monolitik scriptler, UI'a gömülü veri). |
| **Test Kapsamı** | 0 / 100 | Sıfır otomatik test (unit/integration test bulunmuyor). |
| **Dağıtım & Taşınabilirlik** | 50 / 100 | Portable bat scripti var ancak Excel kurulumu ve Trust Center ayarına sıkı sıkıya bağımlı. |

### GENEL PRODUCTION READINESS SKORU:
$$\mathbf{23\ /\ 100}$$

> **Nihai Karar:** Bu uygulama mevcut haliyle **ÜRETİM ORTAMINDA KULLANILAMAZ (NOT PRODUCTION READY)**. Dosya bozulması, yanlış IP değişimi ve yetersiz hata yönetimi nedeniyle kurumsal veriler için yüksek risk taşımaktadır.

---

## 21. TEKNİK BORÇLARIN SINIFLANDIRILMASI

### P0 — HEMEN DÜZELTİLMELİ (Veri Kaybı / Dosya Bozulması)
1. **In-place Overwrite Kaldırılmalı:** Geçici staging dosyası (`temp.xlsx`) oluşturulup doğrulanmadan orijinal dosya ezilmemelidir (`File.Replace` atomik pipeline kurulmalı).
2. **Regex Boundary Düzeltilmeli:** Alt dize çarpışmalarını (`192.168.1.1` -> `192.168.1.100`) önlemek için tam sınır (`\b`) regex değişimine geçilmeli.
3. **Bağlantı Adı Değişimi Kaldırılmalı:** Pivot tabloları ve makroları kırmamak için `$conn.Name = $newName` satırı derhal iptal edilmeli.
4. **Manuel Hesaplama Modu:** Formül bozulmalarını engellemek için açılışta `xlCalculationManual` zorunlu kılınmalı.

### P1 — YÜKSEK ÖNCELİK (Üretimde Ciddi Arıza Çıkaracak Noktalar)
1. **Senkron HTTP Sunucu Blokesi:** Tarama ve güncelleme arka plan iş parçacığına (`Runspace`) taşınarak progress polling ve web sunucusu birbirinden ayrılmalı.
2. **Geri Yükleme (Rollback) Butonu:** Web arayüzüne ve API'ye tek tıkla son yedeğe dönebilen Rollback mekanizması eklenmeli.
3. **Kalıcı Dosya Günlüğü (File Logging):** Her işlem `logs/update_YYYYMMDD.log` dosyasına detaylı denetim iziyle yazılmalı.
4. **Sahte Success Düzeltilmeli:** Hata alan dosyalar varsa arayüze kırmızı bildirim ve hata listesi dönülmeli.
5. **RCE Açığı Kapatılmalı:** `/api/open-folder` içindeki `cmd.exe /c start` kaldırılmalı.

### P2 — İYİLEŞTİRME (Stabilite ve Bakım)
1. **COM Object RCW Yönetimi:** Açılan her alt COM nesnesi `ReleaseComObject` ile serbest bırakılmalı, zombi `excel.exe` süreçleri engellenmeli.
2. **Önizleme (Dry-Run):** Kullanıcı güncellemeden önce nelerin değişeceğini gösteren "Önizleme / Simülasyon" modunu çalıştırabilmeli.
3. **Güven Merkezi Doğrulaması:** Başlangıçta Excel VBA güvenliği kontrol edilmeli, kapalıysa kullanıcı uyarılmalı.
4. **Yedek Konumu:** Yedekler `%TEMP%` yerine güvenli bir dizine alınmalı ve saklama süresi (retention) tanımlanmalı.

### P3 — GELİŞTİRME (UX ve Taşınabilirlik)
1. Hard-coded yolların ve IP'lerin arayüzden temizlenmesi.
2. Hostname ve Named Instance desteği.
3. Pester test paketinin kurulması.

---

## 22. KAPSAMLI DEĞİŞİKLİK PLANI (PHASE 1 - PHASE 10)

*(Not: Bu aşamada hiçbir kod değiştirilmeyecektir. Gelecekteki uygulama adımları için rehberdir.)*

* **PHASE 1 – Veri Güvenliği:** `engine/excel_engine.ps1` dosyasında `Update-ExcelDirectory` fonksiyonu baştan yazılarak atomik dosya değiştirme (`Temp Staging -> Verify -> Backup -> File.Replace`) mimarisine geçirilecek.
* **PHASE 2 – Excel Bütünlüğü:** `excel_engine.ps1` içinde `$conn.Name` değişimi kaldırılacak; `xlCalculationManual` modu eklenecek; regex kelime sınırları (`\b`) zorunlu kılınacak.
* **PHASE 3 – Error Handling:** `excel_engine.ps1` ve `server.ps1` içindeki boş catch'ler kaldırılacak; yapılandırılmış hata nesneleri (`@{ success = $false, errorDetails = ... }`) üretilecek.
* **PHASE 4 – Logging & Audit:** `server.ps1` ve `excel_engine.ps1` motoruna günlük dosyası yazıcı (`Write-AppLog`) eklenecek; `logs/` dizinine UTF-8 JSON/text log yazılacak.
* **PHASE 5 – Backup / Rollback:** Yedekler `%TEMP%` yerine `backups/` altına taşınacak; API'ye `/api/rollback` uç noktası ve UI'a "Yedekten Geri Al" butonu eklenecek.
* **PHASE 6 – Validation & Preview (Dry-Run):** API'ye `/api/dry-run` uç noktası eklenecek; diske yazmadan önce etkilenecek dosya ve satır listesi modalda onaylatılacak.
* **PHASE 7 – Tests:** Pester 5 tabanlı test süiti kurulacak; `fixtures/` altına örnek dosyalar konularak regresyon testleri yazılacak.
* **PHASE 8 – Performance:** `server.ps1` çoklu iş parçacıklı (Runspace) yapıya geçirilecek; progress polling ana iş parçacığından bağımsızlaştırılacak; COM nesneleri için deterministic dispose yazılacak.
* **PHASE 9 – UX İyileştirmeleri:** `index.html` ve `app.js` içerisindeki hard-coded yollar kaldırılacak; kural validasyonları ve gelişmiş filtreleme eklenecek.
* **PHASE 10 – Deployment:** `BAŞLAT_Excel_Updater.bat` güçlendirilecek; Excel ve Trust Center ön gereksinim kontrol scripti eklenecek.

---

## 23. MEVCUT DAVRANIŞI BOZMA İLKESİ

İlerleyen geliştirme sürecinde mevcut iş kurallarını korumak için şu prensipler uygulanmalıdır:
1. **Power Query Formül Sözdizimi:** Power Query M dilinde formül `Sql.Database("IP", "DB")` yapısındadır. Yapılacak regex sadece sunucu parametresini hedeflemeli, parantez veya tırnak yapısına dokunmamalıdır.
2. **OLEDB/ODBC Bağlantı Cümlesi:** `Data Source=` veya `Server=` anahtarları dışındaki parametreler (`Catalog`, `Persist Security Info`) asla bozulmamalıdır.
3. **VBA Modül İsimleri:** Modül isimleri ve bileşen tipleri (`vbext_ct_StdModule`, `vbext_ct_MSForm`) korunmalı, sadece satır içi metinler hedeflenmelidir.
4. **Regresyon Risk Yönetimi:** Gerçek şirket dosyaları üzerinde işlem yapılmadan önce mutlaka hash doğrulamalı (SHA256) kopyalar üzerinde test edilmelidir.

---

## 24. YÖNETİCİ ÖZETİ VE KRİTİK ÇIKTILAR

### A. Projenin 10 Maddelik Özeti
1. Proje, şirket SQL Server IP'si değiştiğinde yüzlerce Excel dosyasındaki bağlantıları toplu güncelleyen yerel bir masaüstü aracıdır.
2. Mimarisi: Windows PowerShell 5.1 HTTP dinleyicisi + Excel COM Interop + Vanilla JS/HTML5 web arayüzüdür.
3. Harici dil bağımlılığı (Node.js, Python vb.) gerektirmez; taşınabilirdir.
4. Power Query formüllerini, OLEDB/ODBC veri bağlantılarını, QueryTable'ları ve VBA modüllerini tarar.
5. Hücre verilerini doğrudan güncellemez; bağlantı metaverilerini ve makro kodlarını günceller.
6. Dosyaları doğrudan kendi üzerlerine kaydeder (`In-place overwrite`).
7. Ayarları ve logları kalıcı olarak hiçbir veritabanı veya dosyada saklamaz.
8. Yedekleri sistemin geçici dizinine (`%TEMP%`) alır ve geri alma (rollback) mekanizması sunmaz.
9. Senkron çalışan tek iş parçacığı nedeniyle işlem anında ilerleme çubuğu donar.
10. Mevcut haliyle güvenilirlik skoru **23/100** olup üretim için kritik riskler barındırmaktadır.

### B. En Kritik 10 Risk
1. **Doğrudan Üzerine Yazma:** Elektrik/sistem kesintisinde dosyanın kurtarılamaz şekilde bozulması.
2. **Alt-IP Çarpışması:** `192.168.1.1` değişiminde `192.168.1.15`'in `10.0.0.105` olarak bozulması.
3. **Rollback Yokluğu:** Yanlış bir kural girildiğinde geri dönüşün olmaması.
4. **Bağlantı Adının Değiştirilmesi:** Pivot tabloların ve Veri Modellerinin yetim kalarak çökmesi.
5. **Otomatik Hesaplama Bozulması:** SQL kapalıyken formüllerin `#REF!` değerine dönüp kaydedilmesi.
6. **Güvenlik Açığı (RCE):** `/api/open-folder` üzerinden komut enjeksiyonu yapılabilmesi.
7. **Geçici Yedeklerin Silinmesi:** Windows Storage Sense tarafından `%TEMP%` yedeklerinin temizlenmesi.
8. **Bellek / Zombi Süreç Sızıntısı:** Kapanmayan `EXCEL.EXE` süreçlerinin sistemi kilitlemesi.
9. **Sahte Başarı Mesajı:** Dosyalar hata alsa bile ekranda "Başarıyla Tamamlandı" yazması.
10. **Single-Thread Kilidi:** Web sunucusunun tarama/güncelleme anında diğer isteklere yanıt verememesi.

### C. Kesin Görülen Buglar
* **BUG-001:** Kelime sınırı (`\b`) olmaması nedeniyle alt dize / IP çarpışması.
* **BUG-002:** `conn.Name` değiştirilerek Pivot Tablo bağlantılarının koparılması.
* **BUG-003:** Senkron tek iş parçacıklı döngünün progress polling'i bloke etmesi.
* **BUG-004:** Atomik olmayan dosya kaydetme (`$wb.Save()`).
* **BUG-005:** Exception yakalandığında HTTP soketinin açık kalıp tarayıcıyı dondurması.
* **BUG-006:** Hata durumlarında bile API'nin `success = $true` dönmesi.
* **BUG-007:** `server.ps1` üzerinde `cmd.exe /c start` ile komut enjeksiyonu (RCE).
* **BUG-008:** Child COM nesnelerinin release edilmemesi sonucu zombi süreçler kalması.

### D. Potansiyel Buglar
* **BUG-009:** SQL sunucu isimlerinin (Hostname / FQDN) regex tarafından tanınmaması.
* **BUG-010:** Dış veri kapalıyken `CalculationAutomatic` nedeniyle hücre formüllerinin bozulması.
* Power Query kimlik bilgilerinin IP değişimi sonrası oturum açma hatası vermesi.
* Çok uzun yollarda (>260 karakter) dosyanın sessizce atlanması.

### E. Veri Kaybına Yol Açabilecek Noktalar
* `$wb.Save()` anında diskin dolması veya sürecin kapanması.
* Kısmi başarısızlıkta transaction desteği olmadığı için bazı dosyaların güncellenip bazılarının yarım kalması.
* `%TEMP%` altındaki yedeğin Windows tarafından silinmesi ve geri dönüşün imkansızlaşması.

### F. Yanlış Excel Verisi Oluşturabilecek Noktalar
* Tam sınır ayracı kullanılmadığı için `192.168.1.1` kuralının benzer tüm IP'leri yanlış adreslere çevirmesi.
* Kademeli kural tanımlandığında (`A -> B` ve `B -> C`) verilerin ikinci kural tarafından yanlış ezilmesi.
* Büyük/küçük harf uyumsuzluğunda bağlantı dizelerinin güncellenmeyip eski kalması.

### G. Dosya Bozulmasına Yol Açabilecek Noktalar
* Excel COM açıkken ağ bağlantısının kopması ve dosya kilitlerinin kırılamaması.
* Makro projelerinde `ReplaceLine` kullanılırken satır devam karakterlerinin (`_`) bölünmesi sonucu VBA syntax error oluşması.

### H. Öncelikli Yapılacak İlk 10 Değişiklik
1. `Update-ExcelDirectory` içinde atomik dosya kaydetme pipeline'ı kurmak (`temp` dosya -> doğrula -> `File.Replace`).
2. IP değişiminde `[regex]::Replace` ile `\b` sınır denetimi getirmek.
3. `$conn.Name = $newName` atamasını derhal kaldırmak.
4. Dosya açılırken Excel hesaplama modunu manuel (`xlCalculationManual`) yapmak.
5. `/api/rollback` uç noktası ve arayüze "Yedekten Geri Al" fonksiyonu eklemek.
6. Yedekleme yolunu `%TEMP%` yerine `backups/` yerel klasörüne çekmek.
7. `/api/open-folder` üzerindeki `cmd.exe` RCE açığını gidermek.
8. HTTP sunucusunda tarama ve güncellemeyi arka plan iş parçacığına (`Runspace`) taşımak.
9. Kalıcı dosya loglama (`Write-AppLog`) mekanizması kurmak.
10. Güncelleme öncesi "Önizleme / Simülasyon" (Dry-run) özelliğini eklemek.

### I. Test Edilmeden Production'a Alınmaması Gereken Noktalar
* Çok parçalı VBA makroları içeren `.xlsm` dosyalarının güncellendikten sonra hatasız derlenmesi (`Compile VBAProject`).
* Veri Modeli ve Pivot Tablo içeren dosyaların güncelleme sonrası veri yenileyebilmesi.
* Elektrik kesintisi simülasyonunda (process kill) orijinal dosyanın sağlam kaldığının doğrulanması.
* Salt-okunur ve şifreli dosyalarda sistemin çökmeden zarif hata üretebilmesi.

### J. Senden Sonraki Adımda İstemem Gereken İşlem
Bu aşamada analiz eksiksiz tamamlanmıştır ve tek bir kod satırına dokunulmamıştır. Bir sonraki adımda benden talep etmeniz gereken işlem:
> **"PHASE 1 (Veri Güvenliği) ve PHASE 2 (Excel Bütünlüğü) adımlarını kapsayan; atomik dosya değiştirme (`File.Replace`), regex kelime sınırı (`\b`), `$conn.Name` koruması ve manuel hesaplama modu eklemelerini içeren uygulama planını (Implementation Plan) hazırla."**
