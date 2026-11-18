# PPTX Türkçe Çeviri

Windows'ta bir PowerPoint sunumundaki düzenlenebilir İngilizce metni yerel bir modelle Türkçeye çevirir. Metin kutularının koordinatları ve boyutları korunur. Betik ayrı bir PPTX ve atlanan öğeleri listeleyen bir CSV raporu oluşturur.

## Gereksinimler

- Windows ve 64 bit CPython. Paket uyumluluğu için Python 3.12 ile başlayabilirsin.
- [pptx_turkce.py](pptx_turkce.py)
- [Argos İngilizce → Türkçe model paketi](https://github.com/iremyilmazi/pptx-turkce/releases/download/argos-en-tr-1.5/translate-en_tr-1_5.argosmodel)
- Python paketleri `lxml`, `ctranslate2` ve `sentencepiece`

İnternet, Python paketlerinin ve modelin ilk indirmesi için gerekir. Çeviri yerel CPU'da çalışır. PowerPoint yalnızca sonucu açıp kontrol etmek için gerekir.

## 1. Dosyaları hazırla

Örneğin Belgeler altında `pptx-turkce` adlı bir klasör oluştur. Betiği ve model paketini bu klasöre koy. Modelin `.argosmodel` uzantısını koru ve paketi açmadan bırak.

Komut İstemi'ni (CMD) aç ve klasöre geç. Klasör farklı bir konumdaysa aşağıdaki yolu değiştir. `cd /d` gerekirse sürücüyü de değiştirir.

```cmd
cd /d "%USERPROFILE%\Documents\pptx-turkce"
python --version
python -c "import struct; print(struct.calcsize('P') * 8)"
```

Son komutun `64` yazması gerekir. `python` komutu bulunamaz ama `py -3.12 --version` çalışırsa sanal ortamı `py -3.12 -m venv .venv` ile oluştur.

## 2. Python paketlerini yükle

Proje klasöründe bir sanal ortam oluştur. Komutlar sanal ortamın Python dosyasını doğrudan çağırır, böylece sanal ortamı etkinleştirmene gerek kalmaz.

```cmd
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install lxml ctranslate2 sentencepiece
.\.venv\Scripts\python.exe -c "import lxml.etree, ctranslate2, sentencepiece; print('Paketler hazır')"
```

### İnternete kapalı şirket bilgisayarında kurulum

İnternet erişimi olan, aynı Windows mimarisine ve Python sürümüne sahip bir bilgisayarda paketleri indir.

```cmd
python -m pip download --only-binary=:all: --dest .\wheels lxml ctranslate2 sentencepiece
```

`wheels` klasörünü, betiği ve model paketini kurumun izin verdiği yöntemle şirket bilgisayarına aktar. Sunum şirket bilgisayarında kalabilir. Orada proje klasöründen şu komutları çalıştır.

```cmd
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --no-index --find-links .\wheels lxml ctranslate2 sentencepiece
```

## 3. Sunumu çevir

Önce sunumu PowerPoint'te kaydet. Aşağıdaki örnek yolu kendi dosyanın yoluyla değiştir. Yol boşluk içeriyorsa tırnakları koru.

```cmd
.\.venv\Scripts\python.exe .\pptx_turkce.py "C:\Dosyalar\Sunum.pptx" --model ".\translate-en_tr-1_5.argosmodel"
```

Kaynak sunumun yanında `Sunum_tr.pptx` ve `Sunum_tr.atlananlar.csv` oluşturulur. Başka bir klasöre yazmak için zaten var olan bir klasörde yeni bir dosya adı belirt.

```cmd
.\.venv\Scripts\python.exe .\pptx_turkce.py "C:\Dosyalar\Sunum.pptx" --model ".\translate-en_tr-1_5.argosmodel" --cikti "C:\Dosyalar\Sunum_Turkce_v2.pptx"
```

## 4. Sonucu kontrol et

- Çıktıyı PowerPoint'te aç. Türkçe metnin uzadığı başlıkları ve dar kutuları gözden geçir. Metin kutusunun yeri ve boyutu korunurken satır kayması veya taşma olabilir.
- CSV raporunu Excel'de aç. Alanlar noktalı virgülle ayrılır. Rapor kaynak metnin kendisi yerine PPTX içindeki XML parçasını, paragraf sırasını ve atlama nedenini gösterir. `slideN.xml` paket içindeki dosya adıdır ve PowerPoint'teki güncel slayt sırasından farklı olabilir.
- Sayı değişikliği, kaybolan kısaltma, boş sonuç, olağandışı uzunluk veya model hatası tespit edildiğinde kaynak metin korunur. Bu filtreler anlam doğruluğunu garanti etmez. Terimleri ve kritik ifadeleri gözden geçir.
- Farklı karakter biçimlerini bir arada kullanan paragraflar, dinamik alanlar ve elle satır sonu verilen paragraflar atlanabilir. Görsel içi yazılar, SmartArt, konuşmacı notları ve veri kaynağına bağlı grafik etiketleri kapsam dışındadır.

## Sorun giderme

| Belirti | Yapılacak işlem |
| --- | --- |
| `python` tanınmıyor | `py -3.12 --version` ile Python başlatıcısını kontrol et veya kurumun sağladığı Python yolunu kullan. |
| `No matching distribution found` | Python sürümünü ve 64 bit mimariyi kontrol et. Çevrimdışı kurulumda wheel dosyalarının hedef Python ve Windows mimarisiyle eşleştiğinden emin ol. |
| `ModuleNotFoundError` | Paketleri betiği çalıştırdığın Python ile yükle. Yukarıdaki `.venv\Scripts\python.exe` yolunu kullan. |
| `DLL load failed` | 64 bit Python kullandığını kontrol et. Kurum BT ekibinden CTranslate2 için gerekli Microsoft Visual C++ çalışma zamanını doğrulamasını iste. |
| Model dosyası bulunamıyor | CMD'nin doğru klasörde olduğunu ve dosyanın gerçek adının `.argosmodel` ile bittiğini kontrol et. Gerekirse modelin tam yolunu yaz. |
| Model başlangıç kontrolü başarısız | Model dosyasının tam indirildiğini, belirtilen İngilizce → Türkçe paketi olduğunu ve Python bağımlılıklarının yüklenebildiğini kontrol et. |
| Çıktı veya rapor yolu kullanımda | `--cikti` ile yeni bir dosya adı seç. Betik kaynak sunumu ve mevcut çıktıları korur. |
| `PermissionError` | Çıktı klasöründe yazma iznini kontrol et. Gerekirse `--cikti` ile yazabildiğin bir klasör seç. |

Komut seçeneklerini görmek için şu komutu çalıştır.

```cmd
.\.venv\Scripts\python.exe .\pptx_turkce.py --help
```
