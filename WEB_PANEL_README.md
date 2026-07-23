# Hepsiburada Fiyat Takip Web Paneli

Bu panel, gömülü 114 ürünün Hepsiburada fiyat taramasını tek düğmeyle başlatır,
ilerlemeyi canlı gösterir ve tamamlanan Excel raporunu indirmeye açar.

## Panelin sundukları

- Harici Excel yüklemeden, koddaki statik ürün tablosuyla çalışma
- Aynı anda yalnızca bir rapor çalıştırma
- Canlı ilerleme, bulunan/bulunamayan/hata sayaçları ve çalışma günlüğü
- Parolalı erişim
- Tamamlanan `.xlsx` raporunu tarayıcıdan indirme
- Docker ve Railway üzerinde çalışma

## Yerelde Docker ile çalıştırma

Docker Desktop kuruluysa proje klasöründe:

```bash
docker build -t hb-fiyat-paneli .
docker run --init --ipc=host -p 8080:8080 \
  -e APP_PASSWORD="guclu-bir-parola" \
  -e APP_SESSION_SECRET="uzun-rastgele-bir-anahtar" \
  hb-fiyat-paneli
```

Ardından `http://localhost:8080` adresini açın.

## Railway'e yayınlama

1. Bu klasörü özel bir GitHub deposuna gönderin.
2. Railway'de **New Project → Deploy from GitHub repo** seçeneğini açın.
3. Depoyu seçin. Railway kökteki `Dockerfile` dosyasını kullanır.
4. Servisin **Variables** bölümüne şunları ekleyin:

   - `APP_PASSWORD`: panel giriş parolanız
   - `APP_SESSION_SECRET`: en az 32 karakterlik rastgele bir değer
   - `SCRAPER_WORKERS`: `4`

5. Deploy tamamlanınca **Settings → Networking → Generate Domain** ile adres üretin.

Playwright tarayıcıları bellek kullandığı için en az 1 GB RAM önerilir. Sunucu yeniden
başlatıldığında eski rapor dosyaları silinebilir; her çalışmadan sonra raporu indirmeniz
önerilir.

## Ayarlar

| Değişken | Varsayılan | Açıklama |
|---|---:|---|
| `APP_PASSWORD` | boş | Boşsa parola ekranı kapatılır |
| `APP_SESSION_SECRET` | rastgele | Oturum imza anahtarı |
| `SCRAPER_WORKERS` | `4` | Paralel ürün tarama sayısı |
| `REPORT_OUTPUT_DIR` | `generated_reports` | Excel çıktı klasörü |
| `PLAYWRIGHT_BROWSER_CHANNEL` | `chrome` | Docker'da boş bırakılır |
| `SCRAPER_EXTRA_ARGS` | boş | Yalnızca test için ek scraper parametreleri |

## Dosya yapısı

- `hepsiburada_price.py`: mevcut fiyat tarama motoru
- `hb_web_panel/app.py`: API, parola, iş kuyruğu ve Excel indirme
- `hb_web_panel/static/`: tarayıcı arayüzü
- `Dockerfile`: Playwright uyumlu sunucu imajı
- `railway.json`: Railway sağlık kontrolü ve dağıtım ayarları
