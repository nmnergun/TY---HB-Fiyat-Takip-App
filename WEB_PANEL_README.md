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
3. Depoyu seçin. Railway kökteki `railway.json`, `railpack.json`,
   `requirements.txt` ve `main.py` dosyalarını kullanır.
4. Servisin **Variables** bölümüne şunları ekleyin:

   - `APP_PASSWORD`: panel giriş parolanız
   - `APP_SESSION_SECRET`: en az 32 karakterlik rastgele bir değer
   - `SCRAPER_WORKERS`: `4`

5. Deploy tamamlanınca **Settings → Networking → Generate Domain** ile adres üretin.

Playwright tarayıcıları bellek kullandığı için en az 1 GB RAM önerilir. Sunucu yeniden
başlatıldığında eski rapor dosyaları silinebilir; her çalışmadan sonra raporu indirmeniz
önerilir.

## Railway Railpack/build hataları

Güncel paket doğrudan Railpack ile çalışacak şekilde yapılandırılmıştır. Railpack,
`requirements.txt` içindeki Playwright bağımlılığını algılayarak Chromium ve gerekli
Linux paketlerini kurar.

1. GitHub deposunun kökünde şu dosyaların doğrudan bulunduğunu kontrol edin:
   `railway.json`, `railpack.json`, `requirements.txt`, `main.py`,
   `hepsiburada_price.py`.
   Dosyalar bir ZIP'in ya da ikinci bir klasörün içinde kalmamalıdır.
2. Railway servisinde **Settings → Build → Builder** değerini **Railpack** seçin.
3. **Settings → Deploy → Custom Start Command** alanını temizleyin. Depodaki
   `railway.json` doğru Uvicorn komutunu tanımlar.
4. Daha önce eklediyseniz `RAILWAY_DOCKERFILE_PATH` değişkenini kaldırın.
5. **Redeploy → Clear build cache and deploy** ile yeniden dağıtın.

Docker kullanmak isterseniz projedeki `Dockerfile` ve `start.sh` alternatif olarak
hazır tutulmuştur.

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
- `main.py`: Railway'in algıladığı FastAPI giriş noktası
- `requirements.txt`: Railpack Python/Playwright bağımlılıkları
- `railpack.json`: Railpack sağlayıcı ve başlangıç ayarları
- `hb_web_panel/app.py`: API, parola, iş kuyruğu ve Excel indirme
- `hb_web_panel/static/`: tarayıcı arayüzü
- `start.sh`: Railway/Docker başlangıç betiği
- `Dockerfile`: Playwright uyumlu sunucu imajı
- `railway.json`: Railway sağlık kontrolü ve dağıtım ayarları
