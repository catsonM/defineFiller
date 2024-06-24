# DefineFiller

DefineFiller, Excel dosyalarını işleyerek eksik tanımları tamamlayan ve çıktıları düzenli bir şekilde oluşturan bir Python uygulamasıdır. Bu uygulama, modül ve FD (Function Definitions) tanımlarını tamamlar ve eksik define değerlerini doldurur.

## Özellikler

- Eksik modül ve FD tanımlarını tamamlar
- Eksik define değerlerini default değerlerle doldurur
- Kullanıcı dostu grafik arayüz
- Çıktı dosyasının başına 'afill_' ibaresini ekler
- PI no. ve Beta Version bilgilerini dosya adına ekler (opsiyonel)

## Gereksinimler

- Python 3.6 veya üzeri
- Aşağıdaki Python kütüphaneleri:
  - pandas
  - openpyxl
  - tkinter

## Kurulum

1. Gerekli kütüphaneleri yükleyin:
    ```sh
    pip install pandas openpyxl
    ```

2. Depoyu klonlayın:
    ```sh
    git clone https://github.com/catsonM/defineFiller.git
    cd defineFiller
    ```

## Kullanım

1. Uygulamayı başlatmak için:
    ```sh
    python definefiller.py
    ```

2. Kullanıcı arayüzünde:
    - 'Browse' butonuna tıklayarak işlemek istediğiniz raw Excel dosyasını seçin.
    - PI no. ve Beta Version bilgilerini girin (opsiyonel).
    - Çıktı klasörünü seçin (opsiyonel).
    - 'Save in subfolder' seçeneğini işaretleyerek çıktı dosyasını alt klasörde saklayabilirsiniz (opsiyonel).
    - 'Complete' butonuna tıklayarak işlemi başlatın.

## Yardım

Program, belirli bir formatta raw Excel dosyalarını tamamlanmış dosyalara dönüştürür. Aşağıdaki adımları takip ederek programı kullanabilirsiniz:

1. 'Browse' butonuna tıklayarak raw Excel dosyasını seçin.
2. PI no. ve Beta Version bilgilerini ilgili alanlara girin. Bu alanlar opsiyoneldir, ancak boş bırakıldığında kullanıcıya bir uyarı verilir ve devam edip etmeyeceği sorulur.
3. Opsiyonel olarak çıktı klasörünü seçin.
4. 'Save in subfolder' seçeneğini işaretleyerek çıktı dosyasını alt klasörde saklayabilirsiniz.
5. 'Complete' butonuna tıklayarak işlemi başlatın.

Not: Çıktı dosyasının başına 'afill_' ibaresi eklenecektir.

## Yazar

- **Mert Can Catoglu**
- E-posta: mertcan.catoglu@tr.bosch.com

## Lisans

Bu proje MIT Lisansı altında lisanslanmıştır - daha fazla bilgi için `LICENSE` dosyasına bakınız.
