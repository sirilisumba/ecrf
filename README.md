# ecrf
Copying data into from excel into website

==============================
======== README FIRST ======== 
==============================
Sebelum mulai menjalankan script, pastikan hal ini sudah dilakukan:
1. Download dan Install python 
2. Install selenium: pip install selenium. Dan install openpyxl: pip install openpyxl.
3. Install Brave (rekomendasi: Chrome) Bisa pakai web browser yg lain.
4. Webdriver browser sudah di copy ke folder tempat script disimpan. Chrome/Chromium/Brave: chromedriver. Download di: https://developer.chrome.com/docs/chromedriver/downloads
5. Pastikan versi ChromeDriver cocok dengan versi Chrome/Brave
6. Jalankan terlebih dulu: "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe" --remote-debugging-port=9222 
7. Jika pake chrome, tidak perlu jalankan no 6.
8. Pastikan rename data yg akan dipakai dengan nama: data.xlsx dan simpan di folder yang sama dengan script.
9. Pastikan data yg di copy-paste sudah ke data.xlsx sudah sesuai.
10. Buka cmd masuk ke folder dimana script tersimpan, jalankan: py ecrf1.py

Catatan:
1. Buka form ke-2 harus ambil dari data-responden-id, karena no inklusi belum di proses
2. Form ke-3 ambil value data-responden-id yg sudah didapat
3. Form ke 4-7 ambil dari no inklusi yang didapat dari form ke-3
4. Jika data ke geser ke halaman berikutnya, harus buka dulu halamannya
5. Looping pencarian data-responden-id, saat mau open form ke-2, hanya sampai row 4 saja, dg asumsi 4 pengisi ecrf aktif isi bersamaan
6. Data: no inklusi, inisial, tgl lahir, dan haspengobatan, diambil dari 1 data aja. 
7. Tgl informed consent dan tgl enrollment, meski selalu sama, tetap dibedakan
8. Jika setelah idle lama (meski tidak di shutdown), browser tdk berreaksi, lakukan step di no 6
Bugs yang BEULM solve:
1. Kalo dibuat manual, script masih bisa jalan, meski user belum pencet SIMPAN
2. Jika manual, pastikan klik SIMPAN dulu, baru tap ENTER di keyboard
3. Jika server loading terlalu lambat, script akan sering break
==============================
==============================
