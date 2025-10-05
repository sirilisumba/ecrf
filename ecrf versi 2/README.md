# ----- Script ECRF Versi 2 ----- 
#
# UPDATE:
# 1. Looping non-stop, baik dari form ke form, maupun dari data ke data.
# 2. Break hanya terjadi jika: 
    a. di stop uesr: ctrl + c atau ctrl + x
    b. row 3 di data.xlsx kosong
    c. server &/ wifi lemot, sehingga script break
# 3. Setiap error, ditandai dengan suara beep 3x
#
# CATATAN:
# 1. Jalankan versi 2, hanya jika, anda yakin dengan file excel sudah benar semua.
# 2. Jika terjadi break, lanjutkan dengan script per form:
    a. form1.py
    b. form2.py
    c. form3.py
    d. formpelapor.py
    e. formpasien.py
    f. formvaksinasi.py
    g. formkipi.py
# 3. File data.xlsx harus satu folder derngan script dan chromedriver.
# 4. Cara jalankan: ketik di cmd: py ecrfv2-1.py
# 5. Jika mulai dari file kedua, maka cara panggil di cmd: py ecrfv2-2.py
#
# TROUBLESHOOT:
# Jika koneksi dari cmd dan brave menggunakan command ini tidak berhasil:
    "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe" --remote-debugging-port=9222
# Gunakan command berikut ini:
    "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe" --remote-debugging-port=9222 --user-data-dir="C:\braveprofile"
# Lalu cek dengan cara ketik di tab baru di Brave:
    http://127.0.0.1:9222/json
# Jika keluar seperti ini:
    {
     "description": "",
     "devtoolsFrontendUrl": "https://chrome-devtools-frontend.appspot.com/serve_rev/@6c9b7bdded46e59d445cb0c067bff9f3bcd8fdd/inspector.html?ws=127.0.0.1:9222/devtools/page/5191F8FE88C9F90F29C5C4E855F4C134",
     "id": "5191F8FE88C9F90F29C5C4E855F4C134",
     "title": "New Tab",
     "type": "page",
     "url": "http://127.0.0.1:9222/json",
     "webSocketDebuggerUrl": "ws://127.0.0.1:9222/devtools/page/5191F8FE88C9F90F29C5C4E855F4C134"
    },
# maka cmd dan brave sudah saling terhubung, dan proses scripting sudah bisa dimulai.
