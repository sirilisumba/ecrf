# VERIFIKASI
# coding ini hanya berlaku untuk pengguna Windows

#
# PERSIAPAN:
# 1. Copy/download semua files, jadikan dalam satu folder
# 2. Siapkan 2 file excel: data.xlsx dan log.xlsx
# 3. File data.xlsx, isikan dengan data yang mau dicek (lihat contoh excel)
# 4. File log.xlsx, tidak perlu diisi apa-apa. Gunanya: untuk mencatat perbedaan data yang ditemukan
# 5. Download web browser: Brave https://brave.com/download/ ikuti petunjuk download (JANGAN UBAH LOKASI INSTALASI, IKUTI SAJA PETUNJUKNYA).

#
# PENGGUNAAN:
# 1. Buka folder, double klik file: start.exe, console command prompt (CMD) akan otomatis terbuka (JANGAN DITUTUP).
# 2. Browser Brave akan terubkan otomatis. Masukkan username dan password, lalu simpan jika ditawarkan oleh browser utk menyimpan username & password.
# 3. Masuk ke website Biofarma, hingga posisi tabel responden.
# 4. Kembali ke console, pencet ENTER di keyboard (ikuti petunjuk di console).
# 5. Aplikasi akan mulai membandingkan dengan data di excel.
# 6. Jika ditemukan data tidak sama, akan ada bunyi alarm, dengan posisi kotak comment sudah terbuka.
# 7. Setelah pencet tombol SIMPAN di kotak comment, pencet ENTER di console.
# 8. Jangan pencet apapun di web browser. Hanya interaksi dengan console.
# 9. Jika pada saat hendak mengisikan comment, ingin melihat nilai di web, kotak comment bisa ditutup dulu. Tapi harus kembali ke posisi terakhir.
# 10. Jika pada saat running aplikasi, lalu putus di tengah. Kembali ke halaman tabel, lalu jalankan sesuai aplikasi sesuai dengan posisi terakhir. Jika terhenti di Form Pelapor, maka selanjutnya dapat menjalankan verifpasien.exe
# 11. Selama brave terbuka, tidak perlu mulai dari Start.exe. Namun jika brave tertutup, selalu mulai dari file start.exe.
# 12. Pastikan kedua excel dalam keadaan tertutup sebelum mulai menjalankan aplikasi.
# 13. Untuk menghentikan jalannya aplikasi, bisa dengan menekan ctrl + C di console, atau langsung saja tutup console nya.

#
# LOGIC:
# 1. Aplikasi akan membandingkan data di data.xlsx dengan data di web.
# 2. Perbedaan data akan dicatat di log.xlsx.
# 3. Aplikasi akan selalu membaca di baris 3. Selesai satu data, baris 3 akan di hapus, dan baris dibawahnya naik ke baris 3, begitu seterusnya.
# 4. File start.exe, hanya dijalankan pada saat awal. Lalu loop akan berjalan dari verif.exe hingga verifkipi.exe dan kembali ke verif.exe.
# 5. Template comment sudah disiapkan, namun bisa diganti sesuai kebutuhan, lalu klik tombol SIMPAN.
# 6. Tanggal Informed Consent dan Tanggal Enrollment dibatasi, start of date nya: 13 September 2025, sebelum dari tanggal itu, dianggap out-of-range.
# 7. Aturan penulisan pada FASYANKES, NO BATCH, dan INISIAL, sudah sesuai aturan penulisan dari Komnas KIPI.
# 8. Pada form KIPI, jika ada kesalahan input di Lokal dan Sistemik, dan console terus membaca, biarkan saja. Karena yg di hide oleh web, tetap diproses oleh aplikasi.
# 9. Setiap ada error, aplikasi akan memberikan bunyi alarm yg berbeda dari bunyi alarm comment.

#
# TROUBLESHOOT:
# 1. Untuk memastikan apakah browser Brave sudah konek dengan aplikasi: buka tab baru di brave, ketikan url ini: http://127.0.0.1:9222/json
# 2. Jika keluar warning: "This site can’t be reached", berarti Brave belum berhasil konek. Selain itu berarti Brave sudah konek.
