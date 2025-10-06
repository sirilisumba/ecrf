from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime
from selenium.common.exceptions import TimeoutException
import time
import subprocess
from ecrfv2_utils import create_driver, load_excel_data, set_text, isi_radio_button, isi_date_indo, set_radio, save_form, play_sound, save_form1


def jalankan_script(attempt=1, max_attempts=3):
    if attempt > max_attempts:
        print("🚫 Terlalu banyak percobaan. Hentikan script.")
        play_sound()
        return

    driver = None
    try:
        # -- bagian utama automasi --
        driver, wait_long, wait_short = create_driver()
        driver.get("https://aplikasi.biofarma.co.id/ecrf/Project/InfoUjiKlinis/1961")

        try:
            data = load_excel_data()
        except Exception as e:
            print(e)
            play_sound()
            exit()

        try:
            ###############################################
            ######### FORM 1 - ADD DATA RESPONDEN #########
            ###############################################

            ############## FORM 1 - BUKA FORM #############
            try:
                print("⏳ ...menunggu data dimuat dan spinner hilang...")
                wait_long.until(EC.invisibility_of_element_located((By.CLASS_NAME, "spinner")))

                # Tunggu tombol muncul, bisa diklik, dan enabled
                button = wait_long.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#add_responden_puskesmas button")))
                wait_long.until(lambda d: button.is_enabled())

                # Scroll dan klik pakai JS
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.3)  # biar scroll stabil
                driver.execute_script("arguments[0].click();", button)
                print("✅ Loading selesai dan button diklik.")

            except TimeoutException:
                print("⚠️ Timeout: Spinner tidak hilang dalam batas waktu.")
            except Exception as e:
                print("❌ Gagal scroll atau klik tombol:", e)
                play_sound()

            # Tunggu modal muncul dan siap diisi ==> PENTING KARENA SERING GAGAL
            try:
                wait_responden = WebDriverWait(driver, 15)  # kasih timeout agak lama
                modal = wait_responden.until(
                    EC.visibility_of_element_located((By.ID, "myModalTambahResponden"))
                )

                wait_responden.until(
                    EC.element_to_be_clickable((By.ID, "puskesmas"))
                )
                print("✅ Modal dan form siap diisi.")

            except TimeoutException:
                print("⚠️ Timeout menunggu modal tambah responden muncul.")
            except Exception as e:
                print("❌ Gagal menunggu modal:", e)
                play_sound()

            ############## FORM 1 - ISI DATA ##############
            # Data dari Excel
            print("Mulai mengisi Form Add Data Responden")
            print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
            print("Data dari Excel:", data['val_puskesmas'], data['val_inisial'], data['val_tgllahir'], data['val_no_inklusi'], data['val_jeniskelamin'], data['val_tglscreening'])
            
            # Mulai isi form
            # WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "puskesmas"))).click()
            # isi_dropdown(driver, wait_long, "puskesmas",  data['val_puskesmas'])
            select_puskesmas = Select(wait_long.until(EC.presence_of_element_located((By.ID, "puskesmas"))))
            select_puskesmas.select_by_visible_text(data['val_puskesmas'])

            set_text(driver, "inisial_nama", data['val_inisial'])
            isi_date_indo(driver, wait_short, "tanggalLahir_s",  data['val_tgllahir'])
            isi_radio_button(driver, "jenis_kelamin",  data['val_jeniskelamin'])
            isi_date_indo(driver, wait_short, "tanggalScreening_s",  data['val_tglscreening'])

            ############## FORM 1 - SAVE DATA ##############
            # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
            save_form1(driver, button_id="btn-add-respondent")

            # Save OPSI 2 : tunggu user ENTER di Keyboard
            # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
            # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")

            ###############################################
            ######### FORM 2 - INKLUSI / EKSKLUSI #########
            ###############################################

            ############## FORM 2 - BUKA FORM #############
            time.sleep(5)
            
            # Konversi dari datetime ke string terlebih dahulu
            if isinstance(data['val_tgllahir'], datetime):
                val_tgllahir_str = data['val_tgllahir'].strftime("%d %B %Y")
            else:
                val_tgllahir_str = str(data['val_tgllahir'])

            def ganti_bulan_ke_inggris(tanggal_str):
                bulan_map = {
                    'Januari': 'January',
                    'Februari': 'February',
                    'Maret': 'March',
                    'April': 'April',
                    'Mei': 'May',
                    'Juni': 'June',
                    'Juli': 'July',
                    'Agustus': 'August',
                    'September': 'September',
                    'Oktober': 'October',
                    'November': 'November',
                    'Desember': 'December',
                }

                for indo, eng in bulan_map.items():
                    if indo in tanggal_str:
                        return tanggal_str.replace(indo, eng)
                return tanggal_str  
            
            inisial_nama = data['val_inisial'].strip()
            tanggalLahir_s = ganti_bulan_ke_inggris(val_tgllahir_str)
            responden_id = None
            rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-responden-id]")

            for i, tr in enumerate(rows[:4]):
                try:
                    nama_td = tr.find_element(By.CSS_SELECTOR, "td.nama_responden")
                    nama = nama_td.text.strip()

                    if nama != inisial_nama:
                        print(f"⏭️ Baris {i+1}: name not match ({nama})")
                        continue

                    # Klik tombol fa-file (lihat detail) untuk ambil tanggal lahir
                    detail_button = tr.find_element(By.CSS_SELECTOR, "button.btn-detail i.fa-file")
                    driver.execute_script("arguments[0].scrollIntoView(true);", detail_button)
                    time.sleep(0.3)
                    detail_button.click()
                    print("🔍 Buka modal untuk mendapatkan data tanggal lahir")

                    # Tunggu modal muncul
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "myModalTambahResponden"))
                    )

                    # Ambil value tanggal lahir
                    input_tanggal = driver.find_element(By.ID, "tanggalLahir_s")
                    tanggal_lahir_modal = input_tanggal.get_attribute("value").strip()

                    # Tutup modal langsung
                    try:
                        close_button = driver.find_element(By.ID, "btn-close")
                        close_button.click()
                    except:
                        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                    print(f"📆 Birth date from modal: {tanggal_lahir_modal}")

                    format_tgl = "%d %B %Y"  # tetap pakai ini

                    try:
                        tgl1_str = ganti_bulan_ke_inggris(tanggal_lahir_modal)
                        tgl2_str = ganti_bulan_ke_inggris(tanggalLahir_s)

                        tgl1 = datetime.strptime(tgl1_str, format_tgl)
                        tgl2 = datetime.strptime(tgl2_str, format_tgl)

                        if tgl1 != tgl2:
                            print(f"⏭️ Birth date is not matched ({tanggal_lahir_modal}) vs ({tanggalLahir_s})")
                            continue
                    except ValueError as e:
                        print(f"⚠️ FAILED parsing on date: {e}")
                        continue

                    # Kalo cocok, ambil responden_id
                    responden_id = tr.get_attribute("data-responden-id")
                    print(f"📌 FOUND data-responden-id: {responden_id}")

                    # Klik chevron untuk buka detail
                    chevron_css = f'tr[data-responden-id="{responden_id}"] td[class^="detail-row"] i.fa-chevron-down'
                    chevron = wait_long.until(EC.element_to_be_clickable((By.CSS_SELECTOR, chevron_css)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", chevron)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", chevron)
                    print("🔽 Click chevron to expand...")

                    # Tunggu detail row muncul
                    detail_selector = f"tr.detail_row{i}"
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, detail_selector))
                    )
                    print(f"✅ Detail row {i+1} terbuka.")

                    # Klik tombol ISI
                    try:
                        tombol_isi = wait_long.until(EC.element_to_be_clickable((
                            By.XPATH,
                            "//tr[td[contains(., 'Inklusi/Eksklusi')]]//button[contains(text(), 'ISI')]"
                        )))
                        print("✅ Tombol 'ISI' ketemu")
                        print("   OnClick attr:", tombol_isi.get_attribute("onclick"))

                        driver.execute_script("arguments[0].scrollIntoView(true);", tombol_isi)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", tombol_isi)
                        print("🖱️ Tombol 'ISI' dklik.")
                    except Exception as e:
                        print("❌ GAGAL menemukan atau klik tombol 'ISI'")
                        play_sound()
                        print(e)
                    break
                except Exception as e:
                    print(f"⚠️ Error pada baris {i+1}: {e}")
                    continue
            if not responden_id:
                print("❌ GAGAL menemukan responden yang cocok pada baris.")
                play_sound()

            print("📌 Responden_id utk buka Form 2:", responden_id)
            ############## FORM 2 - ISI DATA ##############
            # Data dari Excel
            print("Mulai mengisi Form Inklusi/Eksklusi")
            print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
            print("Data dari Excel:", data['val_radio1'], data['val_radio2'], data['val_radio3'])

            # Mulai isi form
            set_radio(driver, "form_group_60312", data['val_radio1'])
            set_radio(driver, "form_group_60313", data['val_radio2'])
            set_radio(driver, "form_group_60314", data['val_radio3'])

            print("✅ Copy-paste INKLUSI/EKSKLUSI: DONE.")

            ############## FORM 2 - SAVE DATA ##############
            # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
            save_form(driver, button_id="btn-submit")

            # Save OPSI 2 : tunggu user ENTER di Keyboard
            # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
            # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")

            ###############################################
            ######### FORM 3 - TAMBAH NO INKLUSI ##########
            ###############################################

            ############## FORM 3 - BUKA FORM #############

            print("Responden_id dari form sebelumnya:", responden_id)

            if responden_id:
                responden_id = str(responden_id) 
                try:
                    print(f"🔍 Mencari baris dengan responden_id = {responden_id}")

                    # Ambil semua baris (maksimal 4 baris awal saja)
                    rows = wait_long.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr[data-responden-id]")))
                    rows = rows[:4]

                    tr_elem = None
                    for i, row in enumerate(rows):
                        rid = row.get_attribute("data-responden-id")
                        print(f"⏳ Cek baris {i+1}: data-responden-id = {rid}")
                        if rid == responden_id:
                            tr_elem = row
                            print(f"✅ Baris ke-{i+1} cocok! data-responden-id = {rid}")
                            break

                    if not tr_elem:
                        raise Exception(f"❌ Tidak ditemukan baris dengan data-responden-id = {responden_id}")
                        play_sound()

                    # Klik ikon chevron-down untuk expand
                    chevron = tr_elem.find_element(By.CSS_SELECTOR, "i.fa-chevron-down")
                    wait_long.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'tr[data-responden-id="{responden_id}"] i.fa-chevron-down')))
                    driver.execute_script("arguments[0].scrollIntoView(true);", chevron)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", chevron)
                    print("✅ Chevron diklik untuk buka detail.")

                    time.sleep(1)  # beri waktu detail muncul

                    # Cari dan klik tombol "ENROLL"
                    tombol_enroll = wait_long.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//tr[td[contains(., 'Inklusi/Eksklusi')]]//button[contains(text(), 'ENROLL')]"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView(true);", tombol_enroll)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", tombol_enroll)
                    print("🖱️ Tombol 'ENROLL' berhasil diklik.")

                except Exception as e:
                    raise Exception(f"❌ GAGAL klik tombol ENROLL: {e}")
                    play_sound()

            else:
                print("❌ GAGAL mendapatkan responden_id. Fungsi ENROLL dibatalkan.")
                play_sound()

            ############## FORM 3 - ISI DATA ##############
            wait_short.until(EC.visibility_of_element_located((By.ID, "myModalTambahNomorInklusi")))

            # Data dari Excel
            isi_date_indo(driver, wait_short, "tanggal_inklusi_s", data['val_tgl_inklusi'])
            set_text(driver, "nomor_inklusi", data['val_no_inklusi'])
            # set_text(driver, "inisial_responden", data['val_inisial'])
            # set_text(driver, "keterangan", data['val_keterangan'])

            ############## FORM 3 - SAVE DATA ##############
            # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
            save_form(driver, button_id="btn-add-nomor-inklusi")
            
            # Save OPSI 2 : tunggu user ENTER di Keyboard
            # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
            # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")


            ###############################################
            ########## LANJUT KE FORM BERIKUTNYA ##########
            ###############################################

            # jawaban = input("➡️  Lanjut data berikutnya? (Y/N): ").strip().lower()
            # if jawaban == 'y':
            #     print("▶️  Next data...")
            # elif jawaban == 'n':
            #     print("⏹️ Proses dihentikan user.")
            #     exit()
            # else:
            #     print("⚠️  Unrecognized. Please tap Y or N on your keyboard.")

            # # >>>> Lanjut ke file Loop1 secara otomatis <<<<
            subprocess.run(["python", "ecrfv1-2.py"])
            
        except KeyboardInterrupt:
            print("\n⏹️ Proses dihentikan user.")

        except Exception as e:
            print(f"\n❌ An unhandled exception occurred: {e}")
            play_sound()
            exit() 

        # # 🧪 Simulasi error untuk testing
        # raise Exception("Simulasi error")
        print("✅ Data berhasil diisi dan disimpan.")
        play_sound()
        driver.quit()

    except Exception as e:
        print(f"❌ Percobaan ke-{attempt} gagal: {e}")
        play_sound()

        if driver:
            try:
                # ✅ Cek apakah modal tambah responden terbuka
                modal = WebDriverWait(driver, 2).until(
                    EC.visibility_of_element_located((By.ID, "myModalTambahResponden"))
                )
                print("🛑 Modal masih terbuka.")
                play_sound()

                # ✅ Klik tombol cancel
                cancel_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#footeraddedit .btn__cancel"))
                )
                cancel_btn.click()
                print("🧹 Tombol 'Batal' diklik untuk tutup modal.")
            except Exception as inner:
                print(f"⚠️ Modal tidak terbuka atau gagal klik tombol Cancel: {inner}")
                play_sound()

            # try:
            #     play_sound()
            #     driver.quit()
            # except:
            #     pass
            try:
                if driver:  # cek dulu driver sudah dibuat
                    driver.quit()
            except Exception as e:
                print(f"⚠️ Gagal quit driver: {e}")

        # 🔁 Coba lagi
        print("🔄 Restarting script...")
        time.sleep(1)
        jalankan_script(attempt + 1)

if __name__ == "__main__":
    jalankan_script()

