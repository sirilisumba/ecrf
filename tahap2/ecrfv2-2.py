from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import time
from openpyxl import load_workbook
import time
import subprocess
from ecrfv2_utils import create_driver, EXCEL_PATH, buka_form, load_excel_data, set_text, isi_time, set_radio, set_checkbox, save_form, play_sound, isi_datepicker

driver, wait_long, wait_short = create_driver()
data = load_excel_data()
# file_path = "data2.xlsx"
wb = load_workbook(EXCEL_PATH)
ws = wb.active

try:
    ###############################################
    ######### FORM 4 - INFORMASI PELAPOR ##########
    ###############################################

    ############## FORM 4 - BUKA FORM #############
    print(f"📘 No. inklusi dari Excel:", data["val_no_inklusi"])
    print("Mulai mengisi Form Pelapor")
    buka_form(driver, wait_long, data["val_no_inklusi"], "Informasi Pelapor")

    ############## FORM 4 - ISI DATA ##############
    # Data dari Excel
    print("Mulai mengisi Form Pelapor")
    print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_provinsi'], data['val_tgl_lapor'], data['val_hasPengobatan'])

    # Isi form
    set_text(driver, "itemid_58832", data['val_no_inklusi'])
    set_text(driver, "itemid_58833", data['val_inisial'])

    # Radio button Provinsi
    try:
        if data['val_provinsi'] is not None and str(data['val_provinsi']).strip() != "":
            css = f"#form_group_58835 input[type='radio'][value='{data['val_provinsi']}']"
            
            radio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, css))
            )
            radio.click() 
            print("→ form_group_58835 =", data['val_provinsi'])
    except Exception as e:
        print("❌ GAGAL select radio form_group_58835:", e)
        play_sound()

    isi_datepicker(driver, wait_short, "itemid_58836", data['val_tgl_lapor'])
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    print("✅ Copy-paste INFORMASI PELAPOR: DONE.")
    

    ############## FORM 4 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user ENTER di Keyboard
    # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
    # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")

    ###############################################
    ######### FORM 5 - INFORMASI PASIEN ###########
    ###############################################

    ############## FORM 5 - BUKA FORM #############
    print("Mulai mengisi Form Pasien")
    print(f"📘 No. inklusi dari Excel:", data["val_no_inklusi"])
    buka_form(driver, wait_long, data["val_no_inklusi"], "Informasi Pasien")

    ############## FORM 5 - ISI DATA ##############
    # Data dari Excel
    print("Mulai mengisi Form Pelapor")
    print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_jeniskelamin'], data['val_tgllahir'], data['val_usia_thn'], data['val_usia_bln'], data['val_usia_hr'], data['val_hasPengobatan'])

    # Isi form
    set_text(driver, "itemid_58337", data['val_no_inklusi'])
    set_text(driver, "itemid_58338", data['val_inisial'])
    set_radio(driver, "form_group_58340", data['val_jeniskelamin'])

    try:
        if data['val_jeniskelamin'] is not None and str(data['val_jeniskelamin']).strip() != "":
            css = f"#form_group_58340 input[type='radio'][value='{data['val_jeniskelamin']}']"
            radio = driver.find_element(By.CSS_SELECTOR, css)
            driver.execute_script("arguments[0].click();", radio)  # JS click to avoid intercepts
            print("→ form_group_58340 =", data['val_jeniskelamin'])
    except Exception as e:
        print("❌ GAGAL to select radio form_group_58340:", e)
        play_sound()

    isi_datepicker(driver, wait_short, "itemid_59060", data['val_tgllahir'])
    set_text(driver, "itemid_59061", data['val_usia_thn'])
    set_text(driver, "itemid_59062", data['val_usia_bln'])
    set_text(driver, "itemid_59063", data['val_usia_hr'])
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 5 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user ENTER di Keyboard
    # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
    # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")

    ###############################################
    ########## FORM 6 - DATA VAKSINASI ############
    ###############################################

    ############## FORM 6 - BUKA FORM #############
    print("Mulai mengisi Form Data Vaksinasi")
    print(f"📘 No. inklusi dari Excel:", data["val_no_inklusi"])
    buka_form(driver, wait_long, data["val_no_inklusi"], "Data Vaksinasi")

    ############## FORM 6 - ISI DATA ##############
    # Data dari Excel
    print("Mulai mengisi Form Pelapor")
    print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_jenis_vaksin'], data['val_manufaktur'], data['val_no_batch'], data['val_dosis'], data['val_tgl_vaksin'], data['val_wkt_vaksin'], data['val_tempat_vaksin'], data['val_vaksin_lain'], data['val_vaksin_lain1'], data['val_vaksin_lain2'], data['val_vaksin_lain3'], data['val_tgl_vaksin'], data['val_haspengobatan'])

    # Isi form
    set_text(driver, "itemid_60292", data['val_no_inklusi'])
    set_text(driver, "itemid_60293", data['val_inisial'])
    set_radio(driver, "form_group_60294", data['val_jenis_vaksin'])
    set_radio(driver, "form_group_60295", data['val_manufaktur'])
    set_text(driver, "itemid_60296", data['val_no_batch'])
    set_radio(driver, "form_group_60297", data['val_dosis'])
    isi_datepicker(driver, wait_short, field_id="itemid_60298", tanggal_obj=data['val_tgl_vaksin'])
    isi_time(driver, wait_short, "itemid_60299", data['val_wkt_vaksin'])
    set_radio(driver, "form_group_60301", data['val_tempat_vaksin'])
    # vaksin lain, ada conditional:
    try:
        applied = False
        if data['val_vaksin_lain'] is not None and str(data['val_vaksin_lain']).strip() != "":
            if set_radio(driver, "form_group_60302", data['val_vaksin_lain']):
                applied = True
                time.sleep(0.5)
                selected_val = str(data['val_vaksin_lain']).strip()
                if selected_val in ["1"]:
                    set_text(driver, "itemid_60303", data['val_vaksin_lain1'])
                    set_text(driver, "itemid_60304", data['val_vaksin_lain2'])
                    set_text(driver, "itemid_60305", data['val_vaksin_lain3'])
                else:
                    print(f"→ form_group_60302 = {selected_val}; tidak ada field tambahan yang perlu diisi")
        if not applied:
            print("→ form_group_60302 tidak diset (tidak ada value atau gagal memilih radio)")
            play_sound()
    except Exception as e:
        print("❌ Terjadi error saat memproses form_group_60302:", e)
        play_sound()
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 6 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user ENTER di Keyboard
    # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
    # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")


    ###############################################
    ################ FORM 7 - KIPI ################
    ###############################################

    ############## FORM 7 - BUKA FORM #############
    print("Mulai mengisi Form KIPI")
    print(f"📘 No. inklusi dari Excel:", data["val_no_inklusi"])
    print("📘 Pengobatan:", data["val_hasPengobatan"])
    buka_form(driver, wait_long, data["val_no_inklusi"], "KIPI")

    ############## FORM 7 - ISI DATA ##############
    # Data dari Excel
    print("Mulai mengisi Form KIPI")
    print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_kategori'], data['val_lokal'], 
          data['val_nyeri'], data['val_tebal'], data['val_bengkak'], data['val_lokal_lain'], 
          data['val_lokal_lain1'], data['val_lokal_lain1_nama'], 
          data['val_lokal_lain2'], data['val_lokal_lain2_nama'], 
          data['val_lokal_lain3'], data['val_lokal_lain3_nama'], 
          data['val_lokal_lain4'], data['val_lokal_lain4_nama'], 
          data['val_sistemik'], data['val_demam'], data['val_rewel'], data['val_nangis'], data['val_sistemik_lain'], 
          data['val_sistemik_lain_1'], data['val_sistemik_lain_1_nama'],
          data['val_sistemik_lain_2'], data['val_sistemik_lain_2_nama'], 
          data['val_sistemik_lain_3'], data['val_sistemik_lain_3_nama'], 
          data['val_sistemik_lain_4'], data['val_sistemik_lain_4_nama'],
          data['val_kondisi_akhir'], data['val_hasPengobatan'])
    
    # Isi form
    wait_short.until(EC.presence_of_element_located((By.ID, "itemid_58646")))
    set_text(driver, "itemid_58646", data['val_no_inklusi'])
    set_text(driver, "itemid_58647", data['val_inisial'])

    if set_radio(driver, "form_group_58650", data['val_kategori']):
        # only handle nested block if value == 2
        try:
            if str(data['val_kategori']).strip() == "2":  
                time.sleep(0.35)  
                # form_group_58766 >>> REAKSI LOKAL
                if set_radio(driver, "form_group_58766", data['val_lokal']):
                    if str(data['val_lokal']).strip() == "1":
                        time.sleep(0.25)
                        set_radio(driver, "form_group_58767", data['val_nyeri'])
                        set_radio(driver, "form_group_58768", data['val_merah'])
                        set_radio(driver, "form_group_58769", data['val_tebal'])
                        set_radio(driver, "form_group_58770", data['val_bengkak'])
                        if set_radio(driver, "form_group_58782", data['val_lokal_lain']):
                            if str(data['val_lokal_lain']).strip() == "1":
                                time.sleep(0.25)
                                if set_radio(driver, "form_group_59084", data['val_lokal_lain1']):
                                    if str(data['val_lokal_lain1']).strip() == "1":
                                        set_text(driver, "itemid_59088", data['val_lokal_lain1_nama']) 
                                if set_radio(driver, "form_group_59085", data['val_lokal_lain2']):
                                    if str(data['val_lokal_lain2']).strip() == "1":
                                        set_text(driver, "itemid_59092", data['val_lokal_lain2_nama']) 
                                if set_radio(driver, "form_group_59086", data['val_lokal_lain3']):
                                    if str(data['val_lokal_lain3']).strip() == "1":
                                        set_text(driver, "itemid_59095", data['val_lokal_lain3_nama']) 
                                if set_radio(driver, "form_group_59087", data['val_lokal_lain4']):
                                    if str(data['val_lokal_lain4']).strip() == "1":
                                        set_text(driver, "itemid_59098", data['val_lokal_lain4_nama']) 
                #form_group_58781 >>> SISTEMIK
                if set_radio(driver, "form_group_58781", data['val_sistemik']):
                    if str(data['val_sistemik']).strip() == "1":
                        time.sleep(0.25)
                        set_radio(driver, "form_group_58795", data['val_demam'])
                        set_radio(driver, "form_group_58796", data['val_rewel'])
                        set_radio(driver, "form_group_58797", data['val_nangis'])
                        if set_radio(driver, "form_group_58807", data['val_sistemik_lain']):
                            if str(data['val_sistemik_lain']).strip() == "1":
                                time.sleep(0.25)
                                if set_radio(driver, "form_group_59129", data['val_sistemik_lain_1']):
                                    if str(data['val_sistemik_lain_1']).strip() == "1":
                                        set_text(driver, "itemid_59113", data['val_sistemik_lain_1_nama']) 
                                if set_radio(driver, "form_group_59130", data['val_sistemik_lain_2']):
                                    if str(data['val_sistemik_lain_2']).strip() == "1":
                                        set_text(driver, "itemid_59116", data['val_sistemik_lain_2_nama']) 
                                if set_radio(driver, "form_group_59131", data['val_sistemik_lain_3']):
                                    if str(data['val_sistemik_lain_3']).strip() == "1":
                                        set_text(driver, "itemid_59119", data['val_sistemik_lain_3_nama']) 
                                if set_radio(driver, "form_group_59132", data['val_sistemik_lain_4']):
                                    if str(data['val_sistemik_lain_4']).strip() == "1":
                                        set_text(driver, "itemid_59122", data['val_sistemik_lain_4_nama']) 
                # DL2 -> form_group_58827 >> SEMBUH
                set_radio(driver, "form_group_58827", data['val_kondisi_akhir'])
        except Exception as e:
            print("❌ GAGAL tangani error dari nested radio:", e)
            play_sound()
    else:
        print("→ main radio form_group_58650 tidak di set (atau kosong)")

    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 7 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user ENTER di Keyboard
    # PAUSE: tunggu Anda klik SAVE manual di browser lalu tekan ENTER di terminal
    # input("👉 Waiting for SIMPAN button is clicked, then click ENTER at keyboard to continue...")

    ###############################################
    ############# LANJUT DELETE ROW 3 #############
    ###############################################
    
    ws.delete_rows(3)
    wb.save(EXCEL_PATH)

    print("✅ Delete baris 3 di Excel: DONE.")

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

    # >>>> Lanjut ke file Loop1 secara otomatis <<<<
    subprocess.run(["python", "ecrfv2-1.py"])

except KeyboardInterrupt:
    print("\n⏹️ Proses dihentikan user.")

except Exception as e:
    print(f"\n❌ An unhandled exception occurred: {e}")
    play_sound()
    exit() 


