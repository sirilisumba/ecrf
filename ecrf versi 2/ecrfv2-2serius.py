from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import time
from openpyxl import load_workbook
import time
import subprocess
from ecrfv2_utils import create_driver, buka_form, load_excel_data, set_text, isi_time, set_radio, set_checkbox, save_form, play_sound, isi_datepicker

driver, wait_long, wait_short = create_driver()
data = load_excel_data()
file_path = "data.xlsx"
wb = load_workbook(file_path)
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
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_kategori'], data['val_lokal'], data['val_nyeri'],
          data['val_lokal_tgl'], data['val_lokal_wkt'], data['val_lokal_hr'], data['val_merah'], data['val_merah_tgl'], 
          data['val_merah_wkt'], data['val_merah_hr'], data['val_tebal'], data['val_tebal_tgl'], data['val_tebal_wkt'], data['val_tebal_hr'],
          data['val_bengkak'], data['val_bengkak_tgl'], data['val_bengkak_wkt'], data['val_bengkak_hr'], data['val_lokal_lain'], data['val_lokal_lain1'],
          data['val_lokal_lain1_nama'], data['val_lokal_lain1_tgl'], data['val_lokal_lain1_wkt'], data['val_lokal_lain1_hr'], data['val_lokal_lain2'],
          data['val_lokal_lain2_nama'], data['val_lokal_lain2_tgl'], data['val_lokal_lain2_wkt'], data['val_lokal_lain2_hr'], data['val_lokal_lain3'], data['val_inisial'], 
          data['val_lokal_lain3_nama'], data['val_lokal_lain3_tgl'], data['val_lokal_lain3_wkt'], data['val_lokal_lain3_hr'], data['val_lokal_lain4'], data['val_inisial'], 
          data['val_lokal_lain4_nama'], data['val_lokal_lain4_tgl'], data['val_lokal_lain4_wkt'], data['val_lokal_lain4_hr'], data['val_sistemik'], data['val_inisial'], 
          data['val_demam'], data['val_demam_tgl'], data['val_demam_wkt'], data['val_demam_hr'], data['val_rewel'], data['val_rewel_tgl'], 
          data['val_rewel_wkt'], data['val_rewel_hr'], data['val_nangis'], data['val_nangis_tgl'], data['val_nangis_wkt'], data['val_nangis_hr'], 
          data['val_sistemik_lain'], data['val_sistemik_lain_1'], data['val_sistemik_lain_1_nama'], data['val_sistemik_lain_1_tgl'], 
          data['val_sistemik_lain_1_wkt'], data['val_sistemik_lain_1_hr'], data['val_sistemik_lain_2'], data['val_sistemik_lain_2_nama'], 
          data['val_sistemik_lain_2_tgl'], data['val_sistemik_lain_2_wkt'], data['val_sistemik_lain_2_hr'], data['val_sistemik_lain_3'],
          data['val_sistemik_lain_3_nama'], data['val_sistemik_lain_3_tgl'], data['val_sistemik_lain_3_wkt'], data['val_sistemik_lain_3_hr'], 
          data['val_sistemik_lain_4'], data['val_sistemik_lain_4_nama'], data['val_sistemik_lain_4_tgl'], data['val_sistemik_lain_4_wkt'],
          data['val_sistemik_lain_4_hr'], data['val_diagnosis_1'], data['val_diagnosis_2'], data['val_diagnosis_3'], data['val_kondisi_akhir'], 
          data['val_kausalitas'], data['val_hasPengobatan'], 
          )
    
    # Isi form
    wait_short.until(EC.presence_of_element_located((By.ID, "itemid_58646")))
    set_text(driver, "itemid_58646", data['val_no_inklusi'])
    set_text(driver, "itemid_58647", data['val_inisial'])

    if set_radio(driver, "form_group_58650", data['val_kategori']):
        # only handle nested block if value == 1
        try:
            if str(data['val_kategori']).strip() == "1": # 
                time.sleep(0.35)  # wait UI to render
                # AR2 -> form_group_58704 >>> REAKSI LOKAL
                if set_radio(driver, "form_group_58704", data['val_lokal']):
                    if str(data['val_lokal']).strip() == "1":
                        time.sleep(0.25)
                        # AS2 -> form_group_58705 >>> NYERI LOKAL
                        set_radio(driver, "form_group_58705", data['val_nyeri'])
                        if set_radio(driver, "form_group_58705", data['val_nyeri']):
                            if str(data['val_nyeri']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58706", data['val_lokal_tgl']) # AT2 (tgl)
                                isi_time(driver, wait_short, "itemid_58707", data['val_lokal_wkt']) # AU2 (wkt)
                                set_text(driver, "itemid_58837", data['val_lokal_hr']) # AV2 (hari)
                        # AW2 -> form_group_58709 >>> KEMERAHAN
                        set_radio(driver, "form_group_58709", data['val_merah'])
                        if set_radio(driver, "form_group_58709", data['val_merah']):
                            if str(data['val_merah']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58710", data['val_merah_tgl']) # AX2 (tgl)
                                isi_time(driver, wait_short, "itemid_58711", data['val_merah_wkt']) # AY2 (wkt)
                                set_text(driver, "itemid_58838", data['val_merah_hr']) # AZ2 (hari)
                        # BA2 -> form_group_58713 >>> PENEBALAN
                        set_radio(driver, "form_group_58713", data['val_tebal'])
                        if set_radio(driver, "form_group_58713", data['val_tebal']):
                            if str(data['val_tebal']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58714", data['val_tebal_tgl']) # BB2 (tgl)
                                isi_time(driver, wait_short, "itemid_58715", data['val_tebal_wkt']) # BC2 (wkt)
                                set_text(driver, "itemid_58839", data['val_tebal_hr']) # BD2 (hari)
                        # BE2 -> form_group_58717 >>> PEMBENGKAKAN
                        set_radio(driver, "form_group_58717", data['val_bengkak'])
                        if set_radio(driver, "form_group_58717", data['val_bengkak']):
                            if str(data['val_bengkak']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58718", data['val_bengkak_tgl']) # BF2 (tgl)
                                isi_time(driver, wait_short, "itemid_58719", data['val_bengkak_wkt']) # BG2 (wkt)
                                set_text(driver, "itemid_58840", data['val_bengkak_hr']) # BH2 (hari)
                        # BI2 -> form_group_58722 >>> LAIN-LAIN
                        if set_radio(driver, "form_group_58722", data['val_lokal_lain']):
                            if str(data['val_lokal_lain']).strip() == "1":
                                time.sleep(0.25)
                                # BJ2 -> form_group_59064 >>> lain-lain 1
                                if set_radio(driver, "form_group_59064", data['val_lokal_lain1']):
                                    if str(data['val_lokal_lain1']).strip() == "1":
                                        set_text(driver, "itemid_59068", data['val_lokal_lain1_nama']) # BK2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59069", data['val_lokal_lain1_tgl']) # BL2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59070", data['val_lokal_lain1_wkt']) # BM2 (wkt)
                                        set_text(driver, "itemid_59071", data['val_lokal_lain1_hr']) # BN2 (hari)
                                # BO2 -> form_group_59065 >>> lain-lain 2
                                if set_radio(driver, "form_group_59065", data['val_lokal_lain2']):
                                    if str(data['val_lokal_lain2']).strip() == "1":
                                        set_text(driver, "itemid_59072", data['val_lokal_lain2_nama']) # BP2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59073", data['val_lokal_lain2_tgl']) # BQ2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59074", data['val_lokal_lain2_wkt']) # BR2 (wkt)
                                        set_text(driver, "itemid_59075", data['val_lokal_lain2_hr']) # BS2 (hari)
                                # BT2 -> form_group_59066 >>> lain-lain 3
                                if set_radio(driver, "form_group_59066", data['val_lokal_lain3']):
                                    if str(data['val_lokal_lain3']).strip() == "1":
                                        set_text(driver, "itemid_59076", data['val_lokal_lain3_nama']) # BU2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59077", data['val_lokal_lain3_tgl']) # BV2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59078", data['val_lokal_lain3_wkt']) # BW2 (wkt)
                                        set_text(driver, "itemid_59079", data['val_lokal_lain3_hr']) # BX2 (hari)
                                # BY2 -> form_group_59067 >>> lain-lain 4
                                if set_radio(driver, "form_group_59067", data['val_lokal_lain4']):
                                    if str(data['val_lokal_lain4']).strip() == "1":
                                        set_text(driver, "itemid_59080", data['val_lokal_lain4_nama']) # BZ2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59081", data['val_lokal_lain4_tgl']) # CA2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59082", data['val_lokal_lain4_wkt']) # CB2 (wkt)
                                        set_text(driver, "itemid_59083", data['val_lokal_lain4_hr']) # CC2 (hari)

                # CD2 -> form_group_58721 >>> SISTEMIK
                if set_radio(driver, "form_group_58721", data['val_sistemik']):
                    if str(data['val_sistemik']).strip() == "1":
                        time.sleep(0.25)
                        # CE2 -> form_group_58735 >>> DEMAM
                        set_radio(driver, "form_group_58735", data['val_demam'])
                        if set_radio(driver, "form_group_58735", data['val_demam']):
                            if str(data['val_demam']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58736", data['val_demam_tgl']) # CF2 (tgl)
                                isi_time(driver, wait_short, "itemid_58737", data['val_demam_wkt']) # CG2 (wkt)
                                set_text(driver, "itemid_58842", data['val_demam_hr']) # CH2 (hari)
                        # CI2 -> form_group_58739 >>> REWEL
                        set_radio(driver, "form_group_58739", data['val_rewel'])
                        if set_radio(driver, "form_group_58739", data['val_rewel']):
                            if str(data['val_rewel']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58741", data['val_rewel_tgl']) # CF2 (tgl)
                                isi_time(driver, wait_short, "itemid_58742", data['val_rewel_wkt']) # CG2 (wkt)
                                set_text(driver, "itemid_58743", data['val_rewel_hr']) # CH2 (hari)
                        # CM2 -> form_group_58740 >>> NANGIS
                        set_radio(driver, "form_group_58740", data['val_nangis'])
                        if set_radio(driver, "form_group_58740", data['val_nangis']):
                            if str(data['val_nangis']).strip() == "1":
                                isi_datepicker(driver, wait_short, "itemid_58744", data['val_nangis_tgl']) # CF2 (tgl)
                                isi_time(driver, wait_short, "itemid_58805", data['val_nangis_wkt']) # CG2 (wkt)
                                set_text(driver, "itemid_58844", data['val_nangis_hr']) # CH2 (hari)
                        # CQ2 -> form_group_58746 >>> LAIN-LAIN
                        if set_radio(driver, "form_group_58746", data['val_sistemik_lain']):
                            if str(data['val_sistemik_lain']).strip() == "1":
                                time.sleep(0.25)
                                # CR2 -> form_group_59125 >>> lain-lain 1
                                if set_radio(driver, "form_group_59125", data['val_sistemik_lain_1']):
                                    if str(data['val_sistemik_lain_1']).strip() == "1":
                                        set_text(driver, "itemid_59101", data['val_sistemik_lain_1_nama']) # CS2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59102", data['val_sistemik_lain_1_tgl']) # CT2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59133", data['val_sistemik_lain_1_wkt']) # CU2 (wkt)
                                        set_text(driver, "itemid_59134", data['val_sistemik_lain_1_hr']) # CV2 (hari)
                                # CW2 -> form_group_59126 >>> lain-lain 2
                                if set_radio(driver, "form_group_59126", data['val_sistemik_lain_2']):
                                    if str(data['val_sistemik_lain_2']).strip() == "1":
                                        set_text(driver, "itemid_59104", data['val_sistemik_lain_2_nama']) # CX2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59105", data['val_sistemik_lain_2_tgl']) # CY2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59135", data['val_sistemik_lain_2_wkt']) # CZ2 (wkt)
                                        set_text(driver, "itemid_59136", data['val_sistemik_lain_2_hr']) # DA2 (hari)
                                # DB2 -> form_group_59127 >>> lain-lain 3
                                if set_radio(driver, "form_group_59127", data['val_sistemik_lain_3']):
                                    if str(data['val_sistemik_lain_3']).strip() == "1":
                                        set_text(driver, "itemid_59107", data['val_sistemik_lain_3_nama']) # DC2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59108", data['val_sistemik_lain_3_tgl']) # DD2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59137", data['val_sistemik_lain_3_wkt']) # DE2 (wkt)
                                        set_text(driver, "itemid_59138", data['val_sistemik_lain_3_hr']) # DF2 (hari)
                                # DG2 -> form_group_59128 >>> lain-lain 4
                                if set_radio(driver, "form_group_59128", data['val_sistemik_lain_4']):
                                    if str(data['val_sistemik_lain_4']).strip() == "1":
                                        set_text(driver, "itemid_59110", data['val_sistemik_lain_4_nama']) # DH2 (nama)
                                        isi_datepicker(driver, wait_short, "itemid_59111", data['val_sistemik_lain_4_tgl']) # DI2 (tgl)
                                        isi_time(driver, wait_short, "itemid_59139", data['val_sistemik_lain_4_wkt']) # DJ2 (wkt)
                                        set_text(driver, "itemid_59140", data['val_sistemik_lain_4_hr']) # DK2 (hari)
                
                set_text(driver, "itemid_58759", data['val_diagnosis_1'])
                set_text(driver, "itemid_58760", data['val_diagnosis_2'])
                set_text(driver, "itemid_58761", data['val_diagnosis_3'])
                set_radio(driver, "form_group_58827", data['val_kondisi_akhir'])
               
                set_radio(driver, "form_group_58763", data['val_kausalitas'])
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
    wb.save(file_path)

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
    subprocess.run(["python", "yoni1.py"])

except KeyboardInterrupt:
    print("\n⏹️ Proses dihentikan user.")

except Exception as e:
    print(f"\n❌ An unhandled exception occurred: {e}")
    play_sound()
    exit() 



