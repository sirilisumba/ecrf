from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import time
from openpyxl import load_workbook
import time
import subprocess
from ecrf_utils import create_driver, buka_form, load_excel_data, isi_textinput, isi_date, isi_time, set_radio, set_checkbox, save_form

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
    buka_form(driver, wait_long, data["val_no_inklusi"], "Informasi Pelapor")

    ############## FORM 4 - ISI DATA ##############
    # Data dari Excel
    print("Mulai mengisi Form Pelapor")
    print(f"📘 No. inklusi dari Excel: {data['val_no_inklusi']}")
    print("Data dari Excel:", data['val_no_inklusi'], data['val_inisial'], data['val_provinsi'], data['val_tgl_lapor'], data['val_hasPengobatan'])

    # Isi form
    isi_textinput(driver, wait_short, "itemid_58832", data['val_no_inklusi'])
    isi_textinput(driver, wait_short, "itemid_58833", data['val_inisial'])

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

    isi_date(driver, wait_short, "itemid_58836", data['val_tgl_lapor'])
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    print("✅ Copy-paste INFORMASI PELAPOR: DONE.")
    

    ############## FORM 4 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    # save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user klik SIMPAN di browser lalu tekan ENTER di terminal
    input("👉 Menunggu tombol SIMPAN diklik, dan pencet ENTER di terminal...")

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
    isi_textinput(driver, wait_short, "itemid_58337", data['val_no_inklusi'])
    isi_textinput(driver, wait_short, "itemid_58338", data['val_inisial'])
    set_radio(driver, "form_group_58340", data['val_jeniskelamin'])

    try:
        if data['val_jeniskelamin'] is not None and str(data['val_jeniskelamin']).strip() != "":
            css = f"#form_group_58340 input[type='radio'][value='{data['val_jeniskelamin']}']"
            radio = driver.find_element(By.CSS_SELECTOR, css)
            driver.execute_script("arguments[0].click();", radio)  # JS click to avoid intercepts
            print("→ form_group_58340 =", data['val_jeniskelamin'])
    except Exception as e:
        print("❌ GAGAL to select radio form_group_58340:", e)

    isi_date(driver, wait_short, "itemid_59060", data['val_tgllahir'])
    isi_textinput(driver, wait_short, "itemid_59061", data['val_usia_thn'])
    isi_textinput(driver, wait_short, "itemid_59062", data['val_usia_bln'])
    isi_textinput(driver, wait_short, "itemid_59063", data['val_usia_hr'])
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 5 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    # save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user klik SIMPAN di browser lalu tekan ENTER di terminal
    input("👉 Menunggu tombol SIMPAN diklik, dan pencet ENTER di terminal...")

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
    isi_textinput(driver, wait_short, "itemid_60292", data['val_no_inklusi'])
    isi_textinput(driver, wait_short, "itemid_60293", data['val_inisial'])
    set_radio(driver, "form_group_60294", data['val_jenis_vaksin'])
    set_radio(driver, "form_group_60295", data['val_manufaktur'])
    isi_textinput(driver, wait_short, "itemid_60296", data['val_no_batch'])
    set_radio(driver, "form_group_60297", data['val_dosis'])
    isi_date(driver, wait_short, "itemid_60298", data['val_tgl_vaksin'])
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
                    isi_textinput(driver, wait_short, "itemid_60303", data['val_vaksin_lain1'])
                    isi_textinput(driver, wait_short, "itemid_60304", data['val_vaksin_lain2'])
                    isi_textinput(driver, wait_short, "itemid_60305", data['val_vaksin_lain3'])
                else:
                    print(f"→ form_group_60302 = {selected_val}; tidak ada field tambahan yang perlu diisi")
        if not applied:
            print("→ form_group_60302 tidak diset (tidak ada value atau gagal memilih radio)")
    except Exception as e:
        print("❌ Terjadi error saat memproses form_group_60302:", e)
    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 6 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    # save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user klik SIMPAN di browser lalu tekan ENTER di terminal
    input("👉 Menunggu tombol SIMPAN diklik, dan pencet ENTER di terminal...")


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
          data['val_sistemik_lain_4_hr'], data['val_kondisi_akhir'], data['val_hasPengobatan'])
    
    # Isi form
    # wait_short.until(EC.presence_of_element_located((By.ID, "itemid_58646")))
    isi_textinput(driver, wait_short, "itemid_58646", data['val_no_inklusi'])
    isi_textinput(driver, wait_short, "itemid_58647", data['val_inisial'])

    if set_radio(driver, "form_group_58650", data['val_kategori']):
        # hanya jika nested block value == 2
        try:
            if str(data['val_kategori']).strip() == "2": # 
                time.sleep(0.35)  

                # AR2 -> form_group_58766 >>> REAKSI LOKAL
                if set_radio(driver, "form_group_58766", data['val_lokal']):
                    if str(data['val_lokal']).strip() == "1":
                        time.sleep(0.25)
                        # AS2 -> form_group_58767 >>> NYERI LOKAL
                        set_radio(driver, "form_group_58767", data['val_nyeri'])
                        if set_radio(driver, "form_group_58767", data['val_nyeri']):
                            if str(data['val_nyeri']).strip() == "1":
                                isi_date(driver, wait_short, "itemid_58772", data['val_lokal_tgl']) # AT2 (tgl)
                                isi_time(driver, wait_short, "itemid_58773", data['val_lokal_wkt']) # AU2 (wkt)
                                isi_textinput(driver, wait_short, "itemid_58846", data['val_lokal_hr']) # AV2 (hari)
                        # AW2 -> form_group_58768 >>> KEMERAHAN
                        set_radio(driver, "form_group_58768", data['val_merah'])
                        if set_radio(driver, "form_group_58768", data['val_merah']):
                            if str(data['val_merah']).strip() == "1":
                                isi_date(driver, wait_short, "itemid_58775", data['val_merah_tgl']) # AX2 (tgl)
                                isi_time(driver, wait_short, "itemid_58799", data['val_merah_wkt']) # AY2 (wkt)
                                isi_textinput(driver, wait_short, "itemid_58847", data['val_merah_hr']) # AZ2 (hari)
                        # BA2 -> form_group_58769 >>> PENEBALAN
                        set_radio(driver, "form_group_58769", data['val_tebal'])
                        if set_radio(driver, "form_group_58769", data['val_tebal']):
                            if str(data['val_tebal']).strip() == "1":
                                isi_date(driver, wait_short, "itemid_58777", data['val_tebal_tgl']) # BB2 (tgl)
                                isi_time(driver, wait_short, "itemid_58801", data['val_tebal_wkt']) # BC2 (wkt)
                                isi_textinput(driver, wait_short, "itemid_58848", data['val_tebal_hr']) # BD2 (hari)
                        # BE2 -> form_group_58770 >>> PEMBENGKAKAN
                        set_radio(driver, "form_group_58770", data['val_bengkak'])
                        if set_radio(driver, "form_group_58770", data['val_bengkak']):
                            if str(data['val_bengkak']).strip() == "1":
                                isi_date(driver, wait_short, "itemid_58779", data['val_bengkak_tgl']) # BF2 (tgl)
                                isi_time(driver, wait_short, "itemid_58803", data['val_bengkak_wkt']) # BG2 (wkt)
                                isi_textinput(driver, wait_short, "itemid_58849", data['val_bengkak_hr']) # BH2 (hari)
                        # BI2 -> form_group_58782 >>> LAIN-LAIN
                        if set_radio(driver, "form_group_58782", data['val_lokal_lain']):
                            if str(data['val_lokal_lain']).strip() == "1":
                                time.sleep(0.25)
                                # BJ2 -> form_group_59084 >>> lain-lain 1
                                if set_radio(driver, "form_group_59084", data['val_lokal_lain1']):
                                    if str(data['val_lokal_lain1']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59088", data['val_lokal_lain1_nama']) # BK2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59089", data['val_lokal_lain1_tgl']) # BL2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59090", data['val_lokal_lain1_wkt']) # BM2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59091", data['val_lokal_lain1_hr']) # BN2 (hari)
                                # BO2 -> form_group_59085 >>> lain-lain 2
                                if set_radio(driver, "form_group_59085", data['val_lokal_lain2']):
                                    if str(data['val_lokal_lain2']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59092", data['val_lokal_lain2_nama']) # BP2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59093", data['val_lokal_lain2_tgl']) # BQ2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59149", data['val_lokal_lain2_wkt']) # BR2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59150", data['val_lokal_lain2_hr']) # BS2 (hari)
                                # BT2 -> form_group_59086 >>> lain-lain 3
                                if set_radio(driver, "form_group_59086", data['val_lokal_lain3']):
                                    if str(data['val_lokal_lain3']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59095", data['val_lokal_lain3_nama']) # BU2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59096", data['val_lokal_lain3_tgl']) # BV2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59151", data['val_lokal_lain3_wkt']) # BW2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59152", data['val_lokal_lain3_hr']) # BX2 (hari)
                                # BY2 -> form_group_59087 >>> lain-lain 4
                                if set_radio(driver, "form_group_59087", data['val_lokal_lain4']):
                                    if str(data['val_lokal_lain4']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59098", data['val_lokal_lain4_nama']) # BZ2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59099", data['val_lokal_lain4_tgl']) # CA2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59153", data['val_lokal_lain4_wkt']) # CB2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59154", data['val_lokal_lain4_hr']) # CC2 (hari)

                # CD2 -> form_group_58781 >>> SISTEMIK
                if set_radio(driver, "form_group_58781", data['val_sistemik']):
                    if str(data['val_sistemik']).strip() == "1":
                        time.sleep(0.25)
                        # CE2 -> form_group_58795 >>> DEMAM
                        set_radio(driver, "form_group_58795", data['val_demam'])
                        # if set_radio(driver, "form_group_58795", data['val_demam']):
                        #     if str(data['val_demam']).strip() == "1":
                        #         isi_date(driver, wait_short, "itemid_58808", data['val_demam_tgl']) # CF2 (tgl)
                        #         isi_time(driver, wait_short, "itemid_58809", data['val_demam_wkt']) # CG2 (wkt)
                        #         isi_textinput(driver, wait_short, "itemid_58851", data['val_demam_hr']) # CH2 (hari)
                        # CI2 -> form_group_58796 >>> REWEL
                        set_radio(driver, "form_group_58796", data['val_rewel'])
                        # if set_radio(driver, "form_group_58796", data['val_rewel']):
                        #     if str(data['val_rewel']).strip() == "1":
                        #         isi_date(driver, wait_short, "itemid_58811", data['val_rewel_tgl']) # CF2 (tgl)
                        #         isi_time(driver, wait_short, "itemid_58828", data['val_rewel_wkt']) # CG2 (wkt)
                        #         isi_textinput(driver, wait_short, "itemid_58852", data['val_rewel_hr']) # CH2 (hari)
                        # CM2 -> form_group_58797 >>> NANGIS
                        set_radio(driver, "form_group_58797", data['val_nangis'])
                        # if set_radio(driver, "form_group_58797", data['val_nangis']):
                        #     if str(data['val_nangis']).strip() == "1":
                        #         isi_date(driver, wait_short, "itemid_58813", data['val_nangis_tgl']) # CF2 (tgl)
                        #         isi_time(driver, wait_short, "itemid_58830", data['val_nangis_wkt']) # CG2 (wkt)
                        #         isi_textinput(driver, wait_short, "itemid_58853", data['val_nangis_hr']) # CH2 (hari)
                        # CQ2 -> form_group_58807 >>> LAIN-LAIN
                        if set_radio(driver, "form_group_58807", data['val_sistemik_lain']):
                            if str(data['val_sistemik_lain']).strip() == "1":
                                time.sleep(0.25)
                                # CR2 -> form_group_59129 >>> lain-lain 1
                                if set_radio(driver, "form_group_59129", data['val_sistemik_lain_1']):
                                    if str(data['val_sistemik_lain_1']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59113", data['val_sistemik_lain_1_nama']) # CS2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59114", data['val_sistemik_lain_1_tgl']) # CT2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59141", data['val_sistemik_lain_1_wkt']) # CU2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59142", data['val_sistemik_lain_1_hr']) # CV2 (hari)
                                # CW2 -> form_group_59130 >>> lain-lain 2
                                if set_radio(driver, "form_group_59130", data['val_sistemik_lain_2']):
                                    if str(data['val_sistemik_lain_2']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59116", data['val_sistemik_lain_2_nama']) # CX2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59117", data['val_sistemik_lain_2_tgl']) # CY2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59143", data['val_sistemik_lain_2_wkt']) # CZ2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59144", data['val_sistemik_lain_2_hr']) # DA2 (hari)
                                # DB2 -> form_group_59131 >>> lain-lain 3
                                if set_radio(driver, "form_group_59131", data['val_sistemik_lain_3']):
                                    if str(data['val_sistemik_lain_3']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59119", data['val_sistemik_lain_3_nama']) # DC2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59120", data['val_sistemik_lain_3_tgl']) # DD2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59145", data['val_sistemik_lain_3_wkt']) # DE2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59146", data['val_sistemik_lain_3_hr']) # DF2 (hari)
                                # DG2 -> form_group_59132 >>> lain-lain 4
                                if set_radio(driver, "form_group_59132", data['val_sistemik_lain_4']):
                                    if str(data['val_sistemik_lain_4']).strip() == "1":
                                        isi_textinput(driver, wait_short, "itemid_59122", data['val_sistemik_lain_4_nama']) # DH2 (nama)
                                        # isi_date(driver, wait_short, "itemid_59123", data['val_sistemik_lain_4_tgl']) # DI2 (tgl)
                                        # isi_time(driver, wait_short, "itemid_59147", data['val_sistemik_lain_4_wkt']) # DJ2 (wkt)
                                        # isi_textinput(driver, wait_short, "itemid_59148", data['val_sistemik_lain_4_hr']) # DK2 (hari)
                # DL2 -> form_group_58827 >> SEMBUH
                set_radio(driver, "form_group_58827", data['val_kondisi_akhir'])
        except Exception as e:
            print("❌ GAGAL tangani error dari nested radio:", e)
    else:
        print("→ main radio form_group_58650 tidak di set (atau kosong)")

    set_checkbox(driver, "hasPengobatan", data['val_hasPengobatan'])

    ############## FORM 7 - SAVE DATA #############
    # Save OPSI 1 : otomatis save, lanjut ke script berikutnya
    # save_form(driver, button_id="btn-submit")

    # Save OPSI 2 : tunggu user klik SIMPAN di browser lalu tekan ENTER di terminal
    input("👉 Menunggu tombol SIMPAN diklik, dan pencet ENTER di terminal...")

    ###############################################
    ############# LANJUT DELETE ROW 3 #############
    ###############################################
    
    ws.delete_rows(3)
    wb.save(file_path)

    print("✅ Delete baris 3 di Excel: DONE.")

    ###############################################
    ########## LANJUT KE FORM BERIKUTNYA ##########
    ###############################################

    jawaban = input("➡️  Lanjut data berikutnya? (Y/N): ").strip().lower()
    if jawaban == 'y':
        print("▶️  Next data...")
    elif jawaban == 'n':
        print("⏹️ Proses dihentikan user.")
        exit()
    else:
        print("⚠️  Unrecognized. Please tap Y or N on your keyboard.")

    # >>>> Lanjut ke file Loop1 secara otomatis <<<<
    subprocess.run(["python", "ecrf1.py"])

except KeyboardInterrupt:
    print("\n⏹️ Proses dihentikan user.")

except Exception as e:
    print(f"\n❌ An unhandled exception occurred: {e}")
    exit() 



