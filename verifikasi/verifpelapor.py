import subprocess
import sys
import time
from datetime import datetime, time as dtime
from openpyxl.utils import column_index_from_string
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from verif_utils import (
    create_driver, load_excel_data, get_mapping, EXCEL_PATH, tulis_log, buka_form2, normalize_attr, 
    cek_dan_klik_verified_smart, get_label_text, play_error
    , wait_excel_closed, load_excel_log
)

driver, wait_long, wait_short = create_driver()
wait_excel_closed()
wb, ws, nomor_inklusi = load_excel_data()
log_wb, log_ws = load_excel_log()
SECTION = "Informasi Pelapor"
mapping = get_mapping(SECTION)


print("▶️ Script berjalan... tekan CTRL+C untuk berhenti manual.")

print("⏳ Menunggu tabel '#tbl_responden' & field siap...")
try:
    wait_long.until(EC.visibility_of_element_located((By.ID, "tbl_responden")))
    wait_long.until(EC.presence_of_element_located((By.ID, "field_cari")))

    # 🧩 Tambahan penting: tunggu sampai spinner / loading hilang
    wait_long.until(EC.invisibility_of_element_located((
        By.CSS_SELECTOR,
        ".loading_spinner[style*='block'], .spinner[style*='block'], .loading[style*='block']"
    )))

    time.sleep(0.5)  # beri waktu agar event JS selesai
    print("✅ Halaman siap, lanjut cari nomor inklusi...")
except TimeoutException:
    play_error()
    print("⚠️ Timeout: tabel atau field_cari tidak muncul. Periksa halaman atau network.")
    exit()

# --- Cari nomor inklusi ---
field = driver.find_element(By.ID, "field_cari")
driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
time.sleep(0.3)
driver.execute_script("arguments[0].click();", field)  # ✅ pakai JS click biar aman
field.clear()
field.send_keys(nomor_inklusi)
field.send_keys(Keys.ENTER)
time.sleep(1)  # beri waktu untuk hasil tabel muncul
wait_long.until(EC.presence_of_element_located(
    (By.XPATH, f"//tr[td[normalize-space(text())='{nomor_inklusi}']]")
))


try:
    print(f"🔍 Memproses nomor inklusi: {nomor_inklusi} pada FORM {SECTION}")
    if not buka_form2(driver, wait_long, nomor_inklusi, nama_form=SECTION):
        play_error()
        print("⛔ Tidak bisa melanjutkan karena form gagal dibuka.")
        exit()

    verified_cache = {}
    for key in mapping:
        try:
            verified_cache[key] = driver.find_element(By.ID, f"verified_{key}")
        except:
            verified_cache[key] = None
    print("✅ Verified cache done")

    mismatches = []

    # --- Loop utama ---
    for key, m in mapping.items():
        col_index = column_index_from_string(m['col'])
        cell_value = ws.cell(row=3, column=col_index).value

        if cell_value is None:
            expected = ""
            field_type = "text"
        elif isinstance(cell_value, dtime):
            expected = cell_value.strftime("%H:%M")
            field_type = "time"
        elif isinstance(cell_value, datetime):
            expected = cell_value.strftime("%d-%m-%Y")
            field_type = "date"
        else:
            expected = str(cell_value).strip()
            field_type = "text"

        try:
            val_elem = driver.find_element(By.ID, f"val_verified_{key}")
            current_val = (val_elem.get_attribute("value") or "").strip()
        except:
            current_val = ""

        if expected == "":
            print(f"⏭️ Skip verify_{key} ({m.get('id') or m.get('radio')}) karena Excel kosong")
            continue

        # --- Ambil actual value ---
        actual = ""
        if "id" in m:
            try:
                el = driver.find_element(By.ID, m["id"])
                actual_raw = el.get_attribute("value-data") or el.get_attribute("value") or ""
                actual = normalize_attr(actual_raw)
            except:
                actual = ""
        elif "radio" in m:
            try:
                actual = driver.find_element(
                    By.CSS_SELECTOR, f"#{m['radio']} input:checked"
                ).get_attribute("value").strip()
            except:
                actual = ""

        # --- Ambil nama field dari cache ---
        field_name = get_label_text(driver, key, m)

        print(f"🔎 verify_{key} ({field_name}): expected={expected!r}, got={actual!r}")

        # --- Klik verified pakai cached element ---
        icon_elem = verified_cache.get(key)
        if icon_elem:
            cek_dan_klik_verified_smart(driver, wait_short, key, field_name, expected, actual, current_val, field_type=field_type)
        else:
            print(f"⚠️ verified_{key} tidak ditemukan ({field_name})")

        if actual != expected:
            mismatches.append(f"{field_name} (expected={expected}, got={actual})")


    if mismatches:
        tulis_log(log_ws, log_wb, nomor_inklusi, SECTION, mismatches)

    # --- Pause manual ---
    ans = input("👉 Cek data, jika sudah OK, tekan ENTER untuk lanjut, atau ketik 'stop' untuk keluar: ").strip().lower()
    if ans == "stop":
        print("⏹️ Proses dihentikan user.")
        exit()

    # --- tunggu tombol VERIFY muncul ---
    verify_btn = wait_short.until(EC.element_to_be_clickable((By.ID, "btn-verify")))
    # scroll ke tombol dan klik
    driver.execute_script("arguments[0].scrollIntoView(true);", verify_btn)
    time.sleep(0.2)
    driver.execute_script("arguments[0].click();", verify_btn)
    print("✅ Tombol VERIFY berhasil diklik")

    # --- tunggu tombol KEMBALI muncul ---
    # back_btn = wait_short.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn__cancel")))
    # driver.execute_script("arguments[0].scrollIntoView(true);", back_btn)
    # time.sleep(0.2)
    # driver.execute_script("arguments[0].click();", back_btn)
    # print("✅ Tombol KEMBALI berhasil diklik")

    # ---  LANJUT KE FORM BERIKUTNYA 
    next_file = "verifpasien.exe" if getattr(sys, 'frozen', False) else "verifpasien.py"
    if next_file.endswith(".exe"):
        subprocess.run([next_file])
    else:
        subprocess.run([sys.executable, next_file])


except KeyboardInterrupt:
    print("\n🛑 Script dihentikan manual oleh user (CTRL+C).")
except Exception as e:
    play_error()
    print(f"❌ Error di verifpelapor: {e}")

