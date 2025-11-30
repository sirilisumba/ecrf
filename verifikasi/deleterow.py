import subprocess
from openpyxl import load_workbook

try:
    
    # --- file deleterow.py
    # Buka file Excel
    file_path = "data.xlsx"
    wb = load_workbook(file_path)
    ws = wb.active

    # Hapus baris ke-3
    ws.delete_rows(3)

    # Simpan kembali file
    wb.save(file_path)

    print("✅ Deleting Row 3 in data.xlsx: SUCCESSFUL.")
    print("📌 Process: END.")

    print("\n Program END successfully.")

    subprocess.run(["python", "verifall.py"])

except KeyboardInterrupt:
    print("\n⏹️ Kill process by user.")

except Exception as e:
    print(f"\n❌ An unhandled exception occurred: {e}")


