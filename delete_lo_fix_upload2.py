import win32com.client
import openpyxl
import time
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

# --- FUNGSI UNTUK ANTARMUKA PENGGUNA (PESAN, INPUT & UPLOAD) ---
def jalankan_antarmuka_pengguna():
    """
    Menampilkan semua interaksi dengan pengguna di awal:
    1. Pesan peringatan.
    2. Form input untuk alasan penghapusan.
    3. Dialog untuk memilih file Excel.
    Mengembalikan (alasan_penghapusan, path_file).
    """
    # Membuat window utama tkinter dan langsung menyembunyikannya
    root = tk.Tk()
    root.withdraw()
    
    # 1. Menampilkan pesan peringatan pertama
    messagebox.showwarning(
        "Peringatan Penting",
        "Pastikan Anda sudah login ke SAP GUI dan T-Code VL02N sudah terbuka standby di layar utama SAP."
    )
    
    # 2. Loop untuk meminta input alasan penghapusan (tidak boleh kosong)
    alasan_text = ""
    while not alasan_text:
        alasan_text = simpledialog.askstring(
            "Alasan Penghapusan",
            "Masukkan alasan penghapusan Delivery Order:",
            parent=root
        )
        
        # Jika pengguna menekan 'Cancel', hentikan program
        if alasan_text is None:
            return None, None # Mengembalikan nilai kosong untuk di-handle di main()
            
        # Jika pengguna menekan 'OK' tapi tidak mengisi apa-apa
        if not alasan_text.strip():
            messagebox.showerror("Input Kosong", "Alasan penghapusan tidak boleh kosong. Silakan isi.")
            alasan_text = "" # Set ulang agar loop berlanjut

    # 3. Setelah user mengisi alasan, tampilkan dialog 'Open File'
    file_path = filedialog.askopenfilename(
        parent=root,
        title="Silakan Upload Format Excel Hapus LO",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    
    # Menutup window tkinter setelah selesai
    root.destroy()
    
    return alasan_text, file_path

# --- DEFINISI FUNGSI 1: UPDATE HEADER TEXT ---
def update_header_text(session, do_number, alasan_penghapusan):
    """Fungsi ini sekarang menerima 'alasan_penghapusan' dari pengguna."""
    try:
        print(f"\n✏️ Menambahkan Header Text untuk DO: {do_number}")
        # ... (kode navigasi SAP tetap sama) ...
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = str(do_number)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = len(str(do_number))
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
        time.sleep(1)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(0, 0)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem("Z004", "Column1")
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem("Z004", "Column1")
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem("Z004", "Column1")
        time.sleep(1)
        
        # PERUBAHAN UTAMA: Menggunakan teks dari input pengguna
        shell_target = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell")
        shell_target.text = alasan_penghapusan
        print(f"   -> Alasan ditambahkan: '{alasan_penghapusan}'")
        
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(1)
        print(f"✅ DO {do_number} berhasil diperbarui.")
    except Exception as e:
        print(f"❌ Gagal update DO {do_number}: {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except: pass
        time.sleep(1)

# --- DEFINISI FUNGSI 2: HAPUS LINE ITEM (Tidak ada perubahan di sini) ---
def hapus_do(session, do_number):
    try:
        print(f"\n🚚 Proses Hapus Item untuk DO: {do_number}")
        # ... (kode di dalam fungsi ini tidak berubah) ...
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = str(do_number)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = len(str(do_number))
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").getAbsoluteRow(0).selected = True
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0,0]").setFocus()
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0,0]").caretPosition = 0
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/subSUBSCREEN_ICONBAR:SAPMV50A:1708/btnBT_POLO_T").press()
        time.sleep(1)
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[1]").sendVKey(0)
        time.sleep(1)
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(1)
        print(f"✅ DO {do_number} berhasil dihapus.")
    except Exception as e:
        print(f"❌ Gagal proses DO {do_number}: {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except: pass
        time.sleep(1)

# --- BAGIAN EKSEKUSI UTAMA ---
def main():
    # Langkah 1: Jalankan antarmuka pengguna (pesan, input, & dialog file)
    alasan_hapus, excel_path = jalankan_antarmuka_pengguna()

    # Jika pengguna menekan 'Cancel' di dialog mana pun, hentikan program
    if not excel_path or alasan_hapus is None:
        print("❌ Proses dibatalkan oleh pengguna.")
        sys.exit()

    # Jendela terminal akan muncul sekarang, menampilkan proses di bawah ini
    print(f"📝 Alasan yang diberikan: '{alasan_hapus}'")
    print(f"📂 File yang akan diproses: {excel_path}\n")

    # Langkah 2: Hubungkan ke SAP
    try:
        print("🔌 Menghubungkan ke SAP GUI...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        print("✅ Berhasil terhubung ke SAP.")
    except Exception as e:
        print(f"❌ Gagal menghubungkan ke SAP: {e}")
        input("Pastikan SAP GUI sudah berjalan. Tekan Enter untuk keluar.")
        sys.exit()

    # Langkah 3: Baca data dari Excel yang sudah dipilih
    try:
        # ... (Logika membaca Excel tetap sama) ...
        print("--- Membaca daftar DO dari Excel ---")
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        list_do_numbers = [row[0] for row in sheet.iter_rows(min_row=2, values_only=True) if row and row[0]]
        print(f"Ditemukan {len(list_do_numbers)} DO untuk diproses.")
        if not list_do_numbers:
            print("Tidak ada data DO yang ditemukan di dalam file. Program berhenti.")
            input("Tekan Enter untuk keluar.")
            sys.exit()
    except Exception as e:
        print(f"❌ Gagal membaca file Excel: {e}")
        input("Tekan Enter untuk keluar.")
        sys.exit()

    # Langkah 4 & 5: Jalankan proses inti
    print("\n==============================================")
    print("     LANGKAH 1: MENAMBAHKAN HEADER TEXT")
    print("==============================================")
    for do_number in list_do_numbers:
        # Kirim alasan yang diinput pengguna ke dalam fungsi
        update_header_text(session, do_number, alasan_hapus)
        time.sleep(2)
    print("\n✅ SELESAI LANGKAH 1.")

    print("\n==============================================")
    print("     LANGKAH 2: MENGHAPUS LINE ITEM DO")
    print("==============================================")
    for do_number in list_do_numbers:
        hapus_do(session, do_number)
        time.sleep(2)
    print("\n✅ SELESAI LANGKAH 2.")
    
    print("\n\n🎯 SEMUA PROSES TELAH SELESAI. 🎯")
    input("Tekan Enter untuk menutup program ini.")

# Menjalankan fungsi utama
if __name__ == "__main__":
    main()