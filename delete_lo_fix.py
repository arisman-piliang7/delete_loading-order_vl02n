import win32com.client
import openpyxl
import time
import sys

# --- KONFIGURASI ---
EXCEL_PATH = r"C:\SAP_Script\hapus LO.xlsx"

# --- INISIALISASI KONEKSI SAP & EXCEL ---
try:
    print("🔌 Menghubungkan ke SAP GUI...")
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not isinstance(SapGuiAuto, win32com.client.CDispatch):
        print("❌ SAP GUI tidak sedang berjalan. Harap buka SAP Logon terlebih dahulu.")
        sys.exit(1)
        
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    print("✅ Berhasil terhubung ke SAP.")
except Exception as e:
    print(f"❌ Gagal menghubungkan ke SAP: {e}")
    print("Pastikan Anda sudah login ke SAP.")
    sys.exit(1)

try:
    print(f"📂 Membuka file Excel: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    sheet = wb.active
    print("✅ File Excel berhasil dimuat.")
except FileNotFoundError:
    print(f"❌ File Excel tidak ditemukan di path: {EXCEL_PATH}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Gagal membuka file Excel: {e}")
    sys.exit(1)

# --- DEFINISI FUNGSI 1: UPDATE HEADER TEXT ---
def update_header_text(do_number):
    try:
        print(f"\n✏️ Menambahkan Header Text untuk DO: {do_number}")

        # Masuk VL02N
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Input nomor DO
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = str(do_number)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = len(str(do_number))
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Navigasi ke Header > Texts
        session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
        time.sleep(1)

        # Pilih line Z004 (Kode asli Anda dipertahankan)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(0, 0)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem("Z004", "Column1")
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem("Z004", "Column1")
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem("Z004", "Column1")
        time.sleep(1)

        # Tulis teks "do outstanding"
        shell_target = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV50A:2120/subTEXTEDIT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell")
        shell_target.text = "do outstanding"
        shell_target.setSelectionIndexes(14, 14)

        # Simpan
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(1)

        print(f"✅ DO {do_number} berhasil diperbarui.")
    except Exception as e:
        print(f"❌ Gagal update DO {do_number}: {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass
        time.sleep(1)

# --- DEFINISI FUNGSI 2: HAPUS LINE ITEM ---
def hapus_do(do_number):
    try:
        print(f"\n🚚 Proses Hapus Item untuk DO: {do_number}")

        # Masuk ke VL02N
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Input DO number
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = str(do_number)
        session.findById("wnd[0]/usr/ctxtLIKP-VBELN").caretPosition = len(str(do_number))
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Select item
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").getAbsoluteRow(0).selected = True
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0,0]").setFocus()
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER/txtLIPS-POSNR[0,0]").caretPosition = 0

        # Tekan tombol hapus
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/subSUBSCREEN_ICONBAR:SAPMV50A:1708/btnBT_POLO_T").press()
        time.sleep(1)

        # Konfirmasi hapus
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        session.findById("wnd[1]").sendVKey(0)
        time.sleep(1)

        # Simpan
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        time.sleep(1)

        print(f"✅ DO {do_number} berhasil dihapus.")
    except Exception as e:
        print(f"❌ Gagal proses DO {do_number}: {e}")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVL02N"
            session.findById("wnd[0]").sendVKey(0)
        except:
            pass
        time.sleep(1)

# --- BAGIAN EKSEKUSI UTAMA ---

# Langkah 0: Baca semua DO dari Excel ke dalam list
print("\n--- Membaca daftar DO dari Excel ---")
list_do_numbers = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    do_num = row[0]
    if do_num:
        list_do_numbers.append(do_num)
print(f"Ditemukan {len(list_do_numbers)} DO untuk diproses.")

# Langkah 1: Jalankan proses update header text untuk semua DO
print("\n==============================================")
print("     LANGKAH 1: MENAMBAHKAN HEADER TEXT")
print("==============================================")
if not list_do_numbers:
    print("Tidak ada DO untuk diproses.")
else:
    for do_number in list_do_numbers:
        update_header_text(do_number)
        time.sleep(2)  # Jeda antar DO

print("\n✅ SELESAI LANGKAH 1: Semua header DO berhasil diproses.")

# Langkah 2: Jalankan proses hapus item untuk semua DO
print("\n==============================================")
print("     LANGKAH 2: MENGHAPUS LINE ITEM DO")
print("==============================================")
if not list_do_numbers:
    print("Tidak ada DO untuk diproses.")
else:
    for do_number in list_do_numbers:
        hapus_do(do_number)
        time.sleep(2) # Jeda antar DO

print("\n✅ SELESAI LANGKAH 2: Semua line item DO berhasil diproses.")
print("\n\n🎯 SEMUA PROSES TELAH SELESAI. 🎯")