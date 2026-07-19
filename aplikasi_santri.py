import subprocess
import sys

# Trik otomatis install jika library belum ada
try:
    import xlsxwriter
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter"])

import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
import time
import io  # Ditambahkan untuk handle eksport ke Excel asli

# =========================================================
# 🔐 KODE BARU 1: SISTEM KEAMANAN 3 AKUN UTAMA
# =========================================================
AKUN_DATABASE = {
    "ADMIN PONTRA": {"password": "PONTRA.1", "akses": "Putra dan Ustadz"},
    "ADMIN PONTRI": {"password": "PONTRI.1", "akses": "Putri dan Ustadzah"},
    "ADMIN PUSAT": {"password": "PUSAT.1", "akses": "Semua"}
}

if "terautentikasi" not in st.session_state:
    st.session_state["terautentikasi"] = False
    st.session_state["hak_akses"] = None

def halaman_login():
    st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🔒 Sistem Data Pusat Pesantren</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #6B7280;'>Silakan login sesuai dengan akun komplek Anda</p>", unsafe_allow_html=True)
    
    with st.form("form_login"):
        username = st.text_input("Username Admin:", placeholder="Masukkan admin_putra / admin_putri / admin_pusat")
        password = st.text_input("Password / Kata Sandi:", type="password", placeholder="Masukkan password")
        tombol_masuk = st.form_submit_button("Masuk ke Aplikasi 🚀", use_container_width=True)
        
        if tombol_masuk:
            if username in AKUN_DATABASE and AKUN_DATABASE[username]["password"] == password:
                st.session_state["terautentikasi"] = True
                st.session_state["hak_akses"] = AKUN_DATABASE[username]["akses"]
                st.success(f"✅ Login Berhasil!")
                st.rerun()
            else:
                st.error("❌ Username atau Password salah!")

if not st.session_state["terautentikasi"]:
    halaman_login()
    st.stop() # <--- Di sini Python mogok kalau belum login!
# =========================================================
# 🚪 TOMBOL LOG OUT (DI SIDEBAR SEBELAH KIRI)
# =========================================================
with st.sidebar:
    st.markdown(f"### 👤 Admin: **{st.session_state['hak_akses']}**")
    st.write("---")
    
    # Tombol merah untuk Log Out
    tombol_logout = st.button("🚪 Keluar / Log Out", use_container_width=True, type="primary")
    
    if tombol_logout:
        # Reset ulang semua status memori login
        st.session_state["terautentikasi"] = False
        st.session_state["hak_akses"] = None
        st.success("Berhasil keluar! Mengunci sistem...")
        time.sleep(1) # Beri jeda 1 detik biar estetik
        st.rerun() # Refresh halaman untuk memunculkan gembok kembali

    st.write("---")

    # ==================================================================
    # REFRESH GLOBAL (UNTUK MULTI-USER/KOMPUTER 1, 2, 3) - VERSI AMAN TOTAL
    # ==================================================================
    st.subheader("🔄 Sinkronisasi Sistem")
    
    if st.button("🔄 Refresh Semua Data", use_container_width=True):
        # Beri notifikasi melayang di layar
        st.toast("⚡ Menyinkronkan data terbaru dari komputer lain...")
        time.sleep(0.5)
        # Paksa Streamlit mengulang skrip dari atas dan membaca ulang semua CSV
        st.rerun()
# ==========================================
# 1. KONFIGURASI UTAMA & DEKLARASI FILE
# ==========================================
st.set_page_config(page_title="Sistem Informasi Pesantren", layout="wide")

FILE_SANTRI = "data_santri.csv"
FILE_ASATIDZ = "data_asatidz.csv"
FILE_PENGURUS_KAMAR = "data_pengurus_kamar.csv"
FOLDER_BACKUP = "backup_data_santri"

NAMA_PONPES = "API RIYADLOTUSSHOLIKHIN WASSHOLIKHAT"
DAFTAR_KAMAR = ["Kamar A1", "Kamar A2", "Kamar B1", "Kamar B2", "Kamar C2","Kamar Putri"]

KOLOM_SANTRI = [
    "NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "AYAH", "IBU", "WALI", 
    "DUKUH", "DESA", "KECAMATAN", "KABUPATEN", 
    "STATUS", "NIK", "KK", "TEMPAT LAHIR", "TGL LAHIR", "KELAS", "KAMAR"
]

KOLOM_ASATIDZ = [
    "NIU", "NAMA USTADZ/AH", "JENIS KELAMIN", "TUGAS/JABATAN",
    "DUKUH", "DESA", "KECAMATAN", "KABUPATEN", 
    "STATUS", "NIK", "KK", "TEMPAT LAHIR", "TGL LAHIR", "NO HP"
]

KOLOM_CETAK_DAERAH = ["NAMA SANTRI", "JENIS KELAMIN", "AYAH", "DUKUH", "DESA", "KECAMATAN", "KABUPATEN", "KELAS", "KAMAR"]

BULAN_INDO = {
    "januari": "01", "februari": "02", "maret": "03", "april": "04", "mei": "05", "juni": "06",
    "juli": "07", "agustus": "08", "september": "09", "oktober": "10", "november": "11", "desember": "12",
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "mei": "05", "jun": "06",
    "jul": "07", "agu": "08", "sep": "09", "okt": "10", "nov": "11", "des": "12"
}

MIN_DATE = datetime(1945, 1, 1)
MAX_DATE = datetime(2050, 12, 31)

# ==========================================
# 2. SISTEM KEAMANAN UNDO / REDO & BACKUP
# ==========================================
if not os.path.exists(FOLDER_BACKUP):
    os.makedirs(FOLDER_BACKUP)

def jalankan_auto_backup(filename, label_aksi="backup"):
    if os.path.exists(filename):
        try:
            stempel_waktu = datetime.now().strftime("%Y%m%d_%H%M%S")
            nama_file_asli, ekstensi = os.path.splitext(filename)
            path_tujuan_backup = os.path.join(FOLDER_BACKUP, f"{nama_file_asli}_{label_aksi}_{stempel_waktu}{ekstensi}")
            shutil.copy2(filename, path_tujuan_backup)
        except: pass

def eksekusi_undo(filename):
    nama_murni, ekstensi = os.path.splitext(filename)
    files = sorted([f for f in os.listdir(FOLDER_BACKUP) if f.startswith(f"{nama_murni}_sebelum_")], 
                   key=lambda x: os.path.getmtime(os.path.join(FOLDER_BACKUP, x)), reverse=True)
    if files:
        jalankan_auto_backup(filename, "sesudah")
        shutil.copy2(os.path.join(FOLDER_BACKUP, files[0]), filename)
        os.remove(os.path.join(FOLDER_BACKUP, files[0]))
        return True
    return False

def eksekusi_redo(filename):
    nama_murni, ekstensi = os.path.splitext(filename)
    files = sorted([f for f in os.listdir(FOLDER_BACKUP) if f.startswith(f"{nama_murni}_sesudah_")], 
                   key=lambda x: os.path.getmtime(os.path.join(FOLDER_BACKUP, x)), reverse=True)
    if files:
        jalankan_auto_backup(filename, "sebelum")
        shutil.copy2(os.path.join(FOLDER_BACKUP, files[0]), filename)
        os.remove(os.path.join(FOLDER_BACKUP, files[0]))
        return True
    return False

# ==========================================
# 3. ENGINE LOADER & SAVER (ANTI DESIMAL & ANTI CRASH)
# ==========================================
def bersihkan_tanggal_indo(tgl_str):
    try:
        tgl_str = str(tgl_str).strip().lower()
        if tgl_str == "nan" or tgl_str == "" or "belum lengkap" in tgl_str:
            return "🔴 BELUM LENGKAP"
        try:
            datetime.strptime(tgl_str, "%Y-%m-%d")
            return tgl_str
        except: pass
        parts = tgl_str.split()
        if len(parts) == 3:
            hari = parts[0].zfill(2)
            bulan_text = parts[1]
            tahun = parts[2]
            if bulan_text in BULAN_INDO:
                bulan = BULAN_INDO[bulan_text]
                return f"{tahun}-{bulan}-{hari}"
        return "🔴 BELUM LENGKAP"
    except: return "🔴 BELUM LENGKAP"

def load_data_santri():
    df_kosong = pd.DataFrame(columns=KOLOM_SANTRI)
    if os.path.exists(FILE_SANTRI):
        try:
            # Membaca sebagai string sejak awal agar terhindar dari desimal otomatis
            df = pd.read_csv(FILE_SANTRI, sep=',', encoding='utf-8-sig', dtype=str)
            df.columns = df.columns.str.upper().str.strip()
            if "NAMA SANTRI" not in df.columns or len(df.columns) <= 1:
                df = pd.read_csv(FILE_SANTRI, sep=';', encoding='utf-8-sig', dtype=str)
                df.columns = df.columns.str.upper().str.strip()
            
            for col in KOLOM_SANTRI:
                if col not in df.columns:
                    if col == "KAMAR": df[col] = "Belum Diatur"
                    elif col == "STATUS": df[col] = "Aktif"
                    elif col == "JENIS KELAMIN": df[col] = "Putra"
                    else: df[col] = "🔴 BELUM LENGKAP"
                else:
                    df[col] = df[col].fillna("🔴 BELUM LENGKAP").astype(str).str.strip()

            # =====================================================================
            # 🗂️ KODE SARINGAN DATA SANTRI (FIX LOGIN SAKTI)
            # =====================================================================
            if not df.empty and "JENIS KELAMIN" in df.columns:
                akses_login = str(st.session_state.get("hak_akses", ""))
                # Pakai "in" karena statusnya mengandung kata "Putra dan Ustadz"
                if "Putra" in akses_login:
                    df = df[df["JENIS KELAMIN"].astype(str).str.upper().isin(["PUTRA"])]
                elif "Putri" in akses_login:
                    df = df[df["JENIS KELAMIN"].astype(str).str.upper().isin(["PUTRI"])]

            return df
            
        except Exception as e:
            st.error(f"Gagal membaca file santri: {e}")
            return df_kosong
    return df_kosong

def load_data_asatidz():
    df_kosong = pd.DataFrame(columns=KOLOM_ASATIDZ)
    if os.path.exists(FILE_ASATIDZ):
        try:
            df = pd.read_csv(FILE_ASATIDZ, sep=',', encoding='utf-8-sig', dtype=str)
            df.columns = df.columns.str.upper().str.strip()
            if "NAMA USTADZ/AH" not in df.columns and "NAMA LENGKAP" in df.columns:
                df = df.rename(columns={"NAMA LENGKAP": "NAMA USTADZ/AH"})
                
            if "NAMA USTADZ/AH" not in df.columns or len(df.columns) <= 1:
                df = pd.read_csv(FILE_ASATIDZ, sep=';', encoding='utf-8-sig', dtype=str)
                df.columns = df.columns.str.upper().str.strip()
                if "NAMA USTADZ/AH" not in df.columns and "NAMA LENGKAP" in df.columns:
                    df = df.rename(columns={"NAMA LENGKAP": "NAMA USTADZ/AH"})
            
            for col in KOLOM_ASATIDZ:
                if col not in df.columns:
                    if col == "NIU": df[col] = "TEMP_GURU"
                    elif col == "STATUS": df[col] = "Aktif"
                    elif col == "JENIS KELAMIN": df[col] = "USTADZ"
                    else: df[col] = "🔴 BELUM LENGKAP"
                else:
                    df[col] = df[col].fillna("🔴 BELUM LENGKAP").astype(str).str.strip()
            
            if "STATUS" in df.columns:
                df["STATUS"] = df["STATUS"].str.strip().str.upper().replace({"AKTIF": "Aktif", "KELUAR": "Keluar", "LULUS": "Lulus"})
                df.loc[~df["STATUS"].isin(["Aktif", "Keluar", "Lulus"]), "STATUS"] = "Aktif"

            if "JENIS KELAMIN" in df.columns:
                df["JENIS KELAMIN"] = df["JENIS KELAMIN"].str.strip().str.upper()
                df.loc[df["JENIS KELAMIN"].isin(["Putra", "Laki-laki", "USTADZ"]), "JENIS KELAMIN"] = "USTADZ"
                df.loc[df["JENIS KELAMIN"].isin(["Putri", "Perempuan", "USTADZAH"]), "JENIS KELAMIN"] = "USTADZAH"
                
            for c_num in ["NIU", "NIK", "KK"]:
                if c_num in df.columns:
                    df[c_num] = df[c_num].astype(str).str.replace(".0", "", regex=False).str.replace("nan", "🔴 BELUM LENGKAP", regex=False).str.strip()
                    df[c_num] = df[c_num].apply(lambda x: x.split('.')[0] if '.' in x and x.split('.')[1] == '0' else x)

            # =====================================================================
            # 🗂️ KODE SARINGAN DATA ASATIDZ (FIX LOGIN SAKTI)
            # =====================================================================
            if not df.empty and "JENIS KELAMIN" in df.columns:
                akses_login = str(st.session_state.get("hak_akses", ""))
                # Pakai "in" karena statusnya mengandung kata "Putra dan Ustadz"
                if "Putra" in akses_login:
                    df = df[df["JENIS KELAMIN"].isin(["Ustadz", "Putra"])]
                elif "Putri" in akses_login:
                    df = df[df["JENIS KELAMIN"].isin(["Ustadzah", "Putri"])]

            return df[KOLOM_ASATIDZ]
        except: 
            return df_kosong
    return df_kosong

def load_pengurus_kamar():
    if os.path.exists(FILE_PENGURUS_KAMAR):
        try: return pd.read_csv(FILE_PENGURUS_KAMAR, index_col="KAMAR", sep=',')
        except: pass
    df_pk = pd.DataFrame(index=DAFTAR_KAMAR, columns=["KETUA", "WAKIL"]).fillna("-")
    df_pk.index.name = "KAMAR"
    return df_pk

def save_all(df, filename):
    jalankan_auto_backup(filename, "sebelum")
    df.to_csv(filename, index=True if "data_pengurus_kamar" in filename else False, sep=',', encoding='utf-8-sig')

def tampilkan_notifikasi_sukses():
    st.toast("PERUBAHAN BERHASIL!", icon="✅")
    st.success("✅ PERUBAHAN BERHASIL!")
    time.sleep(1)

# Fungsi Generator konversi ke Excel Asli (.xlsx) tanpa merusak teks KK/NIK
def konversi_ke_excel_asli(df_input):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_input.to_excel(writer, index=False, sheet_name='Data_Pesantren')
    return output.getvalue()

# ==========================================
# 📋 AMBIL MEMORI DATA AKTIF & FILTER AKSES
# ==========================================
df_santri = load_data_santri()
df_asatidz = load_data_asatidz()
df_pengurus_kamar = load_pengurus_kamar()

# 1. DETEKSI STATUS AKSES GLOBAL (MEMBONGKAR MEMORI LOGIN APLIKASI)
status_login_global = str(st.session_state.items()).upper()

# 2. PROSES FILTER AMAN (TERMASUK UNTUK ADMIN PUSAT)
if "PUSAT" in status_login_global or "SUPER" in status_login_global:
    # JIKA ADMIN PUSAT/SUPERADMIN: JANGAN DI-FILTER, BIARKAN SEMUA DATA MUNCUL
    pass
elif "PUTRA" in status_login_global:
    # JIKA ADMIN PUTRA: HANYA TAMPILKAN USTADZ
    if not df_asatidz.empty and "JENIS KELAMIN" in df_asatidz.columns:
        df_asatidz = df_asatidz[df_asatidz["JENIS KELAMIN"] == "Ustadz"]
elif "PUTRI" in status_login_global:
    # JIKA ADMIN PUTRI: HANYA TAMPILKAN USTADZAH
    if not df_asatidz.empty and "JENIS KELAMIN" in df_asatidz.columns:
        df_asatidz = df_asatidz[df_asatidz["JENIS KELAMIN"] == "Ustadzah"]

if not df_santri.empty:
    list_kelas_terdeteksi = sorted([x for x in df_santri["KELAS"].unique() if x and x not in ["nan", "🔴 BELUM LENGKAP", "[Belum Lengkap]"]])
else: list_kelas_terdeteksi = []
if not list_kelas_terdeteksi:
    list_kelas_terdeteksi = ["Kelas As Shifir", "Kelas Ibtida'iyah", "Kelas Al Jurumiyah", "Kelas As Shorfu", "Kelas Al Fiyah", "Kelas Fathul Wahab", "Kelas Al Makhali"]

# LOGIKA KAKAK BERADIK GLOBAL PESANTREN
kk_santri = []
if not df_santri.empty:
    df_s_valid = df_santri[(df_santri["KK"] != "🔴 BELUM LENGKAP") & (df_santri["STATUS"] == "Aktif")]
    kk_santri = df_s_valid["KK"].tolist()

kk_asatidz = []
if not df_asatidz.empty:
    df_a_valid = df_asatidz[(df_asatidz["KK"] != "🔴 BELUM LENGKAP") & (df_asatidz["STATUS"] == "Aktif")]
    kk_asatidz = df_a_valid["KK"].tolist()

semua_kk_global = kk_santri + kk_asatidz

kk_counts_global = {}
for kk in semua_kk_global:
    kk_counts_global[kk] = kk_counts_global.get(kk, 0) + 1

def beri_tanda_saudara_global(row):
    kk = row["KK"]
    status = row["STATUS"]
    if status == "Aktif" and kk in kk_counts_global and kk_counts_global[kk] > 1:
        return f"👥 Beradik-{kk_counts_global[kk]} (Global)"
    return "Mandiri"

if not df_santri.empty:
    df_santri["HUBUNGAN"] = df_santri.apply(beri_tanda_saudara_global, axis=1)
else:
    df_santri["HUBUNGAN"] = "Mandiri"

if not df_asatidz.empty:
    df_asatidz["HUBUNGAN"] = df_asatidz.apply(beri_tanda_saudara_global, axis=1)
else:
    df_asatidz["HUBUNGAN"] = "Mandiri"

# ==========================================
# 4. INTERFACE STRUKTUR TAB UTAMA
# ==========================================
st.title("🕌 Sistem Informasi & Administrasi Pesantren Terpadu")
st.markdown(f"#### 👋 Selamat Datang, Jangan Lupa Bismillah — Ponpes {NAMA_PONPES}")
st.write("---")

tab_dash, tab_input, tab_kelola, tab_daerah, tab_kamar, tab_asatidz_panel = st.tabs([
    "📊 Dashboard Utama", "📥 Tambah & Impor Excel", "📋 Kelola Data Santri", "📍 Pemetaan Daerah & Kelas", "🛏️ Manajemen Kamar", "👳 Data Asatidz"
])

# ------------------------------------------
# TAB 1: DASHBOARD UTAMA
# ------------------------------------------
with tab_dash:
    st.header("📊 Ringkasan Data & Intisari Pondok")
    df_real = df_santri[~df_santri["NAMA SANTRI"].isin(["nan", "🔴 BELUM LENGKAP", ""])] if not df_santri.empty else pd.DataFrame()
    df_real_as = df_asatidz[~df_asatidz["NAMA USTADZ/AH"].isin(["nan", "🔴 BELUM LENGKAP", ""])] if not df_asatidz.empty else pd.DataFrame()
    
    total_s = len(df_real[df_real["STATUS"] == "Aktif"]) if not df_real.empty else 0
    total_g = len(df_real_as[df_real_as["STATUS"] == "Aktif"]) if not df_real_as.empty else 0
    jumlah_global = total_s + total_g
    
    st.markdown(f"""
    <div style="background-color:#1E6B7B; padding:15px; border-radius:10px; text-align:center; margin-bottom:20px;">
        <h2 style="color:white; margin:0;">🌎 TOTAL JUMLAH GLOBAL (SELURUH JIWA)</h2>
        <h1 style="color:yellow; margin:5px 0 0 0; font-size:45px;">{jumlah_global} Orang</h1>
        <p style="color:white; margin:0; font-size:14px;">Gabungan antara seluruh Santri Aktif & Ustadz/Ustadzah Pengajar</p>
    </div>
    """, unsafe_allow_html=True)
    
    c_tot1, c_tot2, c_tot3, c_tot4 = st.columns(4)
    c_tot1.metric("Total Santri Aktif", f"{total_s} Orang")
    c_tot2.metric("Total Guru/Asatidz Aktif", f"{total_g} Orang")
    
    beradik_2 = sum(1 for jml in kk_counts_global.values() if jml == 2)
    beradik_3 = sum(1 for jml in kk_counts_global.values() if jml == 3)
    c_tot3.metric("Keluarga (Beradik 2)", f"{beradik_2} Kelompok")
    c_tot4.metric("Keluarga (Beradik 3)", f"{beradik_3} Kelompok")
    
    st.write("---")
    st.subheader("📋 Rincian Demografis")
    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    
    # Hitung data menggunakan format HURUF BESAR agar sinkron dengan saringan data
    s_putra = len(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["JENIS KELAMIN"] == "PUTRA")]) if not df_real.empty else 0
    s_putri = len(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["JENIS KELAMIN"] == "PUTRI")]) if not df_real.empty else 0
    
    u_ustadz = len(df_real_as[(df_real_as["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real_as["JENIS KELAMIN"] == "USTADZ")]) if not df_real_as.empty else 0
    u_ustadzah = len(df_real_as[(df_real_as["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real_as["JENIS KELAMIN"] == "USTADZAH")]) if not df_real_as.empty else 0
        
    col_p1.info(f"👦 **SANTRI PUTRA:** {s_putra} Orang")
    col_p2.info(f"👧 **SANRTI PUTRI:** {s_putri} Orang")
    col_p3.success(f"👳 **Ustadz:** {u_ustadz} Orang")
    col_p4.success(f"🧕 **Ustadzah:** {u_ustadzah} Orang")
    
    st.write("---")
    st.subheader("🔍 Cek Data Cepat dari Dashboard")
    pilihan_lihat = st.selectbox("Pilih Kategori Data untuk langsung ditampilkan:", 
                                 ["-- Pilih Kategori --", "Santri Putra", "Santri Putri", "Data Ustadz", "Data Ustadzah", "Hubungan Kakak Beradik (Santri) 👥", "Hubungan Keluarga (Asatidz) 👥"])
    
    if pilihan_lihat != "-- Pilih Kategori --":
        st.write(f"### 📋 Menampilkan Data: {pilihan_lihat}")
        if pilihan_lihat == "Santri Putra" and not df_real.empty:
            st.dataframe(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["JENIS KELAMIN"] == "PUTRA")][["NO INDUK", "NAMA SANTRI", "KELAS", "KAMAR", "DESA"]], use_container_width=True)
        elif pilihan_lihat == "Santri Putri" and not df_real.empty:
            st.dataframe(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["JENIS KELAMIN"] == "PUTRI")][["NO INDUK", "NAMA SANTRI", "KELAS", "KAMAR", "DESA"]], use_container_width=True)
        elif pilihan_lihat == "Data Ustadz" and not df_real_as.empty:
            st.dataframe(df_real_as[(df_real_as["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real_as["JENIS KELAMIN"] == "USTADZ")], use_container_width=True)
        elif pilihan_lihat == "Data Ustadzah" and not df_real_as.empty:
            st.dataframe(df_real_as[(df_real_as["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real_as["JENIS KELAMIN"] == "USTADZAH")], use_container_width=True)
        elif pilihan_lihat == "Kakak Beradik 2 👥" and not df_real.empty:
            st.dataframe(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["HUBUNGAN"] == "👥 Beradik-2")][["NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "KK", "AYAH"]], use_container_width=True)
        elif pilihan_lihat == "Kakak Beradik 3 👥" and not df_real.empty:
            st.dataframe(df_real[(df_real["STATUS"].str.upper().str.strip() == "AKTIF") & (df_real["HUBUNGAN"] == "👥 Beradik-3")][["NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "KK", "AYAH"]], use_container_width=True)

    st.write("---")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("🏫 Distribusi Kelas")
        if not df_real.empty:
            df_ak = df_real[df_real["STATUS"] == "Aktif"]
            if not df_ak.empty: st.bar_chart(df_ak["KELAS"].value_counts())
    with col_g2:
        st.subheader("🛏️ Distribusi Kamar")
        if not df_real.empty:
            df_ak = df_real[df_real["STATUS"] == "Aktif"]
            if not df_ak.empty: st.bar_chart(df_ak["KAMAR"].value_counts())

# ------------------------------------------
# TAB 2: TAMBAH & IMPOR EXCEL SANTRI
# ------------------------------------------
with tab_input:
    st.header("📥 Input Data Santri")
    metode = st.radio("Pilih Metode:", ["Unggah / Upload Massal dari Excel (CSV)", "Ketik Manual Satu per Satu"])
    
    if metode == "Ketik Manual Satu per Satu":
        with st.form("form_santri", clear_on_submit=True):
            st.subheader("📝 Data Pribadi")
            c1, c2, c3 = st.columns(3)
            with c1:
                no_induk = st.text_input("NO INDUK")
                nama_santri = st.text_input("NAMA SANTRI*")
                jk = st.selectbox("JENIS KELAMIN", ["PUTRA", "PUTRI"])
                nik = st.text_input("NIK")
                no_kk = st.text_input("KK")
            with c2:
                tmp_lahir = st.text_input("TEMPAT LAHIR")
                tgl_lahir = st.date_input("TGL LAHIR", value=datetime(2010, 1, 1), min_value=MIN_DATE, max_value=MAX_DATE)
                kelas = st.selectbox("KELAS", list_kelas_terdeteksi)
                kamar = st.selectbox("KAMAR", DAFTAR_KAMAR)
            with c3:
                ayah = st.text_input("Nama AYAH")
                ibu = st.text_input("Nama IBU")
                wali_nama = st.text_input("Nama WALI")
                status = st.selectbox("STATUS", ["Aktif", "Keluar", "Lulus"])
            st.subheader("🏡 Data Alamat")
            cb1, cb2 = st.columns(2)
            with cb1:
                dukuh = st.text_input("DUKUH")
                desa = st.text_input("DESA")
            with cb2:
                komm = st.text_input("KECAMATAN")
                kabb = st.text_input("KABUPATEN")
                
            if st.form_submit_button("Simpan Data Santri"):
                if nama_santri:
                    final_no_induk = str(no_induk).strip() if no_induk else f"TEMP_{datetime.now().strftime('%M%S')}"
                    new_data = pd.DataFrame([[
                        final_no_induk, nama_santri, jk, ayah or "🔴 BELUM LENGKAP", ibu or "🔴 BELUM LENGKAP", wali_nama or "🔴 BELUM LENGKAP",
                        dukuh or "🔴 BELUM LENGKAP", desa or "🔴 BELUM LENGKAP", komm or "🔴 BELUM LENGKAP", kabb or "🔴 BELUM LENGKAP", 
                        status, str(nik) or "🔴 BELUM LENGKAP", str(no_kk) or "🔴 BELUM LENGKAP", tmp_lahir or "🔴 BELUM LENGKAP", 
                        tgl_lahir.strftime("%Y-%m-%d"), kelas, kamar
                    ]], columns=KOLOM_SANTRI)
                    df_santri = pd.concat([df_santri, new_data], ignore_index=True)
                    save_all(df_santri[KOLOM_SANTRI], FILE_SANTRI)
                    tampilkan_notifikasi_sukses()
                    st.rerun()
    else:
        st.subheader("📥 Impor via File Excel (.xlsx)")
        # REVISI: Mengubah tipe file yang diterima menjadi .xlsx
        uploaded_file = st.file_uploader("Pilih file Excel (.xlsx) Anda:", type=["xlsx"])
        if uploaded_file is not None:
            try:
                # Membaca file Excel asli dan memaksa semua kolom dibaca sebagai text/string
                df_upload = pd.read_excel(uploaded_file, dtype=str)
                df_upload.columns = df_upload.columns.str.upper().str.strip()
                
                if "NAMA SANTRI" in df_upload.columns:
                    df_upload = df_upload.fillna("🔴 BELUM LENGKAP")
                    if "NO INDUK" not in df_upload.columns: df_upload["NO INDUK"] = "🔴 BELUM LENGKAP"
                        
                    for col in KOLOM_SANTRI:
                        if col not in df_upload.columns:
                            if col == "KAMAR": df_upload[col] = "Belum Diatur"
                            elif col == "STATUS": df_upload[col] = "AKTIF"
                            elif col == "JENIS KELAMIN": df_upload[col] = "PUTRA"
                            else: df_upload[col] = "🔴 BELUM LENGKAP"
                        else:
                            df_upload[col] = df_upload[col].astype(str).str.replace("nan", "🔴 BELUM LENGKAP", regex=False).str.strip()
                    
                    if "STATUS" in df_upload.columns:
                        df_upload["STATUS"] = df_upload["STATUS"].str.strip().str.upper().replace({"AKTIF": "Aktif", "KELUAR": "Keluar", "LULUS": "Lulus"})
                        df_upload.loc[~df_upload["STATUS"].isin(["Aktif", "Keluar", "Lulus"]), "STATUS"] = "Aktif"
                    
                    if "JENIS KELAMIN" in df_upload.columns:
                        df_upload["JENIS KELAMIN"] = df_upload["JENIS KELAMIN"].str.strip().str.upper()

                    if "TGL LAHIR" in df_upload.columns:
                        df_upload["TGL LAHIR"] = df_upload["TGL LAHIR"].apply(bersihkan_tanggal_indo)
                    
                    # Bersihkan sisa desimal .0 pada KK/NIK jika admin telanjur salah ketik di Excel
                    for c_num in ["NO INDUK", "NIK", "KK"]:
                        if c_num in df_upload.columns:
                            df_upload[c_num] = df_upload[c_num].apply(lambda x: x.split('.')[0] if '.' in x and x.split('.')[1] == '0' else x)
                            
                    df_upload = df_upload[KOLOM_SANTRI]
                    st.write("👀 **Pratinjau Data Yang Akan Dimasukkan:**")
                    st.dataframe(df_upload.head(10))
                    
                    if st.button("🚀 Masukkan Semua Data ke Aplikasi", type="primary"):
                        df_santri = df_upload if (df_santri.empty or "🔴 BELUM LENGKAP" in df_santri["NAMA SANTRI"].values) else pd.concat([df_santri, df_upload], ignore_index=True)
                        save_all(df_santri[KOLOM_SANTRI], FILE_SANTRI)
                        tampilkan_notifikasi_sukses()
                        st.rerun()
                else:
                    st.error("❌ Gagal! File Excel harus memiliki kolom bernama 'NAMA SANTRI'")
            except Exception as e: 
                st.error(f"Gagal membaca file Excel: {e}")

# ------------------------------------------
# TAB 3: KELOLA DATA SANTRI (WITH UNDO & REDO)
# ------------------------------------------
with tab_kelola:
    st.header("📋 Daftar & Kelola Data Santri")
    
    uc1, uc2, uc3 = st.columns([1.5, 1.5, 7])
    with uc1:
        if st.button("↩️ Undo Aksi Santri", use_container_width=True):
            if eksekusi_undo(FILE_SANTRI):
                st.toast("Undo Berhasil!", icon="🔄")
                time.sleep(0.5)
                st.rerun()
            else: st.info("Tidak ada riwayat Undo")
    with uc2:
        if st.button("↪️ Redo Aksi Santri", use_container_width=True):
            if eksekusi_redo(FILE_SANTRI):
                st.toast("Redo Berhasil!", icon="🔄")
                time.sleep(0.5)
                st.rerun()
            else: st.info("Tidak ada riwayat Redo")
            
    st.write("---")
    df_santri_clean = df_santri[df_santri["NAMA SANTRI"].notna() & (df_santri["NAMA SANTRI"] != "") & (df_santri["NAMA SANTRI"] != "🔴 BELUM LENGKAP")]
    
    if not df_santri_clean.empty:
        f1, f2, f3 = st.columns(3)
        with f1: filter_status = st.multiselect("Filter STATUS:", ["Aktif", "Keluar", "Lulus"], default=["Aktif"])
        with f2: status_berkas = st.radio("Filter Kelengkapan:", ["Semua", "Hanya Belum Lengkap"], horizontal=True)
        with f3: filter_saudara = st.radio("Filter Saudara:", ["Semua", "Hanya Kakak Beradik 👥"], horizontal=True)
        
        search_query = st.text_input("🔍 Cari nama / No Induk / No KK:")
        df_filtered = df_santri_clean[df_santri_clean["STATUS"].isin(filter_status)]
        
        if status_berkas == "Hanya Belum Lengkap":
            df_filtered = df_filtered[df_filtered.apply(lambda r: "🔴 BELUM LENGKAP" in r.values, axis=1)]
        if filter_saudara == "Hanya Kakak Beradik 👥":
            df_filtered = df_filtered[df_filtered["HUBUNGAN"].str.contains("Beradik", na=False)]
        if search_query:
            df_filtered = df_filtered[
                df_filtered["NAMA SANTRI"].str.contains(search_query, case=False, na=False) | 
                df_filtered["NO INDUK"].str.contains(search_query, case=False, na=False) |
                df_filtered["KK"].str.contains(search_query, case=False, na=False)
            ]
            
        def beri_warna_tabel(val):
            if val == "🔴 BELUM LENGKAP": return "color: #FF4B4B; font-weight: bold; background-color: #FFEBEB;"
            if "👥" in str(val): return "color: #1E6B7B; font-weight: bold; background-color: #E8F8F5;"
            return ""

        kolom_tampil = ["NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "KELAS", "KAMAR", "KK", "HUBUNGAN", "STATUS"]
        st.dataframe(df_filtered[kolom_tampil].style.map(beri_warna_tabel), use_container_width=True)
        
        st.write("---")
        st.subheader("🛠️ Ubah / Hapus Data Santri")
        santri_terpilih = st.selectbox("Pilih Nama Santri:", ["-- Pilih Santri --"] + df_santri_clean["NAMA SANTRI"].tolist())
        if santri_terpilih != "-- Pilih Santri --":
            idx = df_santri[df_santri["NAMA SANTRI"] == santri_terpilih].index[0]
            dl = df_santri.iloc[idx]
            def g_k(t): return "" if t in ["🔴 BELUM LENGKAP", "[Belum Lengkap]"] else t

            e1, e2, e3 = st.columns(3)
            with e1:
                enama = st.text_input("Ubah NAMA", value=g_k(dl["NAMA SANTRI"]))
                einduk = st.text_input("Ubah NO INDUK", value=g_k(dl["NO INDUK"]))
                e_jk = st.selectbox("Ubah JK", ["PUTRA", "PUTRI"], index=0 if dl["JENIS KELAMIN"] == "PUTRA" else 1)
                enik = st.text_input("Ubah NIK", value=g_k(dl["NIK"]))
                ekk = st.text_input("Ubah KK", value=g_k(dl["KK"]))
            with e2:
                etmp = st.text_input("Ubah TEMPAT LAHIR", value=g_k(dl["TEMPAT LAHIR"]))
                try: etgl = st.date_input("Ubah TGL LAHIR", value=datetime.strptime(str(dl["TGL LAHIR"]), "%Y-%m-%d"), min_value=MIN_DATE, max_value=MAX_DATE)
                except: etgl = st.date_input("Ubah TGL LAHIR", value=datetime(2010, 1, 1), min_value=MIN_DATE, max_value=MAX_DATE)
                ekelas = st.selectbox("Ubah KELAS", list_kelas_terdeteksi, index=list_kelas_terdeteksi.index(dl["KELAS"]) if dl["KELAS"] in list_kelas_terdeteksi else 0)
                ekamar = st.selectbox("Ubah KAMAR", DAFTAR_KAMAR, index=DAFTAR_KAMAR.index(dl["KAMAR"]) if dl["KAMAR"] in DAFTAR_KAMAR else 0)
                estatus = st.selectbox("Ubah STATUS", ["Aktif", "Keluar", "Lulus"], index=["Aktif", "Keluar", "Lulus"].index(dl["STATUS"]))
            with e3:
                eayah = st.text_input("Ubah AYAH", value=g_k(dl["AYAH"]))
                eibu = st.text_input("Ubah IBU", value=g_k(dl["IBU"]))
                ewali = st.text_input("Ubah WALI", value=g_k(dl["WALI"]))
                edukuh = st.text_input("Ubah DUKUH", value=g_k(dl["DUKUH"]))
                edesa = st.text_input("Ubah DESA", value=g_k(dl["DESA"]))
                ekec = st.text_input("Ubah KECAMATAN", value=g_k(dl["KECAMATAN"]))
                ekab = st.text_input("Ubah KABUPATEN", value=g_k(dl["KABUPATEN"]))
                b1, b2 = st.columns([2, 4])
            with b1:
                if st.button("💾 Simpan Perubahan Santri", type="primary"):
                    df_santri.at[idx, "NISN"] = locals().get('e_nisn', locals().get('edit_nisn', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "NAMA"] = locals().get('e_nama', locals().get('edit_nama', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "JENIS KELAMIN"] = locals().get('e_jk', locals().get('edit_jk', 'PUTRA'))
                    df_santri.at[idx, "STATUS"] = locals().get('e_status', locals().get('edit_status', 'Aktif'))
                    df_santri.at[idx, "DUKUH"] = locals().get('e_dukuh', locals().get('edit_dukuh', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "DESA"] = locals().get('e_desa', locals().get('edit_desa', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "KECAMATAN"] = locals().get('e_kec', locals().get('edit_kec', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "KABUPATEN"] = locals().get('e_kab', locals().get('edit_kab', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "NIK"] = locals().get('e_nik', locals().get('edit_nik', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "KK"] = locals().get('e_kk', locals().get('edit_kk', '')) or "🔴 BELUM LENGKAP"
                    df_santri.at[idx, "TEMPAT LAHIR"] = locals().get('e_tmp', locals().get('edit_tmp', '')) or "🔴 BELUM LENGKAP"
                
                    # Pengaman untuk tanggal lahir
                    tgl_input = locals().get('e_tgl', locals().get('edit_tgl', None))
                    df_santri.at[idx, "TGL LAHIR"] = tgl_input.strftime("%Y-%m-%d") if hasattr(tgl_input, "strftime") else str(tgl_input or "")
                
                    df_santri.at[idx, "NO HP WA ORTU"] = locals().get('e_hp', locals().get('edit_hp', '')) or "🔴 BELUM LENGKAP"
                
                    # Proses simpan file
                    save_all(df_santri[KOLOM_SANTRI], FILE_SANTRI)
                    tampilkan_notifikasi_sukses()
                    st.rerun()
            with b2:
                if st.button(f"🗑️ HAPUS PERMANEN SANTRI: {santri_terpilih}"):
                    save_all(df_santri[KOLOM_SANTRI], FILE_SANTRI)
                    df_santri = df_santri.drop(idx).reset_index(drop=True)
                    df_santri.to_csv(FILE_SANTRI, index=False, sep=',', encoding='utf-8-sig')
                    tampilkan_notifikasi_sukses()
                    st.rerun()

        st.write("---")
        st.subheader("🚨 Fitur Bahaya (Reset Total Data Santri)")
        if st.button("⚠️ HAPUS SEMUA DATA SANTRI PERMANEN"):
            if os.path.exists(FILE_SANTRI):
                jalankan_auto_backup(FILE_SANTRI, "sebelum")
                os.remove(FILE_SANTRI)
            tampilkan_notifikasi_sukses()
            st.rerun()

# ------------------------------------------
# TAB 4: PEMETAAN DAERAH & KELAS (REVISI PRINT KE EXCEL MS ASLI)
# ------------------------------------------
with tab_daerah:
    st.header("📍 Pemetaan & Fitur Print Per Kelas")
    sub_fitur = st.radio("Pilih Jenis Penggolongan Alamat / Kelas:", ["Berdasarkan Wilayah Alamat", "Berdasarkan Kelas Santri"], horizontal=True)
    st.write("---")
    
    if not df_santri.empty and len(df_santri_clean) > 0:
        if sub_fitur == "Berdasarkan Wilayah Alamat":
            df_santri_clean["KECAMATAN"] = df_santri_clean["KECAMATAN"].str.upper()
            list_kec = sorted([x for x in df_santri_clean["KECAMATAN"].unique() if x and x != "🔴 BELUM LENGKAP"])
            pilihan_kec = st.selectbox("Pilih Kecamatan:", ["-- Pilih Kecamatan --"] + list_kec)
            if pilihan_kec != "-- Pilih Kecamatan --":
                df_hasil_d = df_santri_clean[df_santri_clean["KECAMATAN"] == pilihan_kec]
                st.dataframe(df_hasil_d[KOLOM_CETAK_DAERAH], use_container_width=True)
        else:
            pilihan_kls = st.selectbox("Pilih Kelas Yang Ingin Ditampilkan & Di-print:", ["-- Pilih Kelas --"] + list_kelas_terdeteksi)
            if pilihan_kls != "-- Pilih Kelas --":
                df_hasil_kelas = df_santri_clean[(df_santri_clean["KELAS"] == pilihan_kls) & (df_santri_clean["STATUS"] == "Aktif")]
                st.metric(label=f"Total Santri Aktif di {pilihan_kls}", value=f"{len(df_hasil_kelas)} Orang")
                
                # REVISI: GENERATE DAN DOWNLOAD SEBAGAI FILE EXCEL ASLI (.XLSX)
                excel_kelas_data = konversi_ke_excel_asli(df_hasil_kelas[KOLOM_SANTRI])
                st.download_button(
                    label=f"🟢 Download File Excel (.xlsx) data {pilihan_kls}",
                    data=excel_kelas_data,
                    file_name=f"Data_Santri_{pilihan_kls.replace(' ', '_')}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type="primary"
                )
                st.dataframe(df_hasil_kelas[["NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "AYAH", "DESA", "KAMAR", "HUBUNGAN"]], use_container_width=True)
            
            st.write("---")
            st.subheader("📋 Ringkasan Total Seluruh Kelas")
            df_hanya_aktif = df_santri_clean[df_santri_clean["STATUS"] == "Aktif"]
            if not df_hanya_aktif.empty:
                ringkasan_kelas = df_hanya_aktif["KELAS"].value_counts().reindex(list_kelas_terdeteksi, fill_value=0).reset_index()
                ringkasan_kelas.columns = ["Nama Kelas", "Jumlah Santri Aktif"]
                st.table(ringkasan_kelas)

# ------------------------------------------
# TAB 5: MANAJEMEN KAMAR (REVISI PRINT KE EXCEL MS ASLI)
# ------------------------------------------
with tab_kamar:
    st.header("🛏️ Manajemen Kamar & Fitur Print Per Kamar")
    km = st.selectbox("Pilih Ruang Kamar:", DAFTAR_KAMAR)
    
    st.subheader(f"👑 Struktur Pengurus {km}")
    cp1, cp2 = st.columns(2)
    with cp1:
        ketua_s = df_pengurus_kamar.loc[km, "KETUA"] if km in df_pengurus_kamar.index else "-"
        nama_ketua = st.text_input(f"Nama Ketua {km}:", value=ketua_s if ketua_s != "-" else "")
    with cp2:
        wakil_s = df_pengurus_kamar.loc[km, "WAKIL"] if km in df_pengurus_kamar.index else "-"
        nama_wakil = st.text_input(f"Nama Wakil {km}:", value=wakil_s if wakil_s != "-" else "")
        
    if st.button(f"💾 Perbarui Pengurus {km}"):
        df_pengurus_kamar.loc[km, "KETUA"] = nama_ketua.strip().upper() if nama_ketua else "-"
        df_pengurus_kamar.loc[km, "WAKIL"] = nama_wakil.strip().upper() if nama_wakil else "-"
        save_all(df_pengurus_kamar, FILE_PENGURUS_KAMAR)
        tampilkan_notifikasi_sukses()
        st.rerun()
        
    st.write("---")
    st.subheader(f"📋 Daftar Anggota {km}")
    if not df_santri.empty:
        dp = df_santri_clean[(df_santri_clean["KAMAR"] == km) & (df_santri_clean["STATUS"] == "Aktif")]
        st.metric(label=f"Total Santri di {km}", value=f"{len(dp)} Orang")
        
        # REVISI: GENERATE DAN DOWNLOAD SEBAGAI FILE EXCEL ASLI (.XLSX)
        excel_kamar_data = konversi_ke_excel_asli(dp[KOLOM_SANTRI])
        st.download_button(
            label=f"🟢 Download File Excel (.xlsx) data {km}",
            data=excel_kamar_data,
            file_name=f"Data_Santri_{km.replace(' ', '_')}.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary"
        )
        st.dataframe(dp[["NO INDUK", "NAMA SANTRI", "JENIS KELAMIN", "KELAS", "DESA", "KECAMATAN", "HUBUNGAN"]], use_container_width=True)

    st.write("---")
    st.subheader("📋 Ringkasan Total Seluruh Kamar")
    if not df_santri_clean.empty:
        df_h_aktif_km = df_santri_clean[df_santri_clean["STATUS"] == "Aktif"]
        if not df_h_aktif_km.empty:
            ringkasan_kamar = df_h_aktif_km["KAMAR"].value_counts().reindex(DAFTAR_KAMAR, fill_value=0).reset_index()
            ringkasan_kamar.columns = ["Nama Kamar", "Jumlah Santri Aktif"]
            st.table(ringkasan_kamar)

# ------------------------------------------
# TAB 6: DATA ASATIDZ (WITH MANUAL & IMPORT EXCEL/CSV) - FIX ANTI-LOCK
# ------------------------------------------
with tab_asatidz_panel:
    st.header("👳 Data & Administrasi Asatidz / Pengajar")
    
    u_as1, u_as2, u_as3 = st.columns([1.5, 1.5, 7])
    with u_as1:
        if st.button("↩️ Undo Aksi Asatidz", use_container_width=True):
            if eksekusi_undo(FILE_ASATIDZ):
                st.toast("Undo Berhasil!", icon="🔄")
                time.sleep(0.5)
                st.rerun()
            else: st.info("Tidak ada riwayat Undo")
    with u_as2:
        if st.button("↪️ Redo Aksi Asatidz", use_container_width=True):
            if eksekusi_redo(FILE_ASATIDZ):
                st.toast("Redo Berhasil!", icon="🔄")
                time.sleep(0.5)
                st.rerun()
            else: st.info("Tidak ada riwayat Redo")
            
    st.write("---")

    metode_asatidz = st.radio("Pilih Metode Input Data Asatidz:", ["Ketik Manual Satu per Satu", "Unggah / Impor Massal dari Excel (CSV)"], key="metode_as")
    
    if metode_asatidz == "Ketik Manual Satu per Satu":
        with st.form("form_asatidz", clear_on_submit=True):
            st.subheader("📝 Data Pribadi Pengajar")
            as1, as2, as3 = st.columns(3)
            with as1:
                niu_guru = st.text_input("NIU (Nomor Induk Ustadz)")
                nama_guru = st.text_input("NAMA USTADZ/AH*")
                jk_guru = st.selectbox("JENIS KELAMIN", ["USTADZ", "USTADZAH"])
                jabatan_guru = st.text_input("TUGAS/JABATAN")
                nik_guru = st.text_input("NIK GURU")
            with as2:
                kk_guru = st.text_input("KK GURU")
                tmp_lguru = st.text_input("TEMPAT LAHIR GURU")
                tgl_lguru = st.date_input("TANGGAL LAHIR GURU", value=datetime(1990, 1, 1), min_value=MIN_DATE, max_value=MAX_DATE)
                hp_guru = st.text_input("NO HP / WA")
                status_guru = st.selectbox("STATUS KEAKTIFAN", ["Aktif", "Keluar", "Lulus"])
            with as3:
                dk_guru = st.text_input("DUKUH GURU")
                ds_guru = st.text_input("DESA GURU")
                kc_guru = st.text_input("KECAMATAN GURU")
                kb_guru = st.text_input("KABUPATEN GURU")
                
            if st.form_submit_button("Simpan Data Asatidz Baru"):
                if nama_guru:
                    final_niu = str(niu_guru).strip() if niu_guru else f"GURU_{datetime.now().strftime('%M%S')}"
                    new_guru = pd.DataFrame([[
                        final_niu, nama_guru, jk_guru, jabatan_guru or "🔴 BELUM LENGKAP",
                        dk_guru or "🔴 BELUM LENGKAP", ds_guru or "🔴 BELUM LENGKAP", kc_guru or "🔴 BELUM LENGKAP", kb_guru or "🔴 BELUM LENGKAP", 
                        status_guru, str(nik_guru) or "🔴 BELUM LENGKAP", str(kk_guru) or "🔴 BELUM LENGKAP", tmp_lguru or "🔴 BELUM LENGKAP", 
                        tgl_lguru.strftime("%Y-%m-%d"), hp_guru or "🔴 BELUM LENGKAP"
                    ]], columns=KOLOM_ASATIDZ)
                    df_asatidz = pd.concat([df_asatidz, new_guru], ignore_index=True)
                    save_all(df_asatidz[KOLOM_ASATIDZ], FILE_ASATIDZ)
                    tampilkan_notifikasi_sukses()
                    st.rerun()
    else:
        st.subheader("📥 Impor Data Asatidz via Excel (CSV)")
        uploaded_file_as = st.file_uploader("Pilih file CSV hasil Excel Data Asatidz Anda:", type=["csv"], key="upload_as")
        if uploaded_file_as is not None:
            try:
                try: df_upload_as = pd.read_csv(uploaded_file_as, sep=None, engine='python', encoding='utf-8-sig')
                except: df_upload_as = pd.read_csv(uploaded_file_as, sep=';', encoding='utf-8-sig')
                df_upload_as.columns = df_upload_as.columns.str.upper().str.strip()
                
                if "NAMA LENGKAP" in df_upload_as.columns and "NAMA USTADZ/AH" not in df_upload_as.columns:
                    df_upload_as = df_upload_as.rename(columns={"NAMA LENGKAP": "NAMA USTADZ/AH"})
                
                if "NAMA USTADZ/AH" in df_upload_as.columns:
                    df_upload_as = df_upload_as.fillna("🔴 BELUM LENGKAP")
                    if "NIU" not in df_upload_as.columns: df_upload_as["NIU"] = "🔴 BELUM LENGKAP"
                        
                    for col in KOLOM_ASATIDZ:
                        if col not in df_upload_as.columns:
                            if col == "STATUS": df_upload_as[col] = "Aktif"
                            elif col == "JENIS KELAMIN": df_upload_as[col] = "USTADZ"
                            else: df_upload_as[col] = "🔴 BELUM LENGKAP"
                        else:
                            df_upload_as[col] = df_upload_as[col].astype(str).str.replace("nan", "🔴 BELUM LENGKAP", regex=False).str.strip()
                    
                    if "STATUS" in df_upload_as.columns:
                        df_upload_as["STATUS"] = df_upload_as["STATUS"].str.strip().str.upper().replace({"AKTIF": "Aktif", "KELUAR": "Keluar", "LULUS": "Lulus"})
                        df_upload_as.loc[~df_upload_as["STATUS"].isin(["Aktif", "Keluar", "Lulus"]), "STATUS"] = "Aktif"
                    
                    if "JENIS KELAMIN" in df_upload_as.columns:
                        df_upload_as["JENIS KELAMIN"] = df_upload_as["JENIS KELAMIN"].str.strip().str.upper()
                        df_upload_as.loc[df_upload_as["JENIS KELAMIN"].isin(["USTADZ"]), "JENIS KELAMIN"] = "USTADZ"
                        df_upload_as.loc[df_upload_as["JENIS KELAMIN"].isin(["USTADZAH"]), "JENIS KELAMIN"] = "USTADZAH"

                    if "TGL LAHIR" in df_upload_as.columns:
                        df_upload_as["TGL LAHIR"] = df_upload_as["TGL LAHIR"].apply(bersihkan_tanggal_indo)
                            
                    df_upload_as = df_upload_as[KOLOM_ASATIDZ]
                    st.write("### 👀 Pratinjau Data yang Akan Diimpor:")
                    st.dataframe(df_upload_as.head(10), use_container_width=True)
                    
                    if st.button("🚀 Masukkan Semua Data Asatidz ke Aplikasi", type="primary"):
                        df_asatidz = df_upload_as if (df_asatidz.empty or "🔴 BELUM LENGKAP" in df_asatidz["NAMA USTADZ/AH"].values) else pd.concat([df_asatidz, df_upload_as], ignore_index=True)
                        save_all(df_asatidz[KOLOM_ASATIDZ], FILE_ASATIDZ)
                        tampilkan_notifikasi_sukses()
                        st.rerun()
                else:
                    st.error("❌ Kolom 'NAMA USTADZ/AH' atau 'NAMA LENGKAP' tidak ditemukan dalam file Anda.")
            except Exception as e: 
                st.error(f"Gagal membaca file asatidz: {e}")

    st.write("---")
    st.subheader("📋 Data Seluruh Asatidz Saat Ini")
    
    df_asatidz_clean = df_asatidz[df_asatidz["NAMA USTADZ/AH"].notna() & (df_asatidz["NAMA USTADZ/AH"] != "") & (df_asatidz["NAMA USTADZ/AH"] != "🔴 BELUM LENGKAP")]
    
    if not df_asatidz_clean.empty:
        st.dataframe(df_asatidz_clean, use_container_width=True)
        
        st.write("---")
        st.subheader("🛠️ Ubah Data Pribadi / Hapus Asatidz Tunggal")
        guru_terpilih = st.selectbox("Pilih Nama Ustadz/Ustadzah:", ["-- Pilih Asatidz --"] + df_asatidz_clean["NAMA USTADZ/AH"].tolist())
        
        # 1. Kita bikin fungsi khusus biar spasinya aman sentosa
    def proses_edit_asatidz(guru_terpilih, df_asatidz):
        idx_g = df_asatidz[df_asatidz["NAMA USTADZ/AH"] == guru_terpilih].index[0]
        dg = df_asatidz.iloc[idx_g]
        def g_k_g(t): return "" if t in ["🔴 BELUM LENGKAP", "[Belum Lengkap]"] else t

        eg1, eg2, eg3 = st.columns(3)
        with eg1:
            eg_nama = st.text_input("Ubah Nama Ustadz/ah", value=g_k_g(dg["NAMA USTADZ/AH"]))
            eg_niu = st.text_input("Ubah NIU", value=g_k_g(dg["NIU"]))
            eg_jk = st.selectbox("Ubah Jenis Kelamin", ["USTADZ", "USTADZAH"], index=0 if dg["JENIS KELAMIN"] == "USTADZ" else 1)
            eg_jbt = st.text_input("Ubah Tugas/Jabatan", value=g_k_g(dg["TUGAS/JABATAN"]))
            eg_nik = st.text_input("Ubah NIK", value=g_k_g(dg["NIK"]))
        with eg2:
            eg_kk = st.text_input("Ubah KK", value=g_k_g(dg["KK"]))
            eg_tmp = st.text_input("Ubah Tempat Lahir", value=g_k_g(dg["TEMPAT LAHIR"]))
            try: eg_tgl = st.date_input("Ubah Tanggal Lahir", value=datetime.strptime(str(dg["TGL LAHIR"]), "%Y-%m-%d"), min_value=MIN_DATE, max_value=MAX_DATE)
            except: eg_tgl = st.date_input("Ubah Tanggal Lahir", value=datetime(1990, 1, 1), min_value=MIN_DATE, max_value=MAX_DATE)
            eg_hp = st.text_input("Ubah No HP", value=g_k_g(dg["NO HP"]))
            eg_status = st.selectbox("Ubah Status Keaktifan", ["Aktif", "Keluar", "Lulus"], index=["Aktif", "Keluar", "Lulus"].index(dg["STATUS"]))
        with eg3:
            eg_dk = st.text_input("Ubah Dukuh", value=g_k_g(dg["DUKUH"]))
            eg_ds = st.text_input("Ubah Desa", value=g_k_g(dg["DESA"]))
            eg_kc = st.text_input("Ubah Kecamatan", value=g_k_g(dg["KECAMATAN"]))
            eg_kb = st.text_input("Ubah Kabupaten", value=g_k_g(dg["KABUPATEN"]))
            
        b1_as, b2_as = st.columns([2, 4])
        with b1_as:
            if st.button("💾 Simpan Perubahan Asatidz", type="primary"):
                df_asatidz.iloc[idx_g] = [
                    eg_niu or "🔴 BELUM LENGKAP", 
                    eg_nama or "🔴 BELUM LENGKAP", 
                    eg_jk, 
                    eg_jbt or "🔴 BELUM LENGKAP",
                    eg_dk or "🔴 BELUM LENGKAP", 
                    eg_ds or "🔴 BELUM LENGKAP", 
                    eg_kc or "🔴 BELUM LENGKAP", 
                    eg_kb or "🔴 BELUM LENGKAP",
                    eg_status, 
                    eg_nik or "🔴 BELUM LENGKAP", 
                    eg_kk or "🔴 BELUM LENGKAP", 
                    eg_tmp or "🔴 BELUM LENGKAP", 
                    eg_tgl.strftime("%Y-%m-%d"), 
                    eg_hp or "🔴 BELUM LENGKAP",
                    "Mandiri"
                ]
                save_all(df_asatidz[KOLOM_ASATIDZ], FILE_ASATIDZ)
                tampilkan_notifikasi_sukses()
                st.rerun()

    # 2. Ini bagian pemicunya, pastikan tulisan 'if' ini sejajar dengan menu atasnya
    if 'guru_terpilih' in locals() or 'guru_terpilih' in globals():
        if guru_terpilih != "-- Pilih Asatidz --":
            proses_edit_asatidz(guru_terpilih, df_asatidz)
    else:
        st.info("Belum ada data asatidz/pengajar aktif yang tersimpan di database.")

# Footer Halaman
st.write("---")
st.caption("Aplikasi Administrasi Pondok Pesantren Terpadu v2.6 • Sistem Anti-Error Massal Impor Excel Asatidz")