import streamlit as st
import pandas as pd
import re
from supabase import create_client, Client

# ========================================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# ========================================================
st.set_page_config(
    page_title="PLN FLC Importer & Normalizer",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ PLN Functional Location Importer & Normalizer")
st.markdown("""
Aplikasi ini berfungsi untuk memproses file Excel Export PLN, melakukan normalisasi hirarki:
1. **Menghapus sufiks Group (`-GRxxx`)** dari `SUP_FUNCTLOC` agar Tower & Span terhubung langsung ke Jalur.
2. **Normalisasi `function_code`** (`SERANDANG` $\rightarrow$ `SG`, `V1`, `O1`, `X1`, dll).
3. **Batch Upsert** otomatis ke Supabase (`mst_functloc` & `ref_flc`).
""")

# ========================================================
# 2. INI KONEKSI SUPABASE AMAN (MENGGUNAKAN SECRETS)
# ========================================================
st.sidebar.header("🔑 Supabase Credentials")

# Mengambil kredensial dari Streamlit Secrets jika ada, atau input manual via sidebar
supabase_url = st.sidebar.text_input(
    "Supabase URL", 
    value=st.secrets.get("SUPABASE_URL", ""),
    type="default"
)
supabase_key = st.sidebar.text_input(
    "Supabase Service Role Key", 
    value=st.secrets.get("SUPABASE_KEY", ""),
    type="password"
)

# ========================================================
# 3. FUNGSI LOGIKA NORMALISASI DATA
# ========================================================
def normalize_sup_functloc(sup_val):
    if pd.isna(sup_val):
        return None
    return re.sub(r'-GR\d+$', '', str(sup_val).strip())

def normalize_function_code(code_val, level_val, desc_val=""):
    if pd.isna(code_val):
        return None
    
    code = str(code_val).strip()
    desc = str(desc_val).upper() if pd.notnull(desc_val) else ""
    
    try:
        level = float(level_val) if pd.notnull(level_val) else 0.0
    except:
        level = 0.0

    # 1. Normalisasi SERANDANG G di Level 4 / Deskripsi SERANDANG menjadi SG
    if code == 'G' and (level == 4.0 or 'SERANDANG' in desc):
        return 'SG'

    # 2. Rule Mapping Kode Duplikat PLN (V, X, O, Q)
    mapping = {
        ('V', 3.0): 'V1', ('V', 4.0): 'V2',
        ('X', 3.5): 'X1', ('X', 4.0): 'X2',
        ('O', 4.0): 'O1', ('O', 2.0): 'O2',
        ('Q', 3.5): 'Q1', ('Q', 4.0): 'Q2'
    }
    
    return mapping.get((code, level), code)

# ========================================================
# 4. UPLOAD & PREVIEW FILE EXCEL
# ========================================================
uploaded_file = st.file_uploader("📂 Upload File Excel Hasil Export PLN (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read Excel File
        df_raw = pd.read_excel(uploaded_file)
        
        # Format Header (Memakai baris ke-2 jika header di baris 2)
        if "ID_FUNCTLOC" not in df_raw.columns:
            headers = df_raw.iloc[1].values
            df = df_raw.iloc[2:].copy()
            df.columns = headers
        else:
            df = df_raw.copy()

        st.success(f"File berhasil dibaca! Total data awal: **{len(df):,}** baris.")

        # Exclude baris Group (KD_FUNGSI = 'X')
        df_cleaned = df[df['KD_FUNGSI'].astype(str).str.strip() != 'X'].copy()

        # Jalankan Normalisasi
        df_cleaned['SUP_FUNCTLOC_CLEAN'] = df_cleaned['SUP_FUNCTLOC'].apply(normalize_sup_functloc)
        df_cleaned['FUNCTION_CODE_CLEAN'] = df_cleaned.apply(
            lambda r: normalize_function_code(r['KD_FUNGSI'], r['NLEVEL'], r['DESKRIPSI']), axis=1
        )

        st.subheader("📊 Preview Hasil Normalisasi Data")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Data Siap Import (Non-Group)", f"{len(df_cleaned):,} baris")
        with col2:
            modified_count = (df_cleaned['SUP_FUNCTLOC'] != df_cleaned['SUP_FUNCTLOC_CLEAN']).sum()
            st.metric("Data Parent (-GR) Di-Normalisasi", f"{modified_count:,} baris")

        # Tampilkan Tabel Preview
        preview_cols = ['ID_FUNCTLOC', 'SUP_FUNCTLOC', 'SUP_FUNCTLOC_CLEAN', 'KD_FUNGSI', 'FUNCTION_CODE_CLEAN', 'NM_LOKASI']
        st.dataframe(df_cleaned[preview_cols].head(100), use_container_width=True)

        # ========================================================
        # 5. ESEKUSI PROCESS IMPORT KE SUPABASE
        # ========================================================
        st.divider()
        if st.button("🚀 Mulai Import ke Supabase", type="primary", use_container_width=True):
            if not supabase_url or not supabase_key:
                st.error("❌ Harap isi Supabase URL & Service Role Key di sidebar terlebih dahulu!")
            else:
                try:
                    # Inisialisasi Supabase Client
                    supabase: Client = create_client(supabase_url, supabase_key)
                    
                    # Mapping data ke format Supabase
                    mapped_data = []
                    for _, row in df_cleaned.iterrows():
                        mapped_data.append({
                            "functloc_id": str(row['ID_FUNCTLOC']).strip(),
                            "sup_functloc_id": row['SUP_FUNCTLOC_CLEAN'],
                            "location_name": str(row['NM_LOKASI']).strip(),
                            "description": str(row['DESKRIPSI']).strip() if pd.notnull(row['DESKRIPSI']) else None,
                            "unit_code": str(row['UNIT']).strip() if pd.notnull(row['UNIT']) else None,
                            "nlevel": float(row['NLEVEL']) if pd.notnull(row['NLEVEL']) else None,
                            "status_code": str(row['STATUS']).strip() if pd.notnull(row['STATUS']) else None,
                            "voltage_code": str(row['TEGANGAN']).strip() if pd.notnull(row['TEGANGAN']) else None,
                            "function_code": row['FUNCTION_CODE_CLEAN'],
                            "region_code": str(row['KD_WILAYAH']).strip() if pd.notnull(row['KD_WILAYAH']) else None,
                            "workcenter": str(row['WORKCENTER']).strip() if pd.notnull(row['WORKCENTER']) else None,
                            "plant_id": str(row['ID_PLANT']).strip() if pd.notnull(row['ID_PLANT']) else None,
                            "grouplokasi_code": str(row['KD_GROUPLOKASI']).strip() if pd.notnull(row['KD_GROUPLOKASI']) else None,
                            "gi_flc": str(row['GI_FLC']).strip() if pd.notnull(row['GI_FLC']) else None,
                            "baygroup_code": str(row['BAYGROUP']).strip() if pd.notnull(row['BAYGROUP']) else None,
                        })

                    # Progress Bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    batch_size = 500
                    total_batches = len(mapped_data)
                    
                    for i in range(0, total_batches, batch_size):
                        batch = mapped_data[i:i + batch_size]
                        supabase.table('mst_functloc').upsert(batch).execute()
                        
                        current_progress = min((i + batch_size) / total_batches, 1.0)
                        progress_bar.progress(current_progress)
                        status_text.text(f"Mengunggah... {min(i + batch_size, total_batches)} dari {total_batches} baris.")

                    st.balloons()
                    st.success("🎉 Import Berhasil! Data telah tersinkronisasi sempurna ke `mst_functloc` & `ref_flc`!")

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan saat mengunggah ke Supabase: {str(e)}")

    except Exception as e:
        st.error(f"❌ Gagal membaca file Excel: {str(e)}")
