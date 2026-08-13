import os
import re
import math
import pandas as pd
import streamlit as st
from supabase import create_client, Client, ClientOptions

# Konfigurasi Halaman
st.set_page_config(page_title="PLN FLC Data Engine", page_icon="⚡", layout="wide")
st.title("⚡ PLN Functional Location Data Engine")

# Inisialisasi Supabase
SUPABASE_URL = os.getenv("SUPABASE_URL") or st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY", "")

supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.error(f"Supabase init error: {str(e)}")

# Fungsi Pembersih Data (Agar Tidak Error JSON / Nan)
def clean_str(val):
    if pd.isna(val) or val is None: return None
    s_val = str(val).strip()
    if s_val.lower() in ['nan', 'none', 'null', '']: return None
    return s_val

def clean_float(val):
    if pd.isna(val) or val is None: return None
    try:
        if isinstance(val, str): val = val.replace(',', '.')
        f_val = float(val)
        return None if math.isnan(f_val) or math.isinf(f_val) else f_val
    except: return None

def clean_date(val):
    if pd.isna(val) or val is None: return None
    try:
        dt = pd.to_datetime(val, errors='coerce')
        return None if pd.isna(dt) else dt.strftime('%Y-%m-%d')
    except: return None

# Fungsi Bisnis PLN
def normalize_function_code(code_val, level_val, desc_val=""):
    code = clean_str(code_val)
    if not code: return None
    desc = clean_str(desc_val) or ""
    level = clean_float(level_val) or 0.0
    if code == 'G' and (level == 4.0 or 'SERANDANG' in desc.upper()): return 'SG'
    mapping = {('V',3.0):'V1', ('V',4.0):'V2', ('X',3.5):'X1', ('X',4.0):'X2', ('O',4.0):'O1', ('O',2.0):'O2', ('Q',3.5):'Q1', ('Q',4.0):'Q2'}
    return mapping.get((code, level), code)

def determine_is_active(status_val):
    cleaned = clean_str(status_val)
    return cleaned.split('.')[0].upper() in ['2', '4', '5', '6', '7', '8'] if cleaned else False

# Main Logic
uploaded_file = st.file_uploader("Upload Excel FLC PLN", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    if "ID_FUNCTLOC" not in df.columns:
        df.columns = df.iloc[1].values
        df = df.iloc[2:].copy()
    
    df_cleaned = df[df['KD_FUNGSI'].astype(str).str.strip() != 'X'].copy()
    
    # Mapping Data
    mapped_data = []
    ref_flc_data = []
    
    for _, row in df_cleaned.iterrows():
        flc_id = clean_str(row.get('ID_FUNCTLOC'))
        if not flc_id: continue
        
        status_code = clean_str(row.get('STATUS'))
        fn_code = normalize_function_code(row.get('KD_FUNGSI'), row.get('NLEVEL'), row.get('DESKRIPSI'))
        loc_name = clean_str(row.get('NM_LOKASI'))
        
        # Referensi ref_flc
        ref_flc_data.append({"flc_id": flc_id, "name": loc_name, "function_code": fn_code, "is_active": determine_is_active(status_code)})
        
        # Data Mst_functloc (mapping lengkap)
        mapped_data.append({
            "functloc_id": flc_id,
            "sup_functloc_id": re.sub(r'-GR\d+$', '', clean_str(row.get('SUP_FUNCTLOC')) or ""),
            "location_name": loc_name,
            "description": clean_str(row.get('DESKRIPSI')),
            "short_name": clean_str(row.get('NMSINGKT')), # MAP NMSINGKT KE SHORT NAME
            "address": clean_str(row.get('ALAMAT')),
            "city": clean_str(row.get('KOTA')),
            "latitude": clean_float(row.get('LATITUDE')),
            "longitude": clean_float(row.get('LONGITUDE')),
            "slo_date": clean_date(row.get('TGL_SLO')),
            "slo_number": clean_str(row.get('NO_SLO')),
            "status_code": status_code,
            "function_code": fn_code,
            "voltage_code": clean_str(row.get('TEGANGAN')),
            "region_code": clean_str(row.get('KD_WILAYAH')),
            "grouplokasi_code": clean_str(row.get('KD_GROUPLOKASI')),
            "baygroup_code": clean_str(row.get('BAYGROUP')),
            "unit_code": clean_str(row.get('UNIT')),
            "plant_id": clean_str(row.get('ID_PLANT')),
            "operational_date": clean_date(row.get('TGL_OPRS')),
            "ownership": clean_str(row.get('MILIK'))
        })

    if st.button("Sync to Supabase"):
        # Logic 3-step sync (seperti kode sebelumnya)
        st.success("Sync Started...")
        # (Tambahkan loop 3 step sync disini seperti kode sebelumnya)
