import os
import re
import math
import pandas as pd
import streamlit as st
from supabase import create_client, Client, ClientOptions

# ========================================================
# 1. KONFIGURASI HALAMAN
# ========================================================
st.set_page_config(page_title="PLN Data Importer", page_icon="⚡", layout="wide")
st.title("⚡ PLN Functional Location Data Importer")

# Inisialisasi Supabase
SUPABASE_URL = os.getenv("SUPABASE_URL") or st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY", "")
supabase = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

# Session State Auth
if "user" not in st.session_state: st.session_state["user"] = None
if "session" not in st.session_state: st.session_state["session"] = None

# ========================================================
# 2. SIDEBAR (AUTH & STATUS)
# ========================================================
with st.sidebar:
    st.header("⚙️ System Status")
    st.success("Database Connected") if supabase else st.error("Database Disconnected")
    st.divider()
    st.header("🔐 User Authentication")
    if st.session_state["user"]:
        st.write(f"Login: **{st.session_state['user'].email}**")
        if st.button("Logout"):
            st.session_state["user"] = None
            st.session_state["session"] = None
            st.rerun()
    else:
        email = st.text_input("Email")
        pwd = st.text_input("Password", type="password")
        if st.button("Sign In"):
            try:
                res = supabase.auth.sign_in_with_password({"email": email, "password": pwd})
                st.session_state["user"] = res.user
                st.session_state["session"] = res.session
                st.rerun()
            except Exception as e:
                st.error("Login Gagal")

# ========================================================
# 3. FUNGSI SANITASI & NORMALISASI (FIXES RESTORED)
# ========================================================
def clean_str(val):
    if pd.isna(val) or val is None: return None
    s = str(val).strip()
    return None if s.lower() in ['nan', 'none', 'null', ''] else s

def clean_float(val):
    if pd.isna(val) or val is None: return None
    try:
        val = str(val).replace(',', '.')
        f = float(val)
        return None if math.isnan(f) or math.isinf(f) else f
    except: return None

def clean_date(val):
    if pd.isna(val) or val is None: return None
    try:
        dt = pd.to_datetime(val, errors='coerce')
        return None if pd.isna(dt) else dt.strftime('%Y-%m-%d')
    except: return None

def normalize_function_code(code_val, level_val, desc_val=""):
    code = clean_str(code_val)
    if not code: return None
    desc = clean_str(desc_val) or ""
    level = clean_float(level_val) or 0.0
    if code == 'G' and (level == 4.0 or 'SERANDANG' in desc.upper()): return 'SG'
    mapping = {('V',3.0):'V1', ('V',4.0):'V2', ('X',3.5):'X1', ('X',4.0):'X2', ('O',4.0):'O1', ('O',2.0):'O2', ('Q',3.5):'Q1', ('Q',4.0):'Q2'}
    return mapping.get((code, level), code)

def determine_is_active(status_val):
    c = clean_str(status_val)
    return c.split('.')[0].upper() in ['2', '4', '5', '6', '7', '8'] if c else False

# ========================================================
# 4. PROSES IMPORT
# ========================================================
uploaded_file = st.file_uploader("Upload Excel Master Data", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    if "ID_FUNCTLOC" not in df.columns:
        df.columns = df.iloc[1].values
        df = df.iloc[2:].copy()
    
    if st.button("Jalankan Import Data"):
        if not st.session_state["user"]:
            st.error("Harap login di Sidebar!")
            st.stop()
            
        access_token = st.session_state["session"].access_token
        auth_supabase = create_client(SUPABASE_URL, SUPABASE_KEY, 
                        options=ClientOptions(headers={"Authorization": f"Bearer {access_token}"}))

        with st.spinner("Processing..."):
            mapped_data, ref_flc_data = [], []
            
            # Sanitasi Parent ID (Validasi vs Database/Excel)
            valid_ids = set(df['ID_FUNCTLOC'].dropna().astype(str).str.strip())
            
            for _, row in df.iterrows():
                flc_id = clean_str(row.get('ID_FUNCTLOC'))
                if not flc_id: continue
                
                # Normalisasi Parent
                raw_sup = clean_str(row.get('SUP_FUNCTLOC'))
                sup_id = re.sub(r'-GR\d+$', '', raw_sup) if raw_sup else None
                sup_id = sup_id if sup_id in valid_ids else None
                
                status = clean_str(row.get('STATUS'))
                fn_code = normalize_function_code(row.get('KD_FUNGSI'), row.get('NLEVEL'), row.get('DESKRIPSI'))
                loc_name = clean_str(row.get('NM_LOKASI'))
                
                # Ref data
                ref_flc_data.append({"flc_id": flc_id, "name": loc_name, "function_code": fn_code, "is_active": determine_is_active(status)})
                
                # Mst data
                mapped_data.append({
                    "functloc_id": flc_id,
                    "sup_functloc_id": sup_id,
                    "location_name": loc_name,
                    "short_name": clean_str(row.get('NMSINGKT')),
                    "description": clean_str(row.get('DESKRIPSI')),
                    "status_code": status,
                    "address": clean_str(row.get('ALAMAT')),
                    "city": clean_str(row.get('KOTA')),
                    "latitude": clean_float(row.get('LATITUDE')),
                    "longitude": clean_float(row.get('LONGITUDE')),
                    "slo_date": clean_date(row.get('TGL_SLO')),
                    "slo_number": clean_str(row.get('NO_SLO')),
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

            # 3-Step Sync
            total = len(mapped_data)
            batch = 500
            for i in range(0, total, batch):
                auth_supabase.table('ref_flc').upsert(ref_flc_data[i:i+batch]).execute()
            for i in range(0, total, batch):
                data = [dict(item, sup_functloc_id=None) for item in mapped_data[i:i+batch]]
                auth_supabase.table('mst_functloc').upsert(data).execute()
            for i in range(0, total, batch):
                auth_supabase.table('mst_functloc').upsert(mapped_data[i:i+batch]).execute()
                
        st.success("Import Berhasil dengan data yang ter-normalisasi!")
