import os
import re
import pandas as pd
import streamlit as st
from supabase import create_client, Client, ClientOptions

# ========================================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# ========================================================
st.set_page_config(
    page_title="PLN FLC Data Engine",
    page_icon="⚡",
    layout="wide"
)

# Header Profesional
st.title("⚡ PLN Functional Location Data Engine")
st.caption("EAM Data Transformation & Automated Synchronization Pipeline")

st.divider()

# ========================================================
# 2. INISIALISASI SUPABASE CLIENT & AUTH SESSION
# ========================================================
SUPABASE_URL = os.getenv("SUPABASE_URL") or st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY") or st.secrets.get("SUPABASE_KEY", "")

supabase: Client = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.error(f"Failed to initialize Supabase Client: {str(e)}")

if "user" not in st.session_state:
    st.session_state["user"] = None
if "session" not in st.session_state:
    st.session_state["session"] = None

# Sidebar Authorization Panel
with st.sidebar:
    st.header("⚙️ System Status")
    if supabase:
        st.success("Database Connected")
    else:
        st.error("Database Disconnected")
        st.caption("Configure `SUPABASE_URL` & `SUPABASE_KEY` in Settings -> Secrets.")

    st.divider()
    st.header("🔐 User Authentication")
    
    if st.session_state["user"]:
        st.success(f"Logged in as:\n**{st.session_state['user'].email}** 🟢")
        if st.button("Logout", type="secondary", use_container_width=True):
            st.session_state["user"] = None
            st.session_state["session"] = None
            try:
                supabase.auth.sign_out()
            except:
                pass
            st.rerun()
    else:
        st.info("Public Mode: View & Audit Only 👁️")
        email_input = st.text_input("Email:", placeholder="user@pln.co.id")
        pwd_input = st.text_input("Password:", type="password")
        
        if st.button("Sign In", type="primary", use_container_width=True):
            if not supabase:
                st.error("Database Client not initialized.")
            elif not email_input or not pwd_input:
                st.warning("Please enter both Email and Password.")
            else:
                try:
                    response = supabase.auth.sign_in_with_password({
                        "email": email_input.strip(),
                        "password": pwd_input
                    })
                    if response.user and response.session:
                        st.session_state["user"] = response.user
                        st.session_state["session"] = response.session
                        st.success("Authentication Successful!")
                        st.rerun()
                except Exception as err:
                    st.error(f"Login Failed: {str(err)}")

# ========================================================
# 3. FUNGSI TRANSFOMASI & NORMALISASI DATA
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

    if code == 'G' and (level == 4.0 or 'SERANDANG' in desc):
        return 'SG'

    mapping = {
        ('V', 3.0): 'V1', ('V', 4.0): 'V2',
        ('X', 3.5): 'X1', ('X', 4.0): 'X2',
        ('O', 4.0): 'O1', ('O', 2.0): 'O2',
        ('Q', 3.5): 'Q1', ('Q', 4.0): 'Q2'
    }
    
    return mapping.get((code, level), code)

# ========================================================
# 4. UPLOAD & BACA FILE EXCEL
# ========================================================
uploaded_file = st.file_uploader("Upload Source Dataset (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        if "ID_FUNCTLOC" not in df_raw.columns:
            headers = df_raw.iloc[1].values
            df = df_raw.iloc[2:].copy()
            df.columns = headers
        else:
            df = df_raw.copy()

        # Exclude baris perantara Group (KD_FUNGSI = 'X')
        df_cleaned = df[df['KD_FUNGSI'].astype(str).str.strip() != 'X'].copy()

        # Urutkan NLEVEL Ascending
        df_cleaned['NLEVEL_NUM'] = pd.to_numeric(df_cleaned['NLEVEL'], errors='coerce').fillna(99)
        df_cleaned = df_cleaned.sort_values(by='NLEVEL_NUM', ascending=True)

        # Eksekusi Normalisasi
        df_cleaned['SUP_FUNCTLOC_CLEAN'] = df_cleaned['SUP_FUNCTLOC'].apply(normalize_sup_functloc)
        df_cleaned['FUNCTION_CODE_CLEAN'] = df_cleaned.apply(
            lambda r: normalize_function_code(r['KD_FUNGSI'], r['NLEVEL'], r['DESKRIPSI']), axis=1
        )

        # Metrics Calculation
        total_data = len(df_cleaned)
        gr_modified = (df_cleaned['SUP_FUNCTLOC'] != df_cleaned['SUP_FUNCTLOC_CLEAN']).sum()
        sg_modified = (df_cleaned['FUNCTION_CODE_CLEAN'] == 'SG').sum()
        mapped_codes = df_cleaned['FUNCTION_CODE_CLEAN'].isin(['V1', 'V2', 'O1', 'O2', 'X1', 'X2', 'Q1', 'Q2']).sum()

        # ========================================================
        # 5. EXECUTIVE SUMMARY METRICS
        # ========================================================
        st.subheader("Data Processing Summary")
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Valid FLC Records", f"{total_data:,}")
        m2.metric("Flattened Group Parents", f"{gr_modified:,}")
        m3.metric("Serandang Normalized (SG)", f"{sg_modified:,}")
        m4.metric("Function Codes Re-mapped", f"{mapped_codes:,}")

        # ========================================================
        # 6. INTERACTIVE REVIEW TABLE
        # ========================================================
        st.divider()
        st.subheader("Dataset Review & Audit")

        search_query = st.text_input("Filter Records:", "", placeholder="Type FLC ID, Location Name, Substation, Tower...")
        
        if search_query:
            search_pattern = search_query.strip()
            df_display = df_cleaned[
                df_cleaned['ID_FUNCTLOC'].astype(str).str.contains(search_pattern, case=False, na=False) |
                df_cleaned['NM_LOKASI'].astype(str).str.contains(search_pattern, case=False, na=False) |
                df_cleaned['SUP_FUNCTLOC'].astype(str).str.contains(search_pattern, case=False, na=False)
            ]
        else:
            df_display = df_cleaned

        tab_choice = st.radio(
            "View Mode:", 
            ["Normalized Output", "Hierarchy Diff (Group Removal)", "Full Raw Dataset"],
            horizontal=True
        )

        if tab_choice == "Normalized Output":
            cols_review = ['ID_FUNCTLOC', 'SUP_FUNCTLOC_CLEAN', 'NM_LOKASI', 'FUNCTION_CODE_CLEAN', 'NLEVEL', 'STATUS', 'TEGANGAN', 'WORKCENTER', 'ID_PLANT']
            st.dataframe(
                df_display[cols_review].rename(columns={
                    'SUP_FUNCTLOC_CLEAN': 'SUP_FUNCTLOC (NORMALIZED)',
                    'FUNCTION_CODE_CLEAN': 'FUNCTION_CODE (MAPPED)'
                }), 
                use_container_width=True,
                height=400
            )

        elif tab_choice == "Hierarchy Diff (Group Removal)":
            cols_comp = ['ID_FUNCTLOC', 'SUP_FUNCTLOC', 'SUP_FUNCTLOC_CLEAN', 'KD_FUNGSI', 'FUNCTION_CODE_CLEAN', 'NM_LOKASI']
            df_changed = df_display[df_display['SUP_FUNCTLOC'] != df_display['SUP_FUNCTLOC_CLEAN']]
            st.caption(f"Showing {len(df_changed):,} modified parent relations:")
            st.dataframe(df_changed[cols_comp], use_container_width=True, height=400)

        else:
            st.dataframe(df_display, use_container_width=True, height=400)

        # ========================================================
        # 7. EXECUTE SYNC (3-STEP PIPELINE: REF_FLC -> MST_PASS_1 -> MST_PASS_2)
        # ========================================================
        st.divider()
        
        if st.session_state["user"] and st.session_state["session"]:
            st.success(f"🔓 Authorized Session Active ({st.session_state['user'].email}): Data Synchronization Privileges Granted.")
            
            if st.button("Synchronize to Supabase", type="primary", use_container_width=True):
                if not supabase:
                    st.error("Database connection unavailable.")
                else:
                    try:
                        access_token = st.session_state["session"].access_token

                        auth_supabase = create_client(
                            SUPABASE_URL, 
                            SUPABASE_KEY, 
                            options=ClientOptions(headers={"Authorization": f"Bearer {access_token}"})
                        )

                        # Data untuk ref_flc
                        ref_flc_data = []
                        # Data untuk mst_functloc
                        mapped_data = []

                        for _, row in df_cleaned.iterrows():
                            flc_id = str(row['ID_FUNCTLOC']).strip()
                            loc_name = str(row['NM_LOKASI']).strip()

                            ref_flc_data.append({
                                "flc_id": flc_id,
                                "flc_name": loc_name
                            })

                            mapped_data.append({
                                "functloc_id": flc_id,
                                "sup_functloc_id": row['SUP_FUNCTLOC_CLEAN'],
                                "location_name": loc_name,
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

                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        batch_size = 500
                        total_records = len(mapped_data)

                        # ----------------------------------------------------
                        # TAHAP 1: Sync ke tabel ref_flc
                        # ----------------------------------------------------
                        status_text.text("Step 1/3: Synchronizing reference table (ref_flc)...")
                        for i in range(0, total_records, batch_size):
                            batch_ref = ref_flc_data[i:i + batch_size]
                            auth_supabase.table('ref_flc').upsert(batch_ref).execute()
                            
                            prog = ((i + batch_size) / total_records) * 0.33
                            progress_bar.progress(min(prog, 0.33))

                        # ----------------------------------------------------
                        # TAHAP 2: Sync ke mst_functloc (IDs without Parent)
                        # ----------------------------------------------------
                        status_text.text("Step 2/3: Registering records into mst_functloc...")
                        for i in range(0, total_records, batch_size):
                            batch_p1 = [dict(item, sup_functloc_id=None) for item in mapped_data[i:i + batch_size]]
                            auth_supabase.table('mst_functloc').upsert(batch_p1).execute()
                            
                            prog = 0.33 + (((i + batch_size) / total_records) * 0.33)
                            progress_bar.progress(min(prog, 0.66))

                        # ----------------------------------------------------
                        # TAHAP 3: Sync Parent-Child Hierarchy Relationships
                        # ----------------------------------------------------
                        status_text.text("Step 3/3: Linking parent-child hierarchy relations...")
                        for i in range(0, total_records, batch_size):
                            batch_p2 = mapped_data[i:i + batch_size]
                            auth_supabase.table('mst_functloc').upsert(batch_p2).execute()
                            
                            prog = 0.66 + (((i + batch_size) / total_records) * 0.34)
                            progress_bar.progress(min(prog, 1.0))

                        st.balloons()
                        st.success("🎉 All 3 synchronization steps completed successfully with 100% integrity!")

                    except Exception as e:
                        st.error(f"Synchronization failed: {str(e)}")
        else:
            st.warning("🔒 Synchronization Restricted: Sign in via sidebar using your Supabase account to execute database updates.")

    except Exception as e:
        st.error(f"Error reading dataset: {str(e)}")
