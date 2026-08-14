"""PLN Functional Location Data Importer.

Menjalankan aplikasi:
    streamlit run app.py

Dependensi utama:
    streamlit
    pandas
    openpyxl
    supabase

Konfigurasi `.streamlit/secrets.toml`:
    SUPABASE_URL = "https://<project-ref>.supabase.co"
    SUPABASE_KEY = "<publishable-key-atau-anon-key>"

Catatan keamanan:
    Jangan menggunakan service_role key pada aplikasi Streamlit ini. Import
    dijalankan memakai sesi pengguna Supabase Auth dan tetap mengikuti RLS.
"""

from __future__ import annotations

import io
import math
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Iterable

import pandas as pd
import streamlit as st
from supabase import Client, create_client


APP_TITLE = "PLN Functional Location Data Importer"
BATCH_SIZE = 250

# Baris group transmisi (KD_FUNGSI = X) tidak disimpan. Anak dari group ini
# diarahkan kembali ke functional location induknya melalui normalize_parent_id.
EXCLUDED_RAW_FUNCTION_CODES = {"X"}

REQUIRED_EXCEL_COLUMNS = {
    "ID_FUNCTLOC",
    "SUP_FUNCTLOC",
    "NM_LOKASI",
    "DESKRIPSI",
    "NLEVEL",
    "STATUS",
    "TEGANGAN",
    "KD_FUNGSI",
    "KD_WILAYAH",
    "KD_GROUPLOKASI",
    "BAYGROUP",
    "GI_FLC",
    "BC_FLC",
}

# Kolom mst_functloc yang memiliki FK ke tabel referensi selain ref_flc.
REFERENCE_FIELDS = {
    "function_code": ("ref_function", "function_code"),
    "status_code": ("ref_gi_status", "gi_status_code"),
    "voltage_code": ("ref_voltage", "voltage_code"),
    "region_code": ("ref_region", "region_code"),
    "grouplokasi_code": ("ref_grouplokasi", "grouplokasi_code"),
    "baygroup_code": ("ref_baygroup", "baygroup_code"),
}


@dataclass
class PreparedImport:
    review_df: pd.DataFrame
    ref_flc_rows: list[dict[str, Any]]
    functloc_rows: list[dict[str, Any]]
    audit_rows: list[dict[str, Any]]
    excluded_count: int


def clean_str(value: Any) -> str | None:
    """Membersihkan teks umum tanpa mengubah isi identifier."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "nat"}:
        return None
    return text


def clean_code(value: Any) -> str | None:
    """Membersihkan kode dan mengubah angka Excel 103.0 menjadi '103'."""
    text = clean_str(value)
    if text is None:
        return None

    normalized = text.replace(",", ".")
    try:
        number = float(normalized)
        if math.isfinite(number) and number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return text


def clean_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    try:
        number = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def clean_date(value: Any) -> str | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def normalize_function_code(
    code_value: Any,
    level_value: Any,
    description_value: Any = None,
) -> str | None:
    code = clean_code(code_value)
    if code is None:
        return None

    level = clean_float(level_value) or 0.0
    description = (clean_str(description_value) or "").upper()

    if code == "G" and (level == 4.0 or "SERANDANG" in description):
        return "SG"

    mapping = {
        ("V", 3.0): "V1",
        ("V", 4.0): "V2",
        ("X", 3.5): "X1",
        ("X", 4.0): "X2",
        ("O", 4.0): "O1",
        ("O", 2.0): "O2",
        ("Q", 3.5): "Q1",
        ("Q", 4.0): "Q2",
    }
    return mapping.get((code, level), code)


def determine_is_active(status_value: Any) -> bool:
    status = clean_code(status_value)
    return status in {"2", "4", "5", "6", "7", "8"}


def normalize_parent_id(raw_parent: Any) -> str | None:
    parent_id = clean_str(raw_parent)
    if parent_id is None:
        return None
    # Export PLN ditemukan memakai dua pola group: "-GR217" dan ".GR001".
    return re.sub(r"(?:-|\.)GR\d+$", "", parent_id, flags=re.IGNORECASE)


def read_master_excel(file_bytes: bytes) -> pd.DataFrame:
    """Mencari baris header secara otomatis lalu membaca data master."""
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object)

    header_row: int | None = None
    for index, row in raw.iterrows():
        values = {str(value).strip().upper() for value in row if not pd.isna(value)}
        if {"ID_FUNCTLOC", "KD_FUNGSI", "NLEVEL"}.issubset(values):
            header_row = int(index)
            break

    if header_row is None:
        raise ValueError(
            "Baris header tidak ditemukan. Excel harus memiliki kolom "
            "ID_FUNCTLOC, KD_FUNGSI, dan NLEVEL."
        )

    headers = [clean_str(value) or f"UNNAMED_{i}" for i, value in enumerate(raw.iloc[header_row])]
    headers = [header.strip().upper() for header in headers]

    duplicated_headers = pd.Series(headers)[pd.Series(headers).duplicated()].unique().tolist()
    if duplicated_headers:
        raise ValueError(f"Header Excel duplikat: {', '.join(duplicated_headers)}")

    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)

    missing = sorted(REQUIRED_EXCEL_COLUMNS - set(df.columns))
    if missing:
        raise ValueError("Kolom wajib tidak ditemukan: " + ", ".join(missing))
    return df


def _audit(
    audit_rows: list[dict[str, Any]],
    issue: str,
    flc_id: str | None,
    source_value: Any,
    action: str,
) -> None:
    audit_rows.append(
        {
            "issue": issue,
            "functloc_id": flc_id,
            "source_value": clean_str(source_value),
            "action": action,
        }
    )


def prepare_import(df: pd.DataFrame) -> PreparedImport:
    """Mengubah dataframe Excel menjadi payload ref_flc dan mst_functloc."""
    work = df.copy()
    work["RAW_FUNCTION_CODE"] = work["KD_FUNGSI"].map(clean_code)

    excluded_mask = work["RAW_FUNCTION_CODE"].isin(EXCLUDED_RAW_FUNCTION_CODES)
    excluded_count = int(excluded_mask.sum())
    work = work.loc[~excluded_mask].copy()

    work["FUNCTION_CODE_CLEAN"] = work.apply(
        lambda row: normalize_function_code(
            row.get("KD_FUNGSI"),
            row.get("NLEVEL"),
            row.get("DESKRIPSI"),
        ),
        axis=1,
    )
    work["ID_FUNCTLOC_CLEAN"] = work["ID_FUNCTLOC"].map(clean_str)

    empty_id_count = int(work["ID_FUNCTLOC_CLEAN"].isna().sum())
    if empty_id_count:
        raise ValueError(f"Terdapat {empty_id_count} baris tanpa ID_FUNCTLOC.")

    duplicates = work.loc[
        work["ID_FUNCTLOC_CLEAN"].duplicated(keep=False),
        "ID_FUNCTLOC_CLEAN",
    ].dropna().unique().tolist()
    if duplicates:
        sample = ", ".join(duplicates[:10])
        raise ValueError(f"ID_FUNCTLOC duplikat ditemukan: {sample}")

    missing_names = work["NM_LOKASI"].map(clean_str).isna()
    if missing_names.any():
        sample_ids = work.loc[missing_names, "ID_FUNCTLOC_CLEAN"].head(10).tolist()
        raise ValueError(
            "NM_LOKASI kosong pada functional location: " + ", ".join(sample_ids)
        )

    valid_ids = set(work["ID_FUNCTLOC_CLEAN"].tolist())
    function_by_id = dict(
        zip(work["ID_FUNCTLOC_CLEAN"], work["FUNCTION_CODE_CLEAN"], strict=True)
    )

    ref_flc_rows: list[dict[str, Any]] = []
    functloc_rows: list[dict[str, Any]] = []
    audit_rows: list[dict[str, Any]] = []
    timestamp = datetime.now(timezone.utc).isoformat()

    for _, row in work.iterrows():
        flc_id = clean_str(row.get("ID_FUNCTLOC"))
        assert flc_id is not None

        level = clean_float(row.get("NLEVEL"))
        function_code = clean_str(row.get("FUNCTION_CODE_CLEAN"))
        location_name = clean_str(row.get("NM_LOKASI"))
        status_code = clean_code(row.get("STATUS"))

        raw_parent = clean_str(row.get("SUP_FUNCTLOC"))
        normalized_parent = normalize_parent_id(raw_parent)
        parent_id = normalized_parent if normalized_parent in valid_ids else None
        if raw_parent and parent_id is None:
            _audit(
                audit_rows,
                "SUP_FUNCTLOC tidak ditemukan",
                flc_id,
                raw_parent,
                "sup_functloc_id diisi NULL",
            )

        raw_gi_flc = clean_str(row.get("GI_FLC"))
        if function_code == "G" and level == 3.0:
            gi_flc = flc_id
            if raw_gi_flc and raw_gi_flc != flc_id:
                _audit(
                    audit_rows,
                    "GI_FLC pada baris GI berbeda dari ID_FUNCTLOC",
                    flc_id,
                    raw_gi_flc,
                    "gi_flc menggunakan ID_FUNCTLOC GI",
                )
        elif raw_gi_flc in valid_ids:
            gi_flc = raw_gi_flc
        elif level == 4.0 and parent_id and function_by_id.get(parent_id) == "G":
            gi_flc = parent_id
            if raw_gi_flc is None:
                _audit(
                    audit_rows,
                    "GI_FLC kosong pada aset GI",
                    flc_id,
                    parent_id,
                    "gi_flc diisi dari SUP_FUNCTLOC",
                )
        else:
            gi_flc = None
            if raw_gi_flc:
                _audit(
                    audit_rows,
                    "GI_FLC tidak ditemukan",
                    flc_id,
                    raw_gi_flc,
                    "gi_flc diisi NULL",
                )

        raw_bc_flc = clean_str(row.get("BC_FLC"))
        bc_flc = raw_bc_flc if raw_bc_flc in valid_ids else None
        if raw_bc_flc and bc_flc is None:
            _audit(
                audit_rows,
                "BC_FLC bukan FLC valid pada file",
                flc_id,
                raw_bc_flc,
                "bc_flc diisi NULL",
            )

        ref_flc_rows.append(
            {
                "flc_id": flc_id,
                "name": location_name,
                "function_code": function_code,
                "is_active": determine_is_active(status_code),
                "updated_at": timestamp,
            }
        )

        functloc_rows.append(
            {
                "functloc_id": flc_id,
                "sup_functloc_id": parent_id,
                "location_name": location_name,
                "description": clean_str(row.get("DESKRIPSI")),
                "unit_code": clean_code(row.get("UNIT")),
                "nlevel": level,
                "status_code": status_code,
                "voltage_code": clean_code(row.get("TEGANGAN")),
                "function_code": function_code,
                "region_code": clean_code(row.get("KD_WILAYAH")),
                "operational_date": clean_date(row.get("TGL_OPRS")),
                "non_operational_date": clean_date(row.get("TGL_TDK_OPRS")),
                "workcenter": clean_str(row.get("WORKCENTER")),
                "plant_id": clean_code(row.get("ID_PLANT")),
                "grouplokasi_code": clean_code(row.get("KD_GROUPLOKASI")),
                "short_name": clean_str(row.get("NMSINGKT")),
                "address": clean_str(row.get("ALAMAT")),
                "city": clean_str(row.get("KOTA")),
                "postal_code": clean_code(row.get("KODEPOS")),
                "latitude": clean_float(row.get("LATITUDE")),
                "longitude": clean_float(row.get("LONGITUDE")),
                "altitude": clean_float(row.get("ALTITUDE")),
                "baygroup_code": clean_code(row.get("BAYGROUP")),
                "gi_flc": gi_flc,
                "slo_date": clean_date(row.get("TGL_SLO")),
                "slo_number": clean_str(row.get("NO_SLO")),
                "plant_section": clean_code(row.get("PLANT_SECTION")),
                "ownership": clean_str(row.get("MILIK")),
                "bc_flc": bc_flc,
                "updated_at": timestamp,
            }
        )

    review_columns = [
        "ID_FUNCTLOC",
        "SUP_FUNCTLOC",
        "NM_LOKASI",
        "NLEVEL",
        "STATUS",
        "KD_FUNGSI",
        "FUNCTION_CODE_CLEAN",
        "GI_FLC",
        "BC_FLC",
        "PLANT_SECTION",
    ]
    review_columns = [column for column in review_columns if column in work.columns]

    return PreparedImport(
        review_df=work[review_columns].copy(),
        ref_flc_rows=ref_flc_rows,
        functloc_rows=functloc_rows,
        audit_rows=audit_rows,
        excluded_count=excluded_count,
    )


def chunks(rows: list[dict[str, Any]], size: int) -> Iterable[list[dict[str, Any]]]:
    for start in range(0, len(rows), size):
        yield rows[start : start + size]


def create_authenticated_client(
    supabase_url: str,
    supabase_key: str,
    access_token: str,
    refresh_token: str,
) -> tuple[Client, Any]:
    client = create_client(supabase_url, supabase_key)
    auth_response = client.auth.set_session(access_token, refresh_token)
    return client, auth_response


def fetch_reference_values(client: Client, table: str, column: str) -> set[str]:
    response = client.table(table).select(column).execute()
    return {
        value
        for item in (response.data or [])
        if (value := clean_code(item.get(column))) is not None
    }


def validate_database_references(
    client: Client,
    functloc_rows: list[dict[str, Any]],
) -> list[dict[str, str]]:
    problems: list[dict[str, str]] = []

    for field, (table, column) in REFERENCE_FIELDS.items():
        allowed = fetch_reference_values(client, table, column)
        source_values = {
            value
            for row in functloc_rows
            if (value := clean_code(row.get(field))) is not None
        }
        for unknown in sorted(source_values - allowed):
            problems.append(
                {
                    "field": field,
                    "value": unknown,
                    "reference": f"{table}.{column}",
                }
            )
    return problems


def upsert_batch(
    client: Client,
    table: str,
    rows: list[dict[str, Any]],
    conflict_column: str,
) -> None:
    client.table(table).upsert(rows, on_conflict=conflict_column).execute()


def sync_master_data(
    client: Client,
    prepared: PreparedImport,
    progress_callback: Any = None,
) -> None:
    ref_batches = list(chunks(prepared.ref_flc_rows, BATCH_SIZE))
    base_batches = list(chunks(prepared.functloc_rows, BATCH_SIZE))
    relation_batches = list(chunks(prepared.functloc_rows, BATCH_SIZE))
    total_batches = len(ref_batches) + len(base_batches) + len(relation_batches)
    completed = 0

    def notify(phase: str) -> None:
        nonlocal completed
        completed += 1
        if progress_callback:
            progress_callback(completed, total_batches, phase)

    # Tahap 1: semua ID harus tersedia di ref_flc terlebih dahulu.
    for batch in ref_batches:
        upsert_batch(client, "ref_flc", batch, "flc_id")
        notify("Mengisi ref_flc")

    # Tahap 2: masukkan mst_functloc tanpa self-reference parent.
    for batch in base_batches:
        base_rows = [dict(row, sup_functloc_id=None) for row in batch]
        upsert_batch(client, "mst_functloc", base_rows, "functloc_id")
        notify("Mengisi mst_functloc")

    # Tahap 3: aktifkan kembali relasi parent setelah semua FLC tersedia.
    for batch in relation_batches:
        upsert_batch(client, "mst_functloc", batch, "functloc_id")
        notify("Menghubungkan hierarchy")


def get_configuration() -> tuple[str, str]:
    try:
        secrets_url = st.secrets.get("SUPABASE_URL", "")
        secrets_key = st.secrets.get("SUPABASE_KEY", "")
    except FileNotFoundError:
        secrets_url = ""
        secrets_key = ""

    url = os.getenv("SUPABASE_URL") or secrets_url
    key = os.getenv("SUPABASE_KEY") or secrets_key
    return str(url).strip(), str(key).strip()


def initialize_state() -> None:
    defaults = {
        "user": None,
        "session": None,
        "prepared_import": None,
        "uploaded_signature": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_sidebar(base_client: Client | None, configured: bool) -> None:
    with st.sidebar:
        st.header("⚙️ System Status")
        if configured:
            st.success("Konfigurasi Supabase tersedia")
        else:
            st.error("SUPABASE_URL atau SUPABASE_KEY belum tersedia")

        st.divider()
        st.header("🔐 User Authentication")

        if st.session_state["user"]:
            email = getattr(st.session_state["user"], "email", "-")
            st.write(f"Login: **{email}**")

            if st.button("Logout", use_container_width=True):
                try:
                    if base_client and st.session_state["session"]:
                        session = st.session_state["session"]
                        base_client.auth.set_session(
                            session.access_token,
                            session.refresh_token,
                        )
                        base_client.auth.sign_out()
                except Exception:
                    pass

                st.session_state["user"] = None
                st.session_state["session"] = None
                st.rerun()
        else:
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")

            if st.button("Sign In", use_container_width=True, disabled=not configured):
                if not base_client:
                    st.error("Supabase belum dikonfigurasi.")
                elif not email or not password:
                    st.warning("Email dan password wajib diisi.")
                else:
                    try:
                        response = base_client.auth.sign_in_with_password(
                            {"email": email.strip(), "password": password}
                        )
                        st.session_state["user"] = response.user
                        st.session_state["session"] = response.session
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Login gagal: {exc}")


def main() -> None:
    st.set_page_config(page_title="PLN Data Importer", page_icon="⚡", layout="wide")
    st.title("⚡ PLN Functional Location Data Importer")
    st.caption("Ruang lingkup: sinkronisasi ref_flc dan mst_functloc dari master aset PLN.")

    initialize_state()
    supabase_url, supabase_key = get_configuration()
    configured = bool(supabase_url and supabase_key)
    base_client = create_client(supabase_url, supabase_key) if configured else None
    render_sidebar(base_client, configured)

    uploaded_file = st.file_uploader("Upload Excel Master Data Aset", type=["xlsx"])
    if not uploaded_file:
        st.info("Upload file Excel hasil export master aset PLN untuk memulai audit.")
        return

    file_bytes = uploaded_file.getvalue()
    signature = f"{uploaded_file.name}:{len(file_bytes)}:{hash(file_bytes)}"

    if st.session_state["uploaded_signature"] != signature:
        try:
            dataframe = read_master_excel(file_bytes)
            prepared = prepare_import(dataframe)
        except Exception as exc:
            st.session_state["prepared_import"] = None
            st.error(f"File tidak dapat diproses: {exc}")
            return

        st.session_state["prepared_import"] = prepared
        st.session_state["uploaded_signature"] = signature

    prepared: PreparedImport = st.session_state["prepared_import"]

    st.divider()
    st.subheader("Ringkasan Preflight")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Data siap", f"{len(prepared.functloc_rows):,}")
    col2.metric("KD_FUNGSI X dikeluarkan", f"{prepared.excluded_count:,}")
    col3.metric("Temuan audit", f"{len(prepared.audit_rows):,}")
    col4.metric(
        "BC_FLC valid",
        f"{sum(row['bc_flc'] is not None for row in prepared.functloc_rows):,}",
    )

    review_tab, audit_tab = st.tabs(["Review Data", "Audit Relasi"])

    with review_tab:
        search = st.text_input(
            "Filter data",
            placeholder="Cari ID FLC, nama lokasi, GI_FLC, atau BC_FLC...",
        )
        review = prepared.review_df
        if search:
            mask = pd.Series(False, index=review.index)
            for column in ["ID_FUNCTLOC", "NM_LOKASI", "GI_FLC", "BC_FLC"]:
                if column in review.columns:
                    mask |= review[column].astype(str).str.contains(
                        search,
                        case=False,
                        na=False,
                        regex=False,
                    )
            review = review.loc[mask]

        st.dataframe(review.head(500), use_container_width=True, hide_index=True)
        st.caption(f"Ditemukan {len(review):,} baris; tabel menampilkan maksimal 500 baris.")

    with audit_tab:
        if prepared.audit_rows:
            audit_df = pd.DataFrame(prepared.audit_rows)
            summary = (
                audit_df.groupby(["issue", "action"], dropna=False)
                .size()
                .reset_index(name="jumlah")
                .sort_values("jumlah", ascending=False)
            )
            st.dataframe(summary, use_container_width=True, hide_index=True)
            with st.expander("Lihat detail temuan"):
                st.dataframe(audit_df, use_container_width=True, hide_index=True)
        else:
            st.success("Tidak ada relasi bermasalah pada file.")

    st.divider()
    st.subheader("Sinkronisasi Supabase")
    st.warning(
        "Proses menggunakan upsert dan tidak menghapus data lama yang tidak terdapat "
        "di Excel. Pastikan preflight sudah diperiksa sebelum menjalankan import."
    )

    confirmation = st.checkbox(
        "Saya sudah memeriksa ringkasan dan audit data.",
        value=False,
    )

    can_import = bool(configured and st.session_state["user"] and confirmation)
    if st.button(
        "Jalankan Import Data",
        type="primary",
        disabled=not can_import,
        use_container_width=True,
    ):
        session = st.session_state.get("session")
        if not session:
            st.error("Sesi login tidak tersedia. Silakan login kembali.")
            return

        try:
            auth_client, auth_response = create_authenticated_client(
                supabase_url,
                supabase_key,
                session.access_token,
                session.refresh_token,
            )
            if getattr(auth_response, "session", None):
                st.session_state["session"] = auth_response.session
                st.session_state["user"] = auth_response.user

            with st.spinner("Memvalidasi tabel referensi Supabase..."):
                reference_problems = validate_database_references(
                    auth_client,
                    prepared.functloc_rows,
                )

            if reference_problems:
                st.error(
                    "Import dibatalkan karena terdapat kode yang belum tersedia "
                    "pada tabel referensi Supabase."
                )
                st.dataframe(
                    pd.DataFrame(reference_problems),
                    use_container_width=True,
                    hide_index=True,
                )
                return

            progress = st.progress(0.0, text="Menyiapkan sinkronisasi...")

            def update_progress(completed: int, total: int, phase: str) -> None:
                progress.progress(
                    completed / total,
                    text=f"{phase}: batch {completed}/{total}",
                )

            sync_master_data(auth_client, prepared, update_progress)
            progress.progress(1.0, text="Sinkronisasi selesai")

            st.success(
                f"Import berhasil: {len(prepared.functloc_rows):,} functional location "
                "telah diproses."
            )
        except Exception as exc:
            st.error(f"Import dihentikan karena terjadi kesalahan: {exc}")
            st.info(
                "Sebagian batch mungkin sudah tersimpan. Perbaiki penyebab error lalu "
                "jalankan kembali; upsert aman untuk mengulang data dengan ID yang sama."
            )


if __name__ == "__main__":
    main()
