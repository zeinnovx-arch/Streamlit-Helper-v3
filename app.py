from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Dashboard Filter Pelanggan PLN",
    page_icon="▦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- KONSTANTA & SESSION STATE ---
CUSTOMER_COLUMNS = ["ID_Pelanggan", "Nama", "Daerah", "Jam_Nyala"]
LOGIN_USERNAME_DEFAULT = "PLNDKP@FILTER"
LOGIN_PASSWORD = "DKP.12345"
PLN_LOGO_PATH = "attached_assets/pln-logo.svg"
HOURS_FILTER_OPTIONS = ["Semua Data", "0–50 Jam", "50–80 Jam", "80–150 Jam"]
ADMIN_EMAIL = "zeinnovx@gmail.com"
ADMIN_EMAIL = "fuadmochamad2@gmail.cpm"

# Folder & file untuk penyimpanan persisten (riwayat upload & riwayat pengecekan)
DATA_DIR = "data"
UPLOAD_CACHE_META_PATH = os.path.join(DATA_DIR, "upload_meta.json")
CHECKED_LOG_PATH = os.path.join(DATA_DIR, "checked_log.json")

INDO_MONTHS = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

# Inisialisasi username & password pada session state agar bisa diubah secara dinamis
if "current_username" not in st.session_state:
    st.session_state.current_username = LOGIN_USERNAME_DEFAULT
if "current_password" not in st.session_state:
    st.session_state.current_password = LOGIN_PASSWORD


# --- FUNGSI UTILITY & HELPER (UMUM) ---
def load_workbook(uploaded_file: Any) -> tuple[dict[str, pd.DataFrame], str | None]:
    """Read an uploaded CSV or Excel workbook into named dataframes."""
    try:
        file_bytes = uploaded_file.getvalue()
        file_name = uploaded_file.name.lower()

        if file_name.endswith(".csv"):
            return {"CSV data": pd.read_csv(BytesIO(file_bytes))}, None

        engine = "xlrd" if file_name.endswith(".xls") else "openpyxl"
        workbook = pd.ExcelFile(BytesIO(file_bytes), engine=engine)
        sheets = {
            sheet_name: pd.read_excel(workbook, sheet_name=sheet_name)
            for sheet_name in workbook.sheet_names
        }
        return sheets, None
    except Exception as error:
        return {}, f"Gagal membaca file: {error}"


def normalise_dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.normalize()


def is_date_like(series: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if series.dtype != "object":
        return False
    non_empty = series.dropna()
    if non_empty.empty:
        return False
    converted = normalise_dates(non_empty)
    return converted.notna().mean() >= 0.8


def excel_download(dataframe: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Filtered data")
    return output.getvalue()


def format_value(value: Any) -> str:
    if pd.isna(value):
        return "Kosong"
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value)


def is_hours_column(column: Any) -> bool:
    normalised = "".join(character for character in str(column) if character.isalnum())
    return normalised.casefold() == "jamnyala"


def format_indonesian_date(date_obj) -> str:
    return f"{date_obj.day} {INDO_MONTHS[date_obj.month]} {date_obj.year}"


# --- FUNGSI PENYIMPANAN PERSISTEN: RIWAYAT FILE UPLOAD ---
class CachedUploadedFile:
    """Wrapper agar file yang dimuat ulang dari disk kompatibel dengan load_workbook()."""

    def __init__(self, path: str, name: str) -> None:
        self._path = path
        self.name = name

    def getvalue(self) -> bytes:
        with open(self._path, "rb") as file:
            return file.read()


def _cache_file_path_for(filename: str) -> str:
    ext = os.path.splitext(filename)[1] or ".dat"
    return os.path.join(DATA_DIR, f"cached_upload{ext}")


def save_uploaded_file_to_cache(uploaded_file: Any) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    # Bersihkan file cache lama (kalau ekstensinya beda dari sebelumnya)
    for existing_name in os.listdir(DATA_DIR):
        if existing_name.startswith("cached_upload"):
            os.remove(os.path.join(DATA_DIR, existing_name))

    cache_path = _cache_file_path_for(uploaded_file.name)
    with open(cache_path, "wb") as file:
        file.write(uploaded_file.getvalue())

    meta = {
        "filename": uploaded_file.name,
        "uploaded_at": datetime.now().isoformat(),
        "cache_path": cache_path,
    }
    with open(UPLOAD_CACHE_META_PATH, "w", encoding="utf-8") as file:
        json.dump(meta, file, ensure_ascii=False, indent=2)


def load_upload_cache_meta() -> dict | None:
    if not os.path.exists(UPLOAD_CACHE_META_PATH):
        return None
    try:
        with open(UPLOAD_CACHE_META_PATH, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return None


def load_cached_uploaded_file() -> Any | None:
    meta = load_upload_cache_meta()
    if not meta or not os.path.exists(meta.get("cache_path", "")):
        return None
    return CachedUploadedFile(meta["cache_path"], meta["filename"])


def clear_upload_history() -> None:
    """Hapus file tersimpan & reset riwayat pengecekan (dipanggil manual oleh pengguna)."""
    if os.path.exists(UPLOAD_CACHE_META_PATH):
        os.remove(UPLOAD_CACHE_META_PATH)
    if os.path.isdir(DATA_DIR):
        for existing_name in os.listdir(DATA_DIR):
            if existing_name.startswith("cached_upload"):
                os.remove(os.path.join(DATA_DIR, existing_name))
    save_checked_log({})


# --- FUNGSI PENYIMPANAN PERSISTEN: RIWAYAT PENGECEKAN PER BARIS ---
def load_checked_log() -> dict[str, str]:
    if not os.path.exists(CHECKED_LOG_PATH):
        return {}
    try:
        with open(CHECKED_LOG_PATH, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {}


def save_checked_log(log: dict[str, str]) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CHECKED_LOG_PATH, "w", encoding="utf-8") as file:
        json.dump(log, file, ensure_ascii=False, indent=2)


def compute_row_key(row: pd.Series, customer_mode: bool) -> str:
    """Kunci unik per baris. Pakai ID_Pelanggan kalau ada, kalau tidak pakai hash isi baris."""
    if customer_mode and "ID_Pelanggan" in row.index:
        return f"id:{row['ID_Pelanggan']}"
    raw = "|".join(str(row[column]) for column in row.index)
    return "hash:" + hashlib.md5(raw.encode("utf-8")).hexdigest()


def build_checkable_table(display_df: pd.DataFrame, customer_mode: bool) -> pd.DataFrame:
    """Tambahkan kolom 'Cek' & 'Status Pengecekan' berdasarkan riwayat yang tersimpan."""
    checked_log = load_checked_log()
    row_keys = [compute_row_key(row, customer_mode) for _, row in display_df.iterrows()]

    table = display_df.copy()
    table.index = row_keys
    table.insert(0, "Cek", [key in checked_log for key in row_keys])
    table["Status Pengecekan"] = [
        f"Sudah dicek pada tanggal {format_indonesian_date(datetime.fromisoformat(checked_log[key]).date())}"
        if key in checked_log
        else "Belum dicek"
        for key in row_keys
    ]
    return table


def sync_checked_log_from_editor(edited_table: pd.DataFrame) -> bool:
    """Bandingkan hasil editan kolom 'Cek' dengan riwayat tersimpan, lalu simpan bila berubah."""
    current_log = load_checked_log()
    today_iso = datetime.now().date().isoformat()
    changed = False

    for row_key, is_checked in zip(edited_table.index, edited_table["Cek"]):
        was_checked = row_key in current_log
        if is_checked and not was_checked:
            current_log[row_key] = today_iso
            changed = True
        elif not is_checked and was_checked:
            del current_log[row_key]
            changed = True

    if changed:
        save_checked_log(current_log)
    return changed


# --- FUNGSI FILTER & LOGIKA ---
def get_hours_filter_mask(
    series: pd.Series, widget_key: str
) -> tuple[pd.Series, str | None]:
    selected_range = st.sidebar.selectbox(
        "Rentang JAMNYALA",
        HOURS_FILTER_OPTIONS,
        key=widget_key,
    )
    numeric_series = pd.to_numeric(series, errors="coerce")
    if numeric_series.dropna().empty:
        st.sidebar.warning("Kolom JAMNYALA tidak memiliki angka yang bisa difilter.")
        return pd.Series(True, index=series.index), None

    if selected_range == "0–50 Jam":
        return numeric_series.ge(0) & numeric_series.lt(50), selected_range
    if selected_range == "50–80 Jam":
        return numeric_series.ge(50) & numeric_series.lt(80), selected_range
    if selected_range == "80–150 Jam":
        return numeric_series.ge(80) & numeric_series.le(150), selected_range

    return pd.Series(True, index=series.index), None


def apply_filters(
    dataframe: pd.DataFrame,
    filter_columns: list[str],
    global_search: str,
) -> tuple[pd.DataFrame, list[str]]:
    mask = pd.Series(True, index=dataframe.index)
    active_filters: list[str] = []

    if global_search.strip():
        search_term = global_search.strip().casefold()
        searchable = dataframe.astype("string").fillna("").apply(
            lambda column: column.str.casefold().str.contains(
                search_term, regex=False, na=False
            )
        )
        mask &= searchable.any(axis=1)
        active_filters.append(f"Pencarian: “{global_search.strip()}”")

    for column in filter_columns:
        series = dataframe[column]
        safe_key = "".join(character if character.isalnum() else "_" for character in column)

        if is_hours_column(column):
            hours_mask, selected_range = get_hours_filter_mask(
                series, f"hours_{safe_key}"
            )
            mask &= hours_mask
            if selected_range:
                active_filters.append(f"{column}: {selected_range}")
            continue

        if pd.api.types.is_numeric_dtype(series):
            numeric_series = pd.to_numeric(series, errors="coerce")
            valid_values = numeric_series.dropna()
            if valid_values.empty:
                continue
            minimum = float(valid_values.min())
            maximum = float(valid_values.max())
            if minimum == maximum:
                continue
            selected_range = st.sidebar.slider(
                f"Rentang {column}",
                min_value=minimum,
                max_value=maximum,
                value=(minimum, maximum),
                key=f"numeric_{safe_key}",
            )
            mask &= numeric_series.between(*selected_range) | numeric_series.isna()
            if selected_range != (minimum, maximum):
                active_filters.append(
                    f"{column}: {format_value(selected_range[0])}–{format_value(selected_range[1])}"
                )
        elif is_date_like(series):
            date_series = normalise_dates(series)
            valid_dates = date_series.dropna()
            if valid_dates.empty:
                continue
            minimum_date = valid_dates.min().date()
            maximum_date = valid_dates.max().date()
            selected_dates = st.sidebar.date_input(
                f"Rentang tanggal {column}",
                value=(minimum_date, maximum_date),
                min_value=minimum_date,
                max_value=maximum_date,
                key=f"date_{safe_key}",
            )
            if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
                start_date, end_date = selected_dates
                mask &= date_series.between(
                    pd.Timestamp(start_date), pd.Timestamp(end_date)
                ) | date_series.isna()
                if (start_date, end_date) != (minimum_date, maximum_date):
                    active_filters.append(f"{column}: {start_date}–{end_date}")
        else:
            values = series.dropna().astype(str)
            unique_values = sorted(values.unique().tolist())
            if len(unique_values) <= 150:
                selected_values = st.sidebar.multiselect(
                    f"Nilai {column}",
                    options=unique_values,
                    default=unique_values,
                    key=f"values_{safe_key}",
                )
                if len(selected_values) != len(unique_values):
                    mask &= series.astype("string").isin(selected_values) | series.isna()
                    active_filters.append(f"{column}: {len(selected_values)} nilai")
            else:
                contains = st.sidebar.text_input(
                    f"{column} memuat teks",
                    key=f"contains_{safe_key}",
                    placeholder="Ketik teks yang dicari",
                )
                if contains.strip():
                    mask &= series.astype("string").str.contains(
                        contains.strip(), case=False, na=False, regex=False
                    ) | series.isna()
                    active_filters.append(f"{column} memuat “{contains.strip()}”")

    return dataframe.loc[mask].copy(), active_filters


def apply_customer_filters(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str]]:
    filtered_dataframe = dataframe.copy()
    active_filters: list[str] = []

    with st.sidebar:
        st.divider()
        st.header("Panel filter pelanggan")
        daerah_input = st.text_input("Cari Daerah", placeholder="Contoh: Jakarta")
        pilihan_jam = st.selectbox(
            "Pilih Rentang Jam Nyala",
            HOURS_FILTER_OPTIONS,
            key="customer_hours_filter",
        )

    if daerah_input.strip():
        filtered_dataframe = filtered_dataframe[
            filtered_dataframe["Daerah"]
            .astype("string")
            .str.contains(daerah_input.strip(), case=False, na=False, regex=False)
        ]
        active_filters.append(f"Daerah: {daerah_input.strip()}")

    hours_mask = pd.Series(True, index=dataframe.index)
    if pilihan_jam == "0–50 Jam":
        hours_mask = pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").ge(0) & (
            pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").lt(50)
        )
    elif pilihan_jam == "50–80 Jam":
        hours_mask = pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").ge(50) & (
            pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").lt(80)
        )
    elif pilihan_jam == "80–150 Jam":
        hours_mask = pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").ge(80) & (
            pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce").le(150)
        )

    filtered_dataframe = filtered_dataframe.loc[
        hours_mask.reindex(filtered_dataframe.index, fill_value=False)
    ]

    if pilihan_jam != "Semua Data":
        active_filters.append(f"Jam Nyala: {pilihan_jam}")

    return filtered_dataframe, active_filters


# --- TAMPILAN HIASAN VISUAL / DEKORASI ---
def render_logo_header() -> None:
    header_left, header_right = st.columns([5, 1])
    with header_right:
        st.image(PLN_LOGO_PATH, width=120)


def render_opening_decoration() -> None:
    st.markdown(
        """
        <style>
        .pln-opening {
            position: relative;
            overflow: hidden;
            margin: 0.75rem 0 2rem;
            border-radius: 18px;
            background: #071a3b;
            box-shadow: 0 16px 32px rgba(7, 26, 59, 0.18);
            color: #ffffff;
        }
        .pln-opening__stripe {
            height: 8px;
            background: linear-gradient(
                90deg,
                #f6c700 0%,
                #f6c700 34%,
                #e52b35 34%,
                #e52b35 67%,
                #1254a4 67%,
                #1254a4 100%
            );
        }
        .pln-opening__body {
            position: relative;
            z-index: 1;
            padding: 2rem 2.25rem 2.15rem;
        }
        .pln-opening__eyebrow {
            color: #f6c700;
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.16em;
            text-transform: uppercase;
        }
        .pln-opening__headline {
            max-width: 33rem;
            margin-top: 0.55rem;
            color: #ffffff;
            font-size: clamp(1.55rem, 3vw, 2.45rem);
            font-weight: 750;
            line-height: 1.08;
        }
        .pln-opening__copy {
            max-width: 35rem;
            margin-top: 0.8rem;
            color: #d6e2f5;
            font-size: 1rem;
            line-height: 1.55;
        }
        .pln-opening__dots {
            display: flex;
            gap: 0.5rem;
            margin-top: 1.25rem;
        }
        .pln-opening__dot {
            display: block;
            width: 0.7rem;
            height: 0.7rem;
            border-radius: 999px;
        }
        .pln-opening__dot--yellow { background: #f6c700; }
        .pln-opening__dot--red { background: #e52b35; }
        .pln-opening__dot--blue { background: #2d73c9; }
        .pln-opening__orb {
            position: absolute;
            right: 8%;
            bottom: -5rem;
            width: 16rem;
            height: 16rem;
            border: 1.5rem solid rgba(246, 199, 0, 0.12);
            border-radius: 50%;
        }
        .pln-opening__beam {
            position: absolute;
            right: 15%;
            top: 2.4rem;
            width: 0.35rem;
            height: 9rem;
            transform: rotate(33deg);
            background: linear-gradient(#f6c700, #e52b35);
            opacity: 0.8;
        }
        </style>
        <div class="pln-opening" aria-label="Pembuka Dashboard Filter Pelanggan PLN">
            <div class="pln-opening__stripe"></div>
            <div class="pln-opening__body">
                <div class="pln-opening__eyebrow">Portal Data Pelanggan</div>
                <div class="pln-opening__headline">Satu dashboard untuk data yang lebih terang dan terarah.</div>
                <div class="pln-opening__copy">
                    Kelola, cari, dan filter data pelanggan PLN dengan cepat sebelum masuk ke dashboard.
                </div>
                <div class="pln-opening__dots" aria-hidden="true">
                    <span class="pln-opening__dot pln-opening__dot--yellow"></span>
                    <span class="pln-opening__dot pln-opening__dot--red"></span>
                    <span class="pln-opening__dot pln-opening__dot--blue"></span>
                </div>
            </div>
            <span class="pln-opening__orb" aria-hidden="true"></span>
            <span class="pln-opening__beam" aria-hidden="true"></span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_decorative_widgets() -> None:
    st.markdown(
        """
        <div style="background: linear-gradient(135deg, #071a3b 0%, #1254a4 50%, #f6c700 100%); padding: 16px 22px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h3 style="margin:0; color: #ffffff;">⚡ Selamat Datang di Portal Data PLN</h3>
            <p style="margin:4px 0 0 0; opacity: 0.95; font-size: 0.95rem;">Sistem Manajemen, Filtrasi, & Pelaporan Data Pelanggan Real-Time</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.toast("💡 Tip: Centang kolom 'Cek' untuk menandai data yang sudah diverifikasi!", icon="⚡")


# --- FITUR TOMBOL PRINT / CETAK TERHUBUNG DATA EXCEL ---
def render_print_button(df_to_print: pd.DataFrame) -> None:
    """Mencetak data Excel yang sudah difilter secara bersih dan responsif."""
    html_table = df_to_print.to_html(index=False, classes="print-table")

    js_print = f"""
    <style>
        .print-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }}
        .print-table th {{
            background-color: #1254a4;
            color: white;
            padding: 8px;
            border: 1px solid #ddd;
            text-align: left;
        }}
        .print-table td {{
            padding: 6px 8px;
            border: 1px solid #ddd;
        }}
        .print-table tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}

        @media print {{
            body * {{
                visibility: hidden;
            }}
            #print-area, #print-area * {{
                visibility: visible;
            }}
            #print-area {{
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
            }}
            .no-print {{
                display: none !important;
            }}
        }}
    </style>

    <div class="no-print">
        <button onclick="window.print()" style="
            background-color: #1254a4;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 8px;
            font-weight: bold;
            cursor: pointer;
            margin-top: 5px;
            margin-bottom: 15px;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        ">
            🖨️ Cetak / Print Tabel Excel Terfilter
        </button>
    </div>

    <div id="print-area" style="display: none;">
        <h2 style="text-align: center; color: #071a3b;">Laporan Data Pelanggan PLN (Hasil Filter)</h2>
        <hr>
        {html_table}
    </div>

    <script>
        window.onbeforeprint = function() {{
            document.getElementById('print-area').style.display = 'block';
        }};
        window.onafterprint = function() {{
            document.getElementById('print-area').style.display = 'none';
        }};
    </script>
    """
    st.components.v1.html(js_print, height=65)


def show_empty_state() -> None:
    st.info(
        "Unggah file Excel atau CSV melalui bilah samping untuk mulai memfilter. "
        "Setelah diunggah, file ini akan tersimpan otomatis di dashboard sehingga "
        "tidak hilang saat halaman dibuka kembali."
    )
    st.subheader("Yang dapat dilakukan")
    columns = st.columns(3)
    columns[0].write("**Jelajahi**\n\nPilih sheet dan lihat data secara ringkas.")
    columns[1].write(
        "**Filter**\n\nCari di semua kolom atau gunakan filter teks, angka, dan tanggal."
    )
    columns[2].write("**Cek & Ekspor**\n\nTandai data yang sudah diverifikasi, lalu unduh hasil filter.")


# --- HALAMAN LOGIN DENGAN FITUR LUPA / RESET USERNAME ---
def show_login() -> bool:
    if st.session_state.get("authenticated", False):
        return True

    st.title("Masuk ke Dashboard")
    st.caption("Silakan masukkan akun Anda untuk mengakses Dashboard Filter Pelanggan PLN.")

    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Username", placeholder="Masukkan username")
        password = st.text_input(
            "Password", type="password", placeholder="Masukkan password"
        )
        submitted = st.form_submit_button("Masuk", use_container_width=True)

        if submitted:
            if (
                username == st.session_state.current_username
                and password == st.session_state.current_password
            ):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Username atau password salah.")

    # --- FITUR BUAT / UBAH USERNAME & PASSWORD (AKSES KHUSUS ZEINNOVX@GMAIL.COM) ---
    with st.expander("🔑 Lupa / Ubah Username atau Password Admin?"):
        st.write(
            "Khusus admin **zeinnovx@gmail.com**, Anda dapat mendaftarkan atau merubah "
            "username baru, password baru, atau keduanya sekaligus di sini."
        )
        with st.form("reset_credentials_form"):
            admin_email_input = st.text_input(
                "Masukkan Email Verifikasi Admin", placeholder="zeinnovx@gmail.com"
            )
            new_username_input = st.text_input(
                "Username Baru (kosongkan jika tidak ingin mengubah)",
                placeholder="Ketik username baru",
            )
            new_password_input = st.text_input(
                "Password Baru (kosongkan jika tidak ingin mengubah)",
                type="password",
                placeholder="Ketik password baru",
            )
            confirm_password_input = st.text_input(
                "Konfirmasi Password Baru",
                type="password",
                placeholder="Ulangi password baru",
            )
            reset_submitted = st.form_submit_button(
                "Simpan & Perbarui Kredensial", use_container_width=True
            )

            if reset_submitted:
                if admin_email_input.strip().lower() != ADMIN_EMAIL:
                    st.error(
                        "❌ Verifikasi Gagal! Hanya email zeinnovx@gmail.com yang dapat "
                        "merubah username/password."
                    )
                elif not new_username_input.strip() and not new_password_input:
                    st.error("⚠️ Isi minimal salah satu: username baru atau password baru.")
                elif new_password_input and new_password_input != confirm_password_input:
                    st.error("⚠️ Konfirmasi password baru tidak sama dengan password baru.")
                else:
                    updates = []
                    if new_username_input.strip():
                        st.session_state.current_username = new_username_input.strip()
                        updates.append(f"Username baru: **{st.session_state.current_username}**")
                    if new_password_input:
                        st.session_state.current_password = new_password_input
                        updates.append("Password baru berhasil disimpan.")
                    st.success("✅ Kredensial berhasil diperbarui. " + " ".join(updates) + " Silakan masuk di atas.")

    return False


# --- LOGIKA UTAMA APLIKASI ---
render_logo_header()

if not show_login():
    render_opening_decoration()
    st.stop()

# Menampilkan hiasan dekorasi interaktif
render_decorative_widgets()

st.title("Dashboard Filter Pelanggan PLN")
st.caption("Kelola, filter, cek, cetak, dan unduh data pelanggan dengan lebih cepat.")

with st.sidebar:
    st.header("Akun")
    st.caption(f"Masuk sebagai: **{st.session_state.current_username}**")
    if st.button("Keluar", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

with st.sidebar:
    st.divider()
    st.header("Muat data")
    uploaded_file = st.file_uploader(
        "Pilih file Excel atau CSV",
        type=["xlsx", "xls", "csv"],
        help="Format yang didukung: .xlsx, .xls, dan .csv",
    )

# --- LOGIKA RIWAYAT UPLOAD (FILE TERSIMPAN DI DASHBOARD) ---
cached_meta = load_upload_cache_meta()
loaded_from_cache = False

if uploaded_file is not None:
    upload_signature = (uploaded_file.name, len(uploaded_file.getvalue()))
    previous_signature = st.session_state.get("last_upload_signature")

    if upload_signature != previous_signature:
        # Ini benar-benar unggahan baru pada sesi ini
        is_replacement = cached_meta is not None
        save_uploaded_file_to_cache(uploaded_file)
        if is_replacement:
            # Ganti file = riwayat pengecekan lama ikut hilang
            save_checked_log({})
            st.session_state["show_replaced_notice"] = True
        st.session_state["last_upload_signature"] = upload_signature
        cached_meta = load_upload_cache_meta()
elif cached_meta is not None:
    # Tidak ada unggahan baru pada sesi ini -> muat file yang tersimpan sebelumnya
    cached_file = load_cached_uploaded_file()
    if cached_file is not None:
        uploaded_file = cached_file
        loaded_from_cache = True
        st.session_state.setdefault(
            "last_upload_signature", (cached_meta["filename"], None)
        )

with st.sidebar:
    if cached_meta is not None:
        uploaded_at = datetime.fromisoformat(cached_meta["uploaded_at"])
        st.caption(
            f"📁 File tersimpan: **{cached_meta['filename']}**\n\n"
            f"Diunggah: {format_indonesian_date(uploaded_at.date())}"
        )
        if st.button("🗑️ Hapus Riwayat File Tersimpan", use_container_width=True):
            clear_upload_history()
            st.session_state.pop("last_upload_signature", None)
            st.rerun()

if st.session_state.pop("show_replaced_notice", False):
    st.info(
        "File sebelumnya telah digantikan dengan file baru. "
        "Riwayat pengecekan lama otomatis direset."
    )

if uploaded_file is None:
    show_empty_state()
    st.stop()

if loaded_from_cache:
    st.caption("ℹ️ Data dimuat otomatis dari file yang tersimpan sebelumnya di dashboard.")

sheets, load_error = load_workbook(uploaded_file)
if load_error:
    st.error(load_error)
    st.stop()

if not sheets:
    st.warning("File ini tidak memiliki sheet yang dapat dibaca.")
    st.stop()

with st.sidebar:
    st.divider()
    sheet_name = st.selectbox("Pilih sheet", list(sheets.keys()))

dataframe = sheets[sheet_name].copy()
dataframe.columns = [
    str(column) if str(column).strip() else f"Unnamed column {index + 1}"
    for index, column in enumerate(dataframe.columns)
]

customer_mode = set(CUSTOMER_COLUMNS).issubset(dataframe.columns)

if customer_mode:
    st.sidebar.subheader("Aplikasi Filter Data Pelanggan Excel")
    st.sidebar.caption(
        "Filter data pelanggan berdasarkan Daerah dan Jam Nyala, lalu unduh hasilnya."
    )
    st.sidebar.success("File berhasil diunggah.")
    filtered_dataframe, active_filters = apply_customer_filters(dataframe)
    display_dataframe = filtered_dataframe[CUSTOMER_COLUMNS]
else:
    with st.sidebar:
        st.divider()
        st.header("Filter baris")
        global_search = st.text_input(
            "Cari di semua kolom",
            placeholder="Contoh: pelanggan, Jakarta, pending",
            help="Mencocokkan teks di seluruh kolom.",
        )
        filter_columns = st.multiselect(
            "Pilih kolom untuk difilter",
            options=dataframe.columns.tolist(),
            help="Tambahkan satu atau beberapa kolom filter.",
        )

    filtered_dataframe, active_filters = apply_filters(
        dataframe, filter_columns, global_search
    )
    display_dataframe = filtered_dataframe

total_rows = len(dataframe)
filtered_rows = len(filtered_dataframe)
match_rate = (filtered_rows / total_rows * 100) if total_rows else 0

metric_columns = st.columns(4)
metric_columns[0].metric("Baris ditampilkan", f"{filtered_rows:,}")
metric_columns[1].metric("Total baris", f"{total_rows:,}")
metric_columns[2].metric("Kolom", f"{len(dataframe.columns):,}")
metric_columns[3].metric("Persentase cocok", f"{match_rate:.1f}%")

if active_filters:
    st.caption("Filter aktif: " + "  ·  ".join(active_filters))
else:
    st.caption(f"Menampilkan semua baris dari “{sheet_name}”.")

tab_data, tab_summary = st.tabs(["Data hasil filter", "Ringkasan kolom"])

with tab_data:
    if display_dataframe.empty:
        st.dataframe(display_dataframe, use_container_width=True, hide_index=True, height=200)
    else:
        st.caption(
            "✅ Centang kolom **Cek** pada baris yang sudah Anda verifikasi. "
            "Tanggal pengecekan akan otomatis tercatat dan tersimpan."
        )
        checkable_table = build_checkable_table(display_dataframe, customer_mode)
        non_editable_columns = [
            column for column in checkable_table.columns if column != "Cek"
        ]
        edited_table = st.data_editor(
            checkable_table,
            use_container_width=True,
            hide_index=True,
            height=520,
            key="data_editor_checked",
            column_config={
                "Cek": st.column_config.CheckboxColumn(
                    "Cek", help="Tandai baris ini sebagai sudah dicek/diverifikasi."
                ),
                "Status Pengecekan": st.column_config.TextColumn(
                    "Status Pengecekan", disabled=True
                ),
            },
            disabled=non_editable_columns,
        )

        if sync_checked_log_from_editor(edited_table):
            st.rerun()

    # Menampilkan Tombol Print Cetak yang Terhubung Langsung ke Excel Terfilter
    render_print_button(display_dataframe)

    col_csv, col_excel = st.columns(2)
    with col_csv:
        st.download_button(
            "Unduh hasil filter CSV",
            data=filtered_dataframe.to_csv(index=False).encode("utf-8"),
            file_name="hasil-filter-pelanggan.csv" if customer_mode else "filtered-data.csv",
            mime="text/csv",
            disabled=filtered_dataframe.empty,
            use_container_width=True,
        )
    with col_excel:
        st.download_button(
            "Unduh hasil filter Excel",
            data=excel_download(filtered_dataframe),
            file_name="hasil-filter-pelanggan.xlsx" if customer_mode else "filtered-data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=filtered_dataframe.empty,
            use_container_width=True,
        )

    if filtered_dataframe.empty:
        st.warning("Tidak ada baris yang cocok. Coba longgarkan filter Anda.")

with tab_summary:
    overview = pd.DataFrame(
        {
            "Kolom": dataframe.columns,
            "Tipe": [str(dataframe[column].dtype) for column in dataframe.columns],
            "Tidak kosong": [
                int(dataframe[column].notna().sum()) for column in dataframe.columns
            ],
            "Nilai unik": [
                int(dataframe[column].nunique(dropna=True))
                for column in dataframe.columns
            ],
        }
    )
    st.dataframe(overview, use_container_width=True, hide_index=True)

    numeric_columns = dataframe.select_dtypes(include="number").columns.tolist()
    if numeric_columns:
        st.subheader("Ringkasan numerik")
        st.dataframe(
            dataframe[numeric_columns].describe().transpose(),
            use_container_width=True,
        )
