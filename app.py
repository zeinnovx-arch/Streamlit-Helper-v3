from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Dashboard Filter Pelanggan PLN",
    page_icon="▦",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOMER_COLUMNS = ["ID_Pelanggan", "Nama", "Daerah", "Jam_Nyala"]
LOGIN_USERNAME = "PLNDKP@FILTER"
LOGIN_PASSWORD = "DKP.12345"
PLN_LOGO_PATH = "attached_assets/pln-logo.svg"
HOURS_FILTER_OPTIONS = ["Semua Data", "0–50 Jam", "50–80 Jam", "80–150 Jam"]


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
    """Apply the focused customer filters from the uploaded customer workflow."""
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


def show_empty_state() -> None:
    st.info(
        "Unggah file Excel atau CSV melalui bilah samping untuk mulai memfilter. "
        "File hanya digunakan selama sesi ini dan tidak disimpan."
    )
    st.subheader("Yang dapat dilakukan")
    columns = st.columns(3)
    columns[0].write("**Jelajahi**\n\nPilih sheet dan lihat data secara ringkas.")
    columns[1].write("**Filter**\n\nCari di semua kolom atau gunakan filter teks, angka, dan tanggal.")
    columns[2].write("**Ekspor**\n\nUnduh hasil filter sebagai CSV atau file Excel baru.")


def render_logo_header() -> None:
    header_left, header_right = st.columns([5, 1])
    with header_right:
        st.image(PLN_LOGO_PATH, width=120)


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
        if username == LOGIN_USERNAME and password == LOGIN_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Username atau password salah.")
    return False


render_logo_header()
if not show_login():
    st.stop()

st.title("Dashboard Filter Pelanggan PLN")
st.caption("Kelola, filter, dan unduh data pelanggan dengan lebih cepat.")

with st.sidebar:
    st.header("Akun")
    st.caption(f"Masuk sebagai: {LOGIN_USERNAME}")
    if st.button("Keluar", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
    st.divider()
    st.header("Muat data")
    uploaded_file = st.file_uploader(
        "Pilih file Excel atau CSV",
        type=["xlsx", "xls", "csv"],
        help="Format yang didukung: .xlsx, .xls, dan .csv",
    )

if uploaded_file is None:
    show_empty_state()
    st.stop()

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
    st.subheader("Aplikasi Filter Data Pelanggan Excel")
    st.caption(
        "Filter data pelanggan berdasarkan Daerah dan Jam Nyala, lalu unduh hasilnya."
    )
    st.success("File berhasil diunggah.")
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
    st.dataframe(
        display_dataframe,
        use_container_width=True,
        hide_index=True,
        height=520,
    )
    st.download_button(
        "Unduh hasil filter CSV",
        data=filtered_dataframe.to_csv(index=False).encode("utf-8"),
        file_name="hasil-filter-pelanggan.csv" if customer_mode else "filtered-data.csv",
        mime="text/csv",
        disabled=filtered_dataframe.empty,
    )
    st.download_button(
        "Unduh hasil filter Excel",
        data=excel_download(filtered_dataframe),
        file_name="hasil-filter-pelanggan.xlsx"
        if customer_mode
        else "filtered-data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=filtered_dataframe.empty,
    )
    if filtered_dataframe.empty:
        st.warning("Tidak ada baris yang cocok. Coba longgarkan filter Anda.")

with tab_summary:
    overview = pd.DataFrame(
        {
            "Kolom": dataframe.columns,
            "Tipe": [str(dataframe[column].dtype) for column in dataframe.columns],
            "Tidak kosong": [int(dataframe[column].notna().sum()) for column in dataframe.columns],
            "Nilai unik": [int(dataframe[column].nunique(dropna=True)) for column in dataframe.columns],
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