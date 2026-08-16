from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Excel Filter App",
    page_icon="▦",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOMER_COLUMNS = ["ID_Pelanggan", "Nama", "Daerah", "Jam_Nyala"]


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
        return {}, f"Could not read this file: {error}"


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
        return "Missing"
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value)


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
        active_filters.append(f"Search: “{global_search.strip()}”")

    for column in filter_columns:
        series = dataframe[column]
        safe_key = "".join(character if character.isalnum() else "_" for character in column)

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
                f"{column} range",
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
                f"{column} date range",
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
                    f"{column} values",
                    options=unique_values,
                    default=unique_values,
                    key=f"values_{safe_key}",
                )
                if len(selected_values) != len(unique_values):
                    mask &= series.astype("string").isin(selected_values) | series.isna()
                    active_filters.append(f"{column}: {len(selected_values)} values")
            else:
                contains = st.sidebar.text_input(
                    f"{column} contains",
                    key=f"contains_{safe_key}",
                    placeholder="Type to match values",
                )
                if contains.strip():
                    mask &= series.astype("string").str.contains(
                        contains.strip(), case=False, na=False, regex=False
                    ) | series.isna()
                    active_filters.append(f"{column} contains “{contains.strip()}”")

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
            "Pilih Filter Jam Nyala",
            [
                "Semua Data",
                "Di bawah 50 Jam (< 50)",
                "Antara 80–150 Jam",
                "Atur Rentang Sendiri (Custom)",
            ],
            key="customer_hours_filter",
        )

    if daerah_input.strip():
        filtered_dataframe = filtered_dataframe[
            filtered_dataframe["Daerah"]
            .astype("string")
            .str.contains(daerah_input.strip(), case=False, na=False, regex=False)
        ]
        active_filters.append(f"Daerah: {daerah_input.strip()}")

    hours = pd.to_numeric(dataframe["Jam_Nyala"], errors="coerce")
    valid_hours = hours.dropna()

    if valid_hours.empty:
        st.sidebar.warning("Kolom Jam_Nyala tidak memiliki angka yang bisa difilter.")
    elif pilihan_jam == "Di bawah 50 Jam (< 50)":
        filtered_dataframe = filtered_dataframe[
            pd.to_numeric(filtered_dataframe["Jam_Nyala"], errors="coerce") < 50
        ]
        active_filters.append("Jam Nyala: di bawah 50 jam")
    elif pilihan_jam == "Antara 80–150 Jam":
        filtered_hours = pd.to_numeric(
            filtered_dataframe["Jam_Nyala"], errors="coerce"
        )
        filtered_dataframe = filtered_dataframe[
            filtered_hours.between(80, 150, inclusive="both")
        ]
        active_filters.append("Jam Nyala: 80–150 jam")
    elif pilihan_jam == "Atur Rentang Sendiri (Custom)":
        minimum = 0.0
        maximum = 150.0
        selected_range = st.sidebar.slider(
            "Rentang Jam Nyala (0–150)",
            min_value=minimum,
            max_value=maximum,
            value=(minimum, maximum),
            step=1.0,
            key="customer_hours_range",
        )
        filtered_hours = pd.to_numeric(
            filtered_dataframe["Jam_Nyala"], errors="coerce"
        )
        filtered_dataframe = filtered_dataframe[
            filtered_hours.between(*selected_range, inclusive="both")
        ]
        if selected_range != (minimum, maximum):
            active_filters.append(
                f"Jam Nyala: {format_value(selected_range[0])}–"
                f"{format_value(selected_range[1])} jam"
            )

    return filtered_dataframe, active_filters


def show_empty_state() -> None:
    st.info(
        "Upload an Excel workbook or CSV file in the sidebar to start filtering. "
        "Your file stays in this session and is not saved."
    )
    st.subheader("What you can do here")
    columns = st.columns(3)
    columns[0].write("**Explore**\n\nSwitch between workbook sheets and inspect the data at a glance.")
    columns[1].write("**Filter**\n\nSearch every column or apply targeted text, number, and date filters.")
    columns[2].write("**Export**\n\nDownload the filtered rows as CSV or a new Excel workbook.")


st.title("Excel Filter App")
st.caption("Turn a spreadsheet into a focused, searchable view in a few clicks.")

with st.sidebar:
    st.header("Load your data")
    uploaded_file = st.file_uploader(
        "Choose an Excel workbook or CSV file",
        type=["xlsx", "xls", "csv"],
        help="Supported formats: .xlsx, .xls, and .csv",
    )

if uploaded_file is None:
    show_empty_state()
    st.stop()

sheets, load_error = load_workbook(uploaded_file)
if load_error:
    st.error(load_error)
    st.stop()

if not sheets:
    st.warning("This file does not contain any readable sheets.")
    st.stop()

with st.sidebar:
    st.divider()
    sheet_name = st.selectbox("Select a sheet", list(sheets.keys()))

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
        st.header("Filter rows")
        global_search = st.text_input(
            "Search all columns",
            placeholder="e.g. customer, Jakarta, pending",
            help="Matches text anywhere in a row.",
        )
        filter_columns = st.multiselect(
            "Choose columns to filter",
            options=dataframe.columns.tolist(),
            help="Add as many column filters as you need.",
        )

    filtered_dataframe, active_filters = apply_filters(
        dataframe, filter_columns, global_search
    )
    display_dataframe = filtered_dataframe

total_rows = len(dataframe)
filtered_rows = len(filtered_dataframe)
match_rate = (filtered_rows / total_rows * 100) if total_rows else 0

metric_columns = st.columns(4)
metric_columns[0].metric("Rows shown", f"{filtered_rows:,}")
metric_columns[1].metric("Rows in sheet", f"{total_rows:,}")
metric_columns[2].metric("Columns", f"{len(dataframe.columns):,}")
metric_columns[3].metric("Match rate", f"{match_rate:.1f}%")

if active_filters:
    st.caption("Active filters: " + "  ·  ".join(active_filters))
else:
    st.caption(f"Showing all rows from “{sheet_name}”.")

tab_data, tab_summary = st.tabs(["Filtered data", "Column overview"])

with tab_data:
    st.dataframe(
        display_dataframe,
        use_container_width=True,
        hide_index=True,
        height=520,
    )
    st.download_button(
        "Download hasil filter CSV" if customer_mode else "Download filtered CSV",
        data=filtered_dataframe.to_csv(index=False).encode("utf-8"),
        file_name="hasil-filter-pelanggan.csv" if customer_mode else "filtered-data.csv",
        mime="text/csv",
        disabled=filtered_dataframe.empty,
    )
    st.download_button(
        "Download hasil filter Excel"
        if customer_mode
        else "Download filtered Excel",
        data=excel_download(filtered_dataframe),
        file_name="hasil-filter-pelanggan.xlsx"
        if customer_mode
        else "filtered-data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=filtered_dataframe.empty,
    )
    if filtered_dataframe.empty:
        st.warning("No rows match the current filters. Try widening your selection.")

with tab_summary:
    overview = pd.DataFrame(
        {
            "Column": dataframe.columns,
            "Type": [str(dataframe[column].dtype) for column in dataframe.columns],
            "Non-empty": [int(dataframe[column].notna().sum()) for column in dataframe.columns],
            "Unique values": [int(dataframe[column].nunique(dropna=True)) for column in dataframe.columns],
        }
    )
    st.dataframe(overview, use_container_width=True, hide_index=True)

    numeric_columns = dataframe.select_dtypes(include="number").columns.tolist()
    if numeric_columns:
        st.subheader("Numeric summary")
        st.dataframe(
            dataframe[numeric_columns].describe().transpose(),
            use_container_width=True,
        )