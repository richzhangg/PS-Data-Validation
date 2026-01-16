import streamlit as st
import pandas as pd
import pyodbc

from io import BytesIO
from datetime import datetime

# =============================
# Helpers
# =============================
def connect(server, database, username, password):
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
    )

def read_any(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            dtype=str,
            encoding="utf-8-sig",
            keep_default_na=False,
            na_filter=False,
        )
    return pd.read_excel(
        uploaded_file,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
    )

def stringify_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert every cell to a display-safe string while preserving DB formatting
    (e.g., Decimal('1.0000') -> '1.0000'). Also prevents pandas showing 1.0.
    """
    def _to_str(x):
        # pandas sometimes uses NaN floats for missing values
        if x is None:
            return ""
        try:
            # handle NaN
            if isinstance(x, float) and pd.isna(x):
                return ""
        except Exception:
            pass
        return str(x)
    return df.applymap(_to_str)

def norm_val(x):
    s = "" if x is None else str(x)
    s = s.replace("\u00A0", " ")
    s = s.strip().lower()
    s = " ".join(s.split())
    return s

def format_lines(lines, limit=50):
    if not lines:
        return ""
    if len(lines) <= limit:
        return ", ".join(str(x) for x in lines)
    return f"{', '.join(str(x) for x in lines[:limit])}, ... (+{len(lines)-limit} more)"

def make_excel_download(dfs: dict, filename_prefix: str = "differences"):
    """
    dfs: {"SheetName": dataframe, ...}
    returns (bytes, filename)
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            if df is None:
                df = pd.DataFrame()
            safe_name = str(sheet_name)[:31]  # Excel sheet name max length
            df.to_excel(writer, index=False, sheet_name=safe_name)

    output.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{ts}.xlsx"
    return output.getvalue(), filename

# ---------- Single-column ----------
def build_index_map_single(series: pd.Series):
    """
    Excel-style: header row = 1, first data row = 2
    Returns normalized list + value->list[excel_row_numbers]
    """
    norm_values = []
    index_map = {}
    for row_num, raw in enumerate(series.tolist(), start=2):
        v = norm_val(raw)
        norm_values.append(v)
        if v != "":
            index_map.setdefault(v, []).append(row_num)
    return norm_values, index_map

def compare_values_single(a_series, b_series, remove_dupes: bool):
    total_rows_a = len(a_series)
    total_rows_b = len(b_series)

    a_norm, a_idx = build_index_map_single(a_series)
    b_norm, b_idx = build_index_map_single(b_series)

    a_nonblank = [v for v in a_norm if v]
    b_nonblank = [v for v in b_norm if v]

    a_unique = set(a_nonblank)
    b_unique = set(b_nonblank)

    missing_details = []
    extra_details = []

    if remove_dupes:
        missing_vals = sorted(a_unique - b_unique)
        extra_vals = sorted(b_unique - a_unique)

        for v in missing_vals:
            missing_details.append({"value": v, "source_lines": format_lines(a_idx.get(v, []))})

        for v in extra_vals:
            extra_details.append({"value": v, "d365_lines": format_lines(b_idx.get(v, []))})

        missing_count = len(missing_vals)
        extra_count = len(extra_vals)

    else:
        a_counts = pd.Series(a_nonblank).value_counts()
        b_counts = pd.Series(b_nonblank).value_counts()
        all_vals = sorted(set(a_counts.index).union(b_counts.index))

        missing_count = 0
        extra_count = 0

        for v in all_vals:
            a_c = int(a_counts.get(v, 0))
            b_c = int(b_counts.get(v, 0))

            if a_c > b_c:
                missing_count += (a_c - b_c)
                missing_details.append({
                    "value": v,
                    "source_count": a_c,
                    "d365_count": b_c,
                    "source_lines": format_lines(a_idx.get(v, [])),
                    "d365_lines": format_lines(b_idx.get(v, [])),
                })

            if b_c > a_c:
                extra_count += (b_c - a_c)
                extra_details.append({
                    "value": v,
                    "source_count": a_c,
                    "d365_count": b_c,
                    "source_lines": format_lines(a_idx.get(v, [])),
                    "d365_lines": format_lines(b_idx.get(v, [])),
                })

    return {
        "total_rows_a": total_rows_a,
        "total_rows_b": total_rows_b,
        "unique_a": len(a_unique),
        "unique_b": len(b_unique),
        "missing_count": missing_count,
        "extra_count": extra_count,
        "missing_details": missing_details,
        "extra_details": extra_details,
    }

# ---------- Multi-column (unordered, composite key) ----------
def build_index_map_multi(df: pd.DataFrame, cols: list[str]):
    """
    Returns:
      keys_nb: list[tuple] normalized composite keys (non-blank only)
      idx_map: dict[key_tuple] -> list[excel_row_numbers] where key appears
    Blank key = all parts blank -> ignored.
    """
    keys_nb = []
    idx_map = {}

    for i in range(len(df)):
        excel_row = i + 2  # header=1, first data row=2
        key = tuple(norm_val(df.iloc[i][c]) for c in cols)

        if all(part == "" for part in key):
            continue

        keys_nb.append(key)
        idx_map.setdefault(key, []).append(excel_row)

    return keys_nb, idx_map

def _row_from_key(prefix: str, cols: list[str], key: tuple):
    return {f"{prefix}.{c}": key[i] for i, c in enumerate(cols)}

def compare_values_multi_unordered(src_df: pd.DataFrame, d365_df: pd.DataFrame,
                                  src_cols: list[str], d365_cols: list[str],
                                  remove_dupes: bool):
    """
    Missing in D365 (Source - D365): show SOURCE composite keys not found in D365
    Extra in D365 (D365 - Source): show D365 composite keys not found in Source
    Order does NOT matter.
    """
    src_keys, src_idx = build_index_map_multi(src_df, src_cols)
    d365_keys, d365_idx = build_index_map_multi(d365_df, d365_cols)

    src_unique = set(src_keys)
    d365_unique = set(d365_keys)

    missing_details = []
    extra_details = []

    if remove_dupes:
        missing_keys = sorted(src_unique - d365_unique)
        extra_keys = sorted(d365_unique - src_unique)

        for k in missing_keys:
            row = _row_from_key("source", src_cols, k)
            row["source_lines"] = format_lines(src_idx.get(k, []))
            missing_details.append(row)

        for k in extra_keys:
            row = _row_from_key("d365", d365_cols, k)
            row["d365_lines"] = format_lines(d365_idx.get(k, []))
            extra_details.append(row)

        missing_count = len(missing_keys)
        extra_count = len(extra_keys)

    else:
        src_counts = pd.Series(src_keys).value_counts()
        d365_counts = pd.Series(d365_keys).value_counts()
        all_keys = sorted(set(src_counts.index).union(set(d365_counts.index)))

        missing_count = 0
        extra_count = 0

        for k in all_keys:
            s_c = int(src_counts.get(k, 0))
            d_c = int(d365_counts.get(k, 0))

            if s_c > d_c:
                missing_count += (s_c - d_c)
                row = _row_from_key("source", src_cols, k)
                row.update({
                    "source_count": s_c,
                    "d365_count": d_c,
                    "source_lines": format_lines(src_idx.get(k, [])),
                    "d365_lines": format_lines(d365_idx.get(k, [])),
                })
                missing_details.append(row)

            if d_c > s_c:
                extra_count += (d_c - s_c)
                row = _row_from_key("d365", d365_cols, k)
                row.update({
                    "source_count": s_c,
                    "d365_count": d_c,
                    "source_lines": format_lines(src_idx.get(k, [])),
                    "d365_lines": format_lines(d365_idx.get(k, [])),
                })
                extra_details.append(row)

    return {
        "source_rows": len(src_df),
        "d365_rows": len(d365_df),
        "source_unique": len(src_unique),
        "d365_unique": len(d365_unique),
        "missing_count": missing_count,
        "extra_count": extra_count,
        "missing_details": missing_details,  # source-side rows
        "extra_details": extra_details,      # d365-side rows
    }

def reset_app():
    for k in ["src_df", "d365_df", "loaded"]:
        st.session_state.pop(k, None)
    st.rerun()

# =============================
# App UI
# =============================
st.set_page_config(page_title="Source vs D365 Validator", layout="wide")
st.title("Source File vs D365 SQL Validator")

with st.sidebar:
    st.header("D365 SQL Connection")
    server = st.text_input("Server")
    database = st.text_input("Database")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    st.divider()
    if st.button("Reset App"):
        reset_app()

page = st.radio(
    "Mode",
    ["Single Column Compare", "Multi-Column Compare (Unordered, Max 4)"],
    horizontal=True
)

st.header("1) Upload Source File")
src_file = st.file_uploader("Excel or CSV", type=["xlsx", "xls", "csv"])

st.header("2) Paste D365 SQL Query")
d365_query = st.text_area(
    "D365 SQL",
    height=160,
    placeholder="SELECT col1, col2, col3 FROM dbo.SomeD365Table;"
)

st.header("3) Comparison Options")
remove_dupes = st.checkbox(
    "Remove duplicates before comparing (unique-only)",
    value=True,
    help="ON = unique keys only. OFF = duplicates (counts) matter."
)

if page == "Multi-Column Compare (Unordered, Max 4)":
    num_cols = st.selectbox("How many columns to compare?", options=[2, 3, 4], index=0)
    st.info("Order does NOT matter. Matches by composite key across the dataset.")

if st.button("Load Data"):
    if not src_file:
        st.error("Please upload a source file.")
        st.stop()
    if not d365_query.strip():
        st.error("Please paste a D365 SQL query.")
        st.stop()
    if not (server and database and username and password):
        st.error("Please fill in all SQL connection fields.")
        st.stop()

    # Source: read as strings (keeps file display as the user sees it)
    try:
        st.session_state["src_df"] = read_any(src_file)
    except Exception as e:
        st.error(f"Failed to read source file: {e}")
        st.stop()

    # D365: IMPORTANT - preserve DB formatting (Decimal trailing zeros, etc.)
    try:
        conn = connect(server, database, username, password)
        raw_d365 = pd.read_sql(d365_query, conn, coerce_float=False)
        st.session_state["d365_df"] = stringify_df(raw_d365)
    except Exception as e:
        st.error(f"SQL error: {e}")
        st.stop()

    st.session_state["loaded"] = True
    st.success("Data loaded successfully.")

if st.session_state.get("loaded"):
    src_df = st.session_state["src_df"]
    d365_df = st.session_state["d365_df"]

    st.subheader("Preview")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Source")
        st.caption(f"Rows: {len(src_df)} | Columns: {len(src_df.columns)}")
        if len(src_df) <= 5000:
            st.dataframe(src_df, use_container_width=True)
        else:
            st.dataframe(src_df.head(50), use_container_width=True)
            st.info("Source > 5000 rows — showing first 50.")
    with c2:
        st.markdown("### D365")
        st.caption(f"Rows: {len(d365_df)} | Columns: {len(d365_df.columns)}")
        st.dataframe(d365_df.head(50), use_container_width=True)

    st.divider()

    # =============================
    # Single Column Compare
    # =============================
    if page == "Single Column Compare":
        st.subheader("4) Select Columns to Compare (Single Column)")
        col1, col2 = st.columns(2)
        with col1:
            src_col = st.selectbox("Source column", src_df.columns, key="single_src_col")
        with col2:
            d365_col = st.selectbox("D365 column", d365_df.columns, key="single_d365_col")

        results = compare_values_single(src_df[src_col], d365_df[d365_col], remove_dupes)

        # Download excel (ALL differences)
        missing_df = pd.DataFrame(results["missing_details"]) if results["missing_details"] else pd.DataFrame()
        extra_df = pd.DataFrame(results["extra_details"]) if results["extra_details"] else pd.DataFrame()
        summary_df = pd.DataFrame([{
            "source_rows": results["total_rows_a"],
            "d365_rows": results["total_rows_b"],
            "source_unique": results["unique_a"],
            "d365_unique": results["unique_b"],
            "missing_in_d365": results["missing_count"],
            "extra_in_d365": results["extra_count"],
            "remove_duplicates": remove_dupes,
            "source_column": src_col,
            "d365_column": d365_col,
        }])

        excel_bytes, excel_name = make_excel_download(
            {
                "Summary": summary_df,
                "Missing_in_D365": missing_df,
                "Extra_in_D365": extra_df,
            },
            filename_prefix="single_column_differences",
        )

        st.download_button(
            label="Download differences (Excel)",
            data=excel_bytes,
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Summary metrics
        st.subheader("Summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Source rows", results["total_rows_a"])
        m2.metric("D365 rows", results["total_rows_b"])
        m3.metric("Source unique", results["unique_a"])
        m4.metric("D365 unique", results["unique_b"])

        d1, d2 = st.columns(2)
        d1.metric("In Source but Missing in D365", results["missing_count"])
        d2.metric("In D365 but Missing in Source", results["extra_count"])

        st.divider()

        r1, r2 = st.columns(2)
        with r1:
            st.markdown("### Missing in D365 (Source − D365)")
            if results["missing_details"]:
                st.dataframe(missing_df, use_container_width=True)
                st.code("\n".join(row["value"] for row in results["missing_details"]), language=None)
            else:
                st.code("(none)", language=None)

        with r2:
            st.markdown("### Extra in D365 (D365 − Source)")
            if results["extra_details"]:
                st.dataframe(extra_df, use_container_width=True)
                st.code("\n".join(row["value"] for row in results["extra_details"]), language=None)
            else:
                st.code("(none)", language=None)

    # =============================
    # Multi-Column Compare (Unordered)
    # =============================
    else:
        st.subheader("4) Select Columns to Compare (Multi-Column, Unordered)")

        st.markdown("#### Source columns")
        src_cols = []
        for i in range(num_cols):
            src_cols.append(
                st.selectbox(
                    f"Source column {i+1}",
                    options=list(src_df.columns),
                    key=f"mc_src_{i}",
                )
            )

        st.markdown("#### D365 columns")
        d365_cols = []
        for i in range(num_cols):
            d365_cols.append(
                st.selectbox(
                    f"D365 column {i+1}",
                    options=list(d365_df.columns),
                    key=f"mc_d365_{i}",
                )
            )

        if st.button("Compare Selected Columns"):
            if len(set(src_cols)) != len(src_cols):
                st.error("Source columns contain duplicates. Please pick distinct Source columns.")
                st.stop()
            if len(set(d365_cols)) != len(d365_cols):
                st.error("D365 columns contain duplicates. Please pick distinct D365 columns.")
                st.stop()

            res = compare_values_multi_unordered(src_df, d365_df, src_cols, d365_cols, remove_dupes)

            # Build for download
            missing_df = pd.DataFrame(res["missing_details"]) if res["missing_details"] else pd.DataFrame()
            extra_df = pd.DataFrame(res["extra_details"]) if res["extra_details"] else pd.DataFrame()
            summary_df = pd.DataFrame([{
                "source_rows": res["source_rows"],
                "d365_rows": res["d365_rows"],
                "source_unique_keys": res["source_unique"],
                "d365_unique_keys": res["d365_unique"],
                "missing_in_d365": res["missing_count"],
                "extra_in_d365": res["extra_count"],
                "remove_duplicates": remove_dupes,
                "source_columns": ", ".join(src_cols),
                "d365_columns": ", ".join(d365_cols),
            }])

            excel_bytes, excel_name = make_excel_download(
                {
                    "Summary": summary_df,
                    "Missing_in_D365": missing_df,
                    "Extra_in_D365": extra_df,
                },
                filename_prefix="multi_column_differences",
            )

            st.download_button(
                label="Download differences (Excel)",
                data=excel_bytes,
                file_name=excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Summary metrics
            st.subheader("Summary (Multi-Column Unordered)")
            a, b, c, d = st.columns(4)
            a.metric("Source rows", res["source_rows"])
            b.metric("D365 rows", res["d365_rows"])
            c.metric("Source unique keys", res["source_unique"])
            d.metric("D365 unique keys", res["d365_unique"])

            e1, e2 = st.columns(2)
            e1.metric("In Source but Missing in D365", res["missing_count"])
            e2.metric("In D365 but Missing in Source", res["extra_count"])

            st.divider()

            r1, r2 = st.columns(2)

            with r1:
                st.markdown("### Missing in D365 (Source − D365)")
                if res["missing_details"]:
                    st.dataframe(missing_df, use_container_width=True)
                    # Copy ONLY selected SOURCE columns (tab-separated)
                    copy_lines = []
                    for row in res["missing_details"]:
                        vals = [row.get(f"source.{c}", "") for c in src_cols]
                        copy_lines.append("\t".join(vals))
                    st.code("\n".join(copy_lines), language=None)
                else:
                    st.code("(none)", language=None)

            with r2:
                st.markdown("### Extra in D365 (D365 − Source)")
                if res["extra_details"]:
                    st.dataframe(extra_df, use_container_width=True)
                    # Copy ONLY selected D365 columns (tab-separated)
                    copy_lines = []
                    for row in res["extra_details"]:
                        vals = [row.get(f"d365.{c}", "") for c in d365_cols]
                        copy_lines.append("\t".join(vals))
                    st.code("\n".join(copy_lines), language=None)
                else:
                    st.code("(none)", language=None)

else:
    st.info("Upload a file, paste SQL, and click **Load Data**.")
