import streamlit as st
import pandas as pd
import pyodbc

# -----------------------------
# Helpers
# -----------------------------
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

def norm_val(x):
    s = "" if x is None else str(x)
    s = s.replace("\u00A0", " ")
    s = s.strip().lower()
    s = " ".join(s.split())
    return s

def compare_sets(a_series, b_series):
    total_rows_a = len(a_series)
    total_rows_b = len(b_series)

    a_norm = [norm_val(v) for v in a_series if norm_val(v)]
    b_norm = [norm_val(v) for v in b_series if norm_val(v)]

    a_set = set(a_norm)
    b_set = set(b_norm)

    missing = sorted(a_set - b_set)
    extra = sorted(b_set - a_set)

    return {
        "total_rows_a": total_rows_a,
        "total_rows_b": total_rows_b,
        "unique_a": len(a_set),
        "unique_b": len(b_set),
        "missing": missing,
        "extra": extra,
        "missing_count": len(missing),
        "extra_count": len(extra),
    }

def reset_app():
    for k in ["src_df", "d365_df", "loaded"]:
        st.session_state.pop(k, None)
    st.rerun()

# -----------------------------
# App UI
# -----------------------------
st.set_page_config(page_title="Source vs D365 Validator", layout="wide")
st.title("Source File vs D365 SQL Validator")

# Sidebar
with st.sidebar:
    st.header("D365 SQL Connection")
    server = st.text_input("Server")
    database = st.text_input("Database")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    st.divider()
    if st.button("Reset App"):
        reset_app()

# Main inputs
st.header("1) Upload Source File (Excel / CSV)")
src_file = st.file_uploader(
    "Upload source file",
    type=["xlsx", "xls", "csv"]
)

st.header("2) Paste D365 SQL Query")
d365_query = st.text_area(
    "D365 SQL",
    height=160,
    placeholder="SELECT factoryname FROM dbo.SomeD365Table;"
)

# Load data button
if st.button("Load Data"):
    if not src_file:
        st.error("Please upload a source file.")
    elif not d365_query.strip():
        st.error("Please paste a D365 SQL query.")
    elif not (server and database and username and password):
        st.error("Please fill in all SQL connection fields.")
    else:
        try:
            st.session_state["src_df"] = read_any(src_file)
        except Exception as e:
            st.error(f"Failed to read source file: {e}")
            st.stop()

        try:
            conn = connect(server, database, username, password)
            st.session_state["d365_df"] = pd.read_sql(d365_query, conn)
        except Exception as e:
            st.error(f"SQL error: {e}")
            st.stop()

        st.session_state["loaded"] = True
        st.success("Data loaded successfully!")

# -----------------------------
# Comparison UI
# -----------------------------
if st.session_state.get("loaded"):
    src_df = st.session_state["src_df"]
    d365_df = st.session_state["d365_df"]

    st.subheader("Preview (first 50 rows)")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Source")
        st.dataframe(src_df.head(50), use_container_width=True)
    with c2:
        st.markdown("### D365")
        st.dataframe(d365_df.head(50), use_container_width=True)

    st.subheader("3) Select Columns to Compare")
    col1, col2 = st.columns(2)
    with col1:
        src_col = st.selectbox("Source column", src_df.columns)
    with col2:
        d365_col = st.selectbox("D365 column", d365_df.columns)

    results = compare_sets(src_df[src_col], d365_df[d365_col])

    # -----------------------------
    # Summary metrics
    # -----------------------------
    st.subheader("Summary")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Source rows", results["total_rows_a"])
    m2.metric("D365 rows", results["total_rows_b"])
    m3.metric("Source unique", results["unique_a"])
    m4.metric("D365 unique", results["unique_b"])

    d1, d2, d3 = st.columns(3)
    d1.metric("Missing in D365", results["missing_count"])
    d2.metric("Extra in D365", results["extra_count"])

    if results["unique_a"] > 0:
        mismatch_pct = round(
            (results["missing_count"] + results["extra_count"])
            / results["unique_a"] * 100,
            2,
        )
        d3.metric("Mismatch % (vs Source)", f"{mismatch_pct}%")
    else:
        d3.metric("Mismatch % (vs Source)", "N/A")

    st.divider()

    # -----------------------------
    # COPY-ICON LISTS
    # -----------------------------
    r1, r2 = st.columns(2)

    with r1:
        st.markdown("### Missing in D365 (Source − D365)")
        st.write(f"{results['missing_count']} values")

        missing_text = (
            "\n".join(results["missing"])
            if results["missing"]
            else "(none)"
        )
        st.code(missing_text, language=None)

    with r2:
        st.markdown("### Extra in D365 (D365 − Source)")
        st.write(f"{results['extra_count']} values")

        extra_text = (
            "\n".join(results["extra"])
            if results["extra"]
            else "(none)"
        )
        st.code(extra_text, language=None)

else:
    st.info("Upload a file, paste a D365 SQL query, then click **Load Data**.")
