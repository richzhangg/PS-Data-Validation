import pandas as pd
import re

#converts to str, removes case and spaces
def norm_col(c):
    # header-only normalization
    return str(c).strip().lower().replace(" ", "")

#reads the files
def read_any(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False, na_filter=False)
    return pd.read_excel(path, dtype=str, keep_default_na=False, na_filter=False)

#show invis spaces
def show(x):
    return repr("" if x is None else x)

#finds matching records
def match_key(x):
    x = "" if x is None else str(x)
    return re.sub(r"[^a-z0-9]", "", x.lower())

def compare_files(a_path, b_path, a_key, b_key, compare_cols):
    a_label = a_path.split("\\")[-1]
    b_label = b_path.split("\\")[-1]

    a = read_any(a_path)
    b = read_any(b_path)

    # Normalize HEADERS 
    a.columns = [norm_col(c) for c in a.columns]
    b.columns = [norm_col(c) for c in b.columns]

    a_key_n = norm_col(a_key)
    b_key_n = norm_col(b_key)
    cols_n = [norm_col(c) for c in compare_cols]

    # avoids irrel columns
    a = a[[a_key_n] + cols_n].copy()
    b = b[[b_key_n] + cols_n].copy()

    # Build internal match keys (alignment only)
    a["_match"] = a[a_key_n].map(match_key)
    b["_match"] = b[b_key_n].map(match_key)

    # Use match key as index for fast lookup 
    bmap = b.set_index("_match")
    dup_keys = bmap.index[bmap.index.duplicated()].unique()
    if len(dup_keys) > 0:
        print(f"FAIL: Duplicate keys found in {b_label}")
        for k in dup_keys:
            print(f"  Duplicate key: {show(k)}")

    pass_ct = 0
    fail_ct = 0
    error_rows = []

    for i, r in a.iterrows():
        mk = r["_match"]
        if mk == "" or mk not in bmap.index:
            fail_ct += 1
            continue

        diffs = []

        #VALIDATE KEY FIELD EXACTLY
        a_key_val = r[a_key_n]
        b_key_val = bmap.loc[mk, b_key_n]
        if a_key_val != b_key_val:
            diffs.append(
                f"{a_key_n} (key field):\n"
                f"  {a_label}: {show(a_key_val)}\n"
                f"  {b_label}: {show(b_key_val)}"
            )

        #VALIDATE COMPARE COLS EXACTLY
        for col in cols_n:
            aval = r[col]
            bval = bmap.loc[mk, col]
            if aval != bval:
                diffs.append(
                    f"{col}:\n"
                    f"  {a_label}: {show(aval)}\n"
                    f"  {b_label}: {show(bval)}"
                )

        if diffs:
            error_rows.append((a_key_val, diffs))
        else:
            pass_ct += 1

    #output
    if error_rows:
        print(f"PASS_WITH_ERRORS: {len(error_rows)}")
        for key_val, diffs in error_rows:
            print(f"\nRECORD (from {a_label} key): {show(key_val)}")
            for d in diffs:
                print(d)
    else:
        print(f"PASS: {pass_ct}")

    if fail_ct:
        print(f"FAIL: {fail_ct}")


#CHANGE
#a - source
#b - d365
compare_files(
    a_path=r"C:\Users\rzhang\OneDrive - Pacific Sunwear of California, Inc\validation 122625\HandlingType.csv",
    b_path=r"C:\Users\rzhang\OneDrive - Pacific Sunwear of California, Inc\validation 122625\d365 validation files\handlingtypes_d365.xlsx",
    a_key="Handling Type",
    b_key="HandlingTypeID",
    compare_cols=["Description"]
)
