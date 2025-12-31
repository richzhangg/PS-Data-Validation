import pandas as pd

def norm_col(c):
    return str(c).strip().lower().replace(" ", "")

def read_any(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig",
                           keep_default_na=False, na_filter=False)
    return pd.read_excel(path, dtype=str, keep_default_na=False, na_filter=False)

def show(x):
    return repr("" if x is None else x)

def compare_files(a_path, b_path, a_col, b_col):
    a_label = a_path.split("\\")[-1]
    b_label = b_path.split("\\")[-1]

    a = read_any(a_path)
    b = read_any(b_path)

    # Normalize headers
    a.columns = [norm_col(c) for c in a.columns]
    b.columns = [norm_col(c) for c in b.columns]
    a_col_n = norm_col(a_col)
    b_col_n = norm_col(b_col)

    if a_col_n not in a.columns:
        raise KeyError(f"{a_label} missing column {show(a_col)}. Available: {list(a.columns)}")
    if b_col_n not in b.columns:
        raise KeyError(f"{b_label} missing column {show(b_col)}. Available: {list(b.columns)}")

    # Extract values and remove duplicates
    a_set = set(a[a_col_n].astype(str))
    b_set = set(b[b_col_n].astype(str))

    missing_in_b = sorted(a_set - b_set)
    extra_in_b = sorted(b_set - a_set)

    if not missing_in_b and not extra_in_b:
        print("PASS!")
        print(f"source= {len(a_set)}, D365= {len(b_set)}")
        return

    print("PASS WITH ERRORS")
    print(f"A column: {a_label}.{a_col}")
    print(f"B column: {b_label}.{b_col}")
    print(f"Unique values source= {len(a_set)}, d365= {len(b_set)}")

    if missing_in_b:
        print(f"\nMissing in d365:")
        for v in missing_in_b:
            print(f"  {show(v)}")

    if extra_in_b:
        print(f"\nMissing in source:")
        for v in extra_in_b:
            print(f"  {show(v)}")
            
#CHANGE
#a - source
#b - d365
compare_files(
    a_path=r"C:\Users\rzhang\OneDrive - Pacific Sunwear of California, Inc\validation 122625\wholesale\Customers_Wholesale_122625.xlsx",
    b_path=r"C:\Users\rzhang\OneDrive - Pacific Sunwear of California, Inc\validation 122625\d365 validation files\customerswholesale_d365.xlsx",
    a_col="customeraccount",
    b_col="accountnum"
)
