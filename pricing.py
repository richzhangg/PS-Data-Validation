from pathlib import Path
import csv

# ---------------- CONFIG ----------------
FOLDER = r"C:\Users\rzhang\OneDrive - Pacific Sunwear of California, Inc\validation 011626\pricing validation"
FILE_PREFIX = "StorePrice_"
FILE_RANGE = range(17)
FILE_EXTENSION = ".csv"
DELIMITER = ","

OUTPUT_FILE = str(Path(FOLDER) / "mismatch_rows_only.csv")

D365_ITEM_RELATION = input("Enter D365 ITEMRELATION: ").strip()
D365_COLOR = input("Enter D365 COLOR: ").strip()
D365_AMOUNT = input("Enter D365 AMOUNT: ").strip()

CLASS_COLUMN_CANDIDATES = ["class", "Class", "CLASS"]
VENDOR_COLUMN_CANDIDATES = ["vendor", "Vendor", "VENDOR"]
STYLE_COLUMN_CANDIDATES = ["style", "Style", "STYLE"]
COLOR_COLUMN_CANDIDATES = ["color", "Color", "COLOR", "inventcolorid", "InventColorId", "INVENTCOLORID"]
RETAIL_COLUMN_CANDIDATES = ["retail", "Retail", "RETAIL", "price", "Price", "PRICE"]
# ----------------------------------------


def split_d365_item_relation(val):
    return val.split("-")


def find_column(header, candidates):
    lowered = {col.lower(): col for col in header}
    for c in candidates:
        if c in header:
            return header.index(c)
        if c.lower() in lowered:
            return header.index(lowered[c.lower()])
    return None


def pad(val, width):
    return str(val).strip().zfill(width)


def normalize_amount(val):
    val = str(val).strip().replace("$", "").replace(",", "")
    try:
        return f"{float(val):.2f}"
    except:
        return val


def run():
    c, v, s = split_d365_item_relation(D365_ITEM_RELATION)
    color = D365_COLOR
    amount = normalize_amount(D365_AMOUNT)

    cw, vw, sw, colw = len(c), len(v), len(s), len(color)

    header_written = False
    total_mismatches = 0

    with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as out:
        writer = None

        for i in FILE_RANGE:
            file_path = Path(FOLDER) / f"{FILE_PREFIX}{i}{FILE_EXTENSION}"

            if not file_path.exists():
                print(f"[WARNING] File {i} not found")
                continue

            print(f"\nSearching file {i}...")

            with open(file_path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
                reader = csv.reader(f, delimiter=DELIMITER)

                try:
                    header = next(reader)
                except:
                    continue

                class_i = find_column(header, CLASS_COLUMN_CANDIDATES)
                vendor_i = find_column(header, VENDOR_COLUMN_CANDIDATES)
                style_i = find_column(header, STYLE_COLUMN_CANDIDATES)
                color_i = find_column(header, COLOR_COLUMN_CANDIDATES)
                retail_i = find_column(header, RETAIL_COLUMN_CANDIDATES)

                if None in (class_i, vendor_i, style_i, color_i, retail_i):
                    print(f"[WARNING] Missing columns in file {i}")
                    continue

                if not header_written:
                    writer = csv.writer(out)
                    writer.writerow(header)
                    header_written = True

                for line_num, row in enumerate(reader, start=2):

                    # Progress print every 100k rows
                    if line_num % 100000 == 0:
                        print(f"File {i}: processed {line_num} rows...")

                    if max(class_i, vendor_i, style_i, color_i, retail_i) >= len(row):
                        continue

                    nc = pad(row[class_i], cw)
                    nv = pad(row[vendor_i], vw)
                    ns = pad(row[style_i], sw)
                    ncol = pad(row[color_i], colw)
                    nret = normalize_amount(row[retail_i])

                    if (
                        nc == c and
                        nv == v and
                        ns == s and
                        ncol == color
                    ):
                        if nret != amount:
                            writer.writerow(row)
                            total_mismatches += 1

            print(f"Finished file {i}")

    print("\nDone.")
    print(f"Total mismatches found: {total_mismatches}")
    print(f"Output file: {OUTPUT_FILE}")


if __name__ == "__main__":
    run()