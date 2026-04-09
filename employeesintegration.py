import os
from datetime import datetime

downloads_path = r"C:\Users\rzhang\Downloads"

csv_files = [
    f for f in os.listdir(downloads_path)
    if f.startswith("employee_") and f.endswith(".csv")
]

if not csv_files:
    print("No employee CSV files found.")
    exit()

csv_files.sort(reverse=True)

latest_filename = csv_files[0]
latest_file = os.path.join(downloads_path, latest_filename)

print("Processing file:", latest_file)

today = datetime.now().strftime("%m%d")

output_file = os.path.join(downloads_path, f"{today}_modified_{latest_filename}")

count_100 = 0
count_305 = 0

with open(latest_file, "r", encoding="utf-8-sig") as infile, \
     open(output_file, "w", encoding="utf-8", newline="") as outfile:

    for line in infile:
        stripped_line = line.rstrip("\r\n")

        if stripped_line.startswith('"100"'):
            stripped_line = stripped_line.replace('"en_US"', '"EN"')
            count_100 += 1

        if stripped_line.startswith('"305"'):
            stripped_line += "," * 37
            count_305 += 1

        outfile.write(stripped_line + "\n")

print("Done!")
print("Output saved to:", output_file)
print('Lines starting with "100" modified:', count_100)
print('Lines starting with "305" modified:', count_305)