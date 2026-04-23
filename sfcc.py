import re
from pathlib import Path
from openpyxl import Workbook

xml_file = r"C:\Users\rzhang\Downloads\SfccDropShipMasterProducts.xml"
output_file = str(Path(xml_file).with_suffix(".xlsx"))

text = Path(xml_file).read_text(encoding="utf-8")

product_pattern = re.compile(
    r'<product\s+[^>]*product-id="([^"]+)"[^>]*>(.*?)</product>',
    re.DOTALL
)

short_desc_pattern = re.compile(
    r'<short-description[^>]*>(.*?)</short-description>',
    re.DOTALL
)

products = product_pattern.findall(text)

print(f"Found {len(products):,} products")

wb = Workbook()
ws = wb.active
ws.title = "Products"

ws.append(["PMHCVS", "PMHSHODES"])

for i, (product_id, product_body) in enumerate(products, start=1):
    descriptions = short_desc_pattern.findall(product_body)

    # Keeps exact encoded content, including &lt; &gt; &#13;
    combined_description = "".join(descriptions)

    ws.append([product_id, combined_description])

    if i % 100 == 0 or i == len(products):
        print(f"Processed {i:,}/{len(products):,}")

print("Saving Excel file...")
wb.save(output_file)

print(f"Done. Saved to: {output_file}")