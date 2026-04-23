import xml.etree.ElementTree as ET
from openpyxl import Workbook
from pathlib import Path
import time

xml_file = r"C:\Users\rzhang\Downloads\SfccDropShipMasterProducts.xml"

output_file = str(Path(xml_file).with_suffix(".xlsx"))


def get_namespace(tag):
    if tag.startswith("{"):
        return tag.split("}")[0] + "}"
    return ""


def convert_xml_to_excel(xml_path, excel_path):
    start = time.time()

    print("Reading XML file...")
    tree = ET.parse(xml_path)
    root = tree.getroot()

    ns = get_namespace(root.tag)

    products = root.findall(f".//{ns}product")
    total = len(products)

    print(f"Found {total:,} products")
    print("Starting conversion...\n")

    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    # Headers only once
    ws.append(["PMHCVS", "PMHSHODES"])

    for i, product in enumerate(products, start=1):

        product_id = product.get("product-id", "").strip()

        descriptions = []

        for sd in product.findall(f".//{ns}short-description"):
            if sd.text:
                descriptions.append(sd.text.strip())

        short_desc_combined = " ".join(descriptions)

        ws.append([
            product_id,
            short_desc_combined
        ])

        # Progress every 100 products
        if i % 100 == 0 or i == total:
            pct = (i / total) * 100
            print(f"Processed {i:,}/{total:,} ({pct:.1f}%)")

    print("\nSaving Excel file...")
    wb.save(excel_path)

    elapsed = time.time() - start

    print("\nDONE")
    print(f"Saved to: {excel_path}")
    print(f"Elapsed time: {elapsed:.2f} seconds")


if __name__ == "__main__":
    convert_xml_to_excel(xml_file, output_file)