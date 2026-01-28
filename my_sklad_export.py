import requests
import json
from openpyxl import Workbook

URL = "https://api.moysklad.ru/api/remap/1.2"
AUTH = ("testapi@topmastershop772", "testapi321")

def main():
    r = requests.get(f"{URL}/entity/productfolder",
                     params={"filter": "name=Тест"}, auth=AUTH)
    r.raise_for_status()

    subfolder = None
    for row in r.json().get("rows", []):
        if row.get("pathName", "") == "Дом и красота":
            subfolder = row
            break

    if not subfolder:
        print("ERROR: Subfolder 'Тест' under 'Дом и красота' not found")
        return

    subfolder_href = subfolder["meta"]["href"]
    print(f"Found folder: {subfolder['pathName']}/{subfolder['name']} (id={subfolder['id']})")

    products = []
    offset = 0
    limit = 1000
    while True:
        r = requests.get(f"{URL}/entity/assortment",
                         params={"filter": f"productFolder={subfolder_href};type=product",
                                 "limit": limit, "offset": offset},
                         auth=AUTH)
        r.raise_for_status()
        data = r.json()
        batch = data.get("rows", [])
        products.extend(batch)
        total = data["meta"]["size"]
        print(f"  Fetched {len(products)}/{total} products")
        if len(products) >= total:
            break
        offset += limit

    print(f"Total products: {len(products)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "MS"
    ws.append(["ms_id", "ms_article", "ms_name", "ms_barcodes"])

    for p in products:
        barcodes_raw = p.get("barcodes", [])
        all_codes = [str(v) for bc_dict in barcodes_raw for v in bc_dict.values()]
        barcodes_json = json.dumps(all_codes, ensure_ascii=False)

        ws.append([
            p.get("id", ""),
            p.get("article", ""),
            p.get("name", ""),
            barcodes_json
        ])

    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = '@'

    out = r"D:\Projects\WB\moysklad_export.xlsx"
    wb.save(out)
    print(f"Saved to {out}")

    print("\nExported data:")
    for i, p in enumerate(products):
        barcodes_raw = p.get("barcodes", [])
        codes = [str(v) for bc in barcodes_raw for v in bc.values()]
        print(f"  [{i+1}] article={p.get('article','')}, name={p.get('name','')}, barcodes={codes}")

if __name__ == "__main__":
    main()