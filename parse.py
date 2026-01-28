import pandas as pd
import requests
import io
import json
from datetime import datetime


def parse_barcodes(value):
    if pd.isna(value):
        return set()
    s = str(value).strip()
    if not s:
        return set()
    if s.startswith('['):
        try:
            items = json.loads(s)
            return {str(x).strip().strip('"').strip("'") for x in items if str(x).strip()}
        except (json.JSONDecodeError, TypeError):
            pass
    parts = s.split(',')
    result = set()
    for p in parts:
        cleaned = p.strip().strip('"').strip("'").strip()
        if cleaned:
            result.add(cleaned)
    return result


def main():
    wb_url = 'https://docs.google.com/spreadsheets/d/1-2kDzJnRBqa5R6PrqTNc9bEVfX8caCtXlciUe35NP9U/export?format=csv'
    ms_url = 'https://docs.google.com/spreadsheets/d/1FHDfhdFRkSMJTztLoF_h7W5jyVSdSnmzPnjMrj8ym8M/export?format=csv'

    wb_df = pd.read_csv(io.StringIO(requests.get(wb_url).text), dtype=str)
    ms_df = pd.read_csv(io.StringIO(requests.get(ms_url).text), dtype=str)

    print(f"WB rows: {len(wb_df)}, MS rows: {len(ms_df)}")

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    results = []

    for _, wb_row in wb_df.iterrows():
        article = str(wb_row['wb_supplier_article']).strip() if pd.notna(wb_row['wb_supplier_article']) else ''
        wb_barcodes = parse_barcodes(wb_row['wb_barcodes'])

        matched_ms = ms_df[ms_df['ms_article'].astype(str).str.strip() == article]

        for _, ms_row in matched_ms.iterrows():
            ms_barcodes = parse_barcodes(ms_row['ms_barcodes'])

            if not wb_barcodes & ms_barcodes:
                for bc in sorted(wb_barcodes):
                    results.append({
                        'wb_nm_id': wb_row['wb_nm_id'],
                        'wb_supplier_article': article,
                        'wb_barcode': bc,
                        'ms_article': ms_row['ms_article'],
                        'ms_id': ms_row['ms_id'],
                        'created_at': now,
                    })

    result_df = pd.DataFrame(results, columns=['wb_nm_id', 'wb_supplier_article', 'wb_barcode', 'ms_article', 'ms_id', 'created_at'])

    output_path = r'D:\Projects\WB\result.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        wb_df.to_excel(writer, sheet_name='WB', index=False)
        ms_df.to_excel(writer, sheet_name='MS', index=False)
        result_df.to_excel(writer, sheet_name='Result', index=False)

    print(f"\nProblematic barcodes found: {len(result_df)}")
    print(f"Result saved to: {output_path}")
    if not result_df.empty:
        print(f"\nSample rows:")
        print(result_df.head(10).to_string(index=False))
    else:
        print("No problematic barcodes found - all matched.")


if __name__ == '__main__':
    main()
