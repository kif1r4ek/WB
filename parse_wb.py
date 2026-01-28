import requests
import pandas as pd
import json
import time


API_TOKEN = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjUwOTA0djEiLCJ0eXAiOiJKV1QifQ.eyJhY2MiOjEsImVudCI6MSwiZXhwIjoxNzgwMDEyODQ4LCJpZCI6IjAxOWFjNTMwLTAxNDktNzJlYi1hNzlhLTI1MWI2YzI1ZDk3MSIsImlpZCI6MTIxMDIyMTksIm9pZCI6MjE3MDk3LCJzIjoxMDczNzQxODI2LCJzaWQiOiJhODEwYzI0NC0zNTZkLTQwYmUtYjQzMC00NWQ3NWUzZGY5ODgiLCJ0IjpmYWxzZSwidWlkIjoxMjEwMjIxOX0.yk8MmiYcBYgXVJGg25Vn-4SdcNo2j2w-n5lE-a0ZC5RE0N1ld6ucFyuow3lc6-jXUqRvX-UXrtpuIfYCwbDL-w"
HEADERS = {
    "Authorization": API_TOKEN,
}

def get_all_cards():
    url = "https://content-api.wildberries.ru/content/v2/get/cards/list"
    all_cards = []

    cursor = {
        "limit": 100,
        "nmID": None,
        "updatedAt": None
    }

    while True:
        payload = {
            "settings": {
                "cursor": cursor
            }
        }

        response = requests.post(url, headers=HEADERS, json=payload)

        if response.status_code == 429:
            print("⏳ Rate limit WB, ждём 2 секунды...")
            time.sleep(2)
            continue

        if response.status_code != 200:
            print("❌ Ошибка WB:", response.text)
            response.raise_for_status()

        data = response.json()

        cards = data.get("cards", [])
        if not cards:
            break

        all_cards.extend(cards)

        cursor = data.get("cursor")
        if not cursor:
            break

        time.sleep(0.7)

    return all_cards




def filter_hair_dye(cards):
    result = []

    for card in cards:
        subject = card.get("subjectName", "").lower()  # 'краски для волос'

        # ищем просто слова 'краск' и 'волос' как части слов
        if "краск" in subject and "волос" in subject:
            result.append(card)

    return result



def prepare_rows(cards):
    rows = []

    for card in cards:
        nm_id = card.get("nmID")             # ОБРАТИ ВНИМАНИЕ: в ответе nmID с большой I
        vendor_code = card.get("vendorCode")

        all_barcodes = []

        for size in card.get("sizes", []):
            # используем 'skus', если 'barcodes' нет
            barcodes = size.get("barcodes") or size.get("skus") or []
            all_barcodes.extend(barcodes)

        if not all_barcodes:
            continue

        rows.append({
            "wb_nm_id": nm_id,
            "wb_supplier_article": vendor_code,
            "wb_barcodes": ", ".join(all_barcodes)  # записываем как строку через запятую
        })

    return rows



def save_to_excel(rows, filename="wb_hair_dye_test.xlsx"):
    df = pd.DataFrame(rows)

    with pd.ExcelWriter(
        filename,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="WB",
            index=False
        )


def main():
    print("Получаем карточки WB...")
    cards = get_all_cards()
    print(f"Всего карточек: {len(cards)}")

    print("Фильтруем краску для волос...")
    hair_dye_cards = filter_hair_dye(cards)
    print(f"Краска для волос: {len(hair_dye_cards)}")

    print("Подготавливаем данные...")
    rows = prepare_rows(hair_dye_cards)

    print("Сохраняем в Excel...")
    save_to_excel(rows)

    print("Готово! Файл wb_hair_dye.xlsx создан")
    print(cards[:1])


if __name__ == "__main__":
    main()
