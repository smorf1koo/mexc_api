#!/usr/bin/env python3
"""
MEXC Spot Scraper
Получает список всех спот-пар с текущими ценами и сохраняет в один xlsx.
"""

import os
import requests
import pandas as pd

MEXC_SPOT_API = "https://api.mexc.com/api/v3"

STATUS_MAP = {
    "1": "TRADING",
    "2": "HALT",
    "3": "BREAK",
    "ENABLED": "TRADING",
}


def get_exchange_info():
    r = requests.get(f"{MEXC_SPOT_API}/exchangeInfo", timeout=30)
    r.raise_for_status()
    return r.json().get("symbols", [])


def get_all_prices():
    r = requests.get(f"{MEXC_SPOT_API}/ticker/price", timeout=30)
    r.raise_for_status()
    return {t["symbol"]: float(t["price"]) for t in r.json()}


def scrape_spot(output_dir: str = "mexc", filename: str = "spot.xlsx"):
    print("🚀 MEXC Spot Scraper")
    print("=" * 60)

    symbols = get_exchange_info()
    print(f"✓ Получено {len(symbols)} пар")

    prices = get_all_prices()
    print(f"✓ Получено {len(prices)} цен")

    rows = []
    for s in symbols:
        symbol = s.get("symbol", "N/A")
        status_raw = str(s.get("status", "N/A"))
        rows.append({
            "Символ": symbol,
            "Базовый актив": s.get("baseAsset", ""),
            "Котируемый актив": s.get("quoteAsset", ""),
            "Полное название": s.get("fullName", ""),
            "Текущая цена": prices.get(symbol, 0.0),
            "Статус": STATUS_MAP.get(status_raw, status_raw),
        })

    df = pd.DataFrame(rows).sort_values("Символ").reset_index(drop=True)

    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, filename)

    column_widths = {"A": 18, "B": 15, "C": 18, "D": 30, "E": 18, "F": 12}
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Spot", index=False)
        ws = writer.sheets["Spot"]
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

    print("=" * 60)
    print(f"✓ Сохранено {len(df)} пар в {out_path}")
    print("=" * 60)


if __name__ == "__main__":
    scrape_spot()
