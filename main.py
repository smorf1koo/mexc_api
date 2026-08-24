#!/usr/bin/env python3
"""
MEXC Futures Scraper
Получает список фьючей с максимальным плечом, датой листинга и текущей ценой
Сохраняет результаты в Excel файл
"""

import os
import re
import requests
import pandas as pd
from datetime import datetime
from typing import List, Dict
import json
import time

# Конфигурация
MEXC_API_BASE = "https://contract.mexc.com/api/v1"
MEXC_SPOT_API = "https://api.mexc.com/api/v3"

# Твои API ключи (замени на свои!)
ACCESS_KEY = "your_access_key_here"
SECRET_KEY = "your_secret_key_here"

class MEXCFuturesScraper:
    def __init__(self, access_key: str = None, secret_key: str = None):
        self.access_key = access_key
        self.secret_key = secret_key
        self.session = requests.Session()
        
    def get_futures_list(self) -> List[Dict]:
        """Получает список всех активных фьючей"""
        try:
            url = f"{MEXC_API_BASE}/contract/detail"
            
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            
            if data.get('success') and data.get('data'):
                futures = data['data']
                print(f"✓ Получено {len(futures)} фьючей")
                return futures
            else:
                print(f"✗ Ошибка API: {data.get('msg', 'Unknown error')}")
                return []
                
        except requests.exceptions.RequestException as e:
            print(f"✗ Ошибка подключения: {e}")
            return []
    
    def get_current_price(self, symbol: str) -> float:
        """Получает текущую цену для фьючей"""
        try:
            url = f"{MEXC_API_BASE}/contract/ticker"
            params = {"symbol": symbol}
            
            response = self.session.get(url, params=params, timeout=5)
            response.raise_for_status()
            
            data = response.json()
            
            if data.get('success') and data.get('data'):
                ticker_data = data['data']
                if isinstance(ticker_data, list) and len(ticker_data) > 0:
                    return float(ticker_data[0].get('lastPrice', 0))
                elif isinstance(ticker_data, dict):
                    return float(ticker_data.get('lastPrice', 0))
            
            return 0.0
            
        except Exception as e:
            print(f"⚠ Ошибка получения цены для {symbol}: {e}")
            return 0.0
    
    def parse_launch_date(self, timestamp_ms: int) -> str:
        """Конвертирует timestamp в дату"""
        try:
            if timestamp_ms > 10000000000:  # Если в миллисекундах
                timestamp_ms = timestamp_ms / 1000
            
            dt = datetime.fromtimestamp(timestamp_ms)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return "N/A"
    
    def scrape_and_save(self, output_dir: str = "mexc"):
        """Главная функция - собирает данные и сохраняет в Excel файлы по плечам"""

        print("🚀 Запуск MEXC Futures Scraper...")
        print("=" * 60)

        # Получаем список фьючей
        futures_list = self.get_futures_list()

        if not futures_list:
            print("✗ Не удалось получить список фьючей")
            return

        # Обработка данных
        data = []
        total = len(futures_list)

        for idx, future in enumerate(futures_list, 1):
            try:
                symbol = future.get('symbol', 'N/A')

                # Информация о плече
                max_leverage = future.get('maxLeverage', future.get('maxOpenLeverage', 'N/A'))

                # Дата листинга
                listing_time = future.get('listingTime', future.get('createTime', 0))
                launch_date = self.parse_launch_date(listing_time)

                # Текущая цена
                print(f"[{idx}/{total}] Получаю цену для {symbol}...", end='', flush=True)
                current_price = self.get_current_price(symbol)
                print(f" ✓ {current_price}")

                # Дополнительная информация
                quote_asset = future.get('quoteAsset', 'USDT')
                base_asset = future.get('baseAsset', symbol.replace('USDT', ''))
                status = future.get('status', 'N/A')

                data.append({
                    'Символ': symbol.split('_')[0],
                    'Базовый актив': base_asset,
                    'Котируемый актив': quote_asset,
                    'Максимальное плечо': max_leverage,
                    'Дата листинга': launch_date,
                    'Текущая цена': current_price,
                    'Статус': status,
                })

                time.sleep(0.1)  # Чтобы не перегружать API

            except Exception as e:
                print(f"\n⚠ Ошибка обработки {symbol}: {e}")
                continue

        df = pd.DataFrame(data)
        df = df.sort_values('Символ').reset_index(drop=True)

        save_by_leverage(df, output_dir, prefix="mexc")

def save_by_leverage(df: pd.DataFrame, output_dir: str, prefix: str) -> None:
    """Группирует DataFrame по колонке 'Максимальное плечо' и сохраняет
    по одному xlsx-файлу на плечо в папку output_dir."""

    os.makedirs(output_dir, exist_ok=True)

    def leverage_key(value):
        if value is None:
            return None
        s = str(value).strip()
        if not s or s.upper() in {"N/A", "NAN", "NONE"}:
            return None
        m = re.search(r"\d+", s)
        return int(m.group()) if m else None

    df = df.copy()
    df["_lev"] = df["Максимальное плечо"].map(leverage_key)

    known = df[df["_lev"].notna()]
    unknown = df[df["_lev"].isna()]

    column_widths = {
        'A': 15, 'B': 15, 'C': 18, 'D': 16, 'E': 20, 'F': 15, 'G': 12,
    }

    written = []

    for lev, group in known.groupby("_lev", sort=True):
        lev_int = int(lev)
        out_path = os.path.join(output_dir, f"{prefix}_{lev_int}.xlsx")
        out_df = group.drop(columns=["_lev"]).sort_values("Символ").reset_index(drop=True)
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            out_df.to_excel(writer, sheet_name='Futures', index=False)
            ws = writer.sheets['Futures']
            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width
        written.append((out_path, len(out_df)))

    if not unknown.empty:
        out_path = os.path.join(output_dir, f"{prefix}_unknown.xlsx")
        out_df = unknown.drop(columns=["_lev"]).sort_values("Символ").reset_index(drop=True)
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            out_df.to_excel(writer, sheet_name='Futures', index=False)
            ws = writer.sheets['Futures']
            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width
        written.append((out_path, len(out_df)))

    print("=" * 60)
    print(f"✓ Файлы сохранены в папку: {output_dir}/")
    for path, n in written:
        print(f"   {os.path.basename(path)}  —  {n} шт.")
    print(f"✓ Всего фьючей: {len(df)}")
    print("=" * 60)


def main():
    scraper = MEXCFuturesScraper(
        access_key=ACCESS_KEY,
        secret_key=SECRET_KEY
    )
    scraper.scrape_and_save(output_dir="mexc")

if __name__ == "__main__":
    main()