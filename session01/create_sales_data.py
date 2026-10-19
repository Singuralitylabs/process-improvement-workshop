#!/usr/bin/env python3
"""Create reproducible sample sales data and a nine-sheet Excel workbook."""

import csv
import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
TSV_DIR = OUTPUT / "tsv"
START_DATE = date(2026, 10, 1)
RANDOM_SEED = 20261019
STORES = ["東京駅前店", "横浜店", "大宮店"]
PRODUCTS = [
    ("コーヒー豆", 1200),
    ("緑茶", 800),
    ("クッキー", 450),
    ("マグカップ", 1500),
    ("ドリップバッグ", 600),
]
HEADERS = ["日付", "店舗", "顧客ID", "品目", "単価", "数量", "売上金額"]


def generate_sales_data() -> List[Dict[str, object]]:
    """店舗・日ごとに10顧客分の架空明細を作る。"""
    rng = random.Random(RANDOM_SEED)
    rows = []
    for day_offset in range(3):
        sales_date = START_DATE + timedelta(days=day_offset)
        for store in STORES:
            for customer_number in range(1, 11):
                product_count = rng.randint(1, 3)
                for product_name, unit_price in rng.sample(PRODUCTS, product_count):
                    quantity = rng.randint(1, 4)
                    rows.append(
                        {
                            "日付": sales_date.isoformat(),
                            "店舗": store,
                            "顧客ID": "C{:03d}".format(customer_number),
                            "品目": product_name,
                            "単価": unit_price,
                            "数量": quantity,
                            "売上金額": unit_price * quantity,
                        }
                    )
    return rows


def group_sales(rows: List[Dict[str, object]]) -> Dict[Tuple[str, str], List[Dict[str, object]]]:
    """日付と店舗の組み合わせをキーにして明細をまとめる。"""
    grouped = defaultdict(list)
    for row in rows:
        grouped[(str(row["日付"]), str(row["店舗"]))].append(row)
    return dict(grouped)


def write_tsv_files(grouped: Dict[Tuple[str, str], List[Dict[str, object]]]) -> None:
    """店舗・日別の明細をExcelに貼り付けやすいTSVとして保存する。"""
    for (sales_date, store), rows in grouped.items():
        path = TSV_DIR / "{}_{}.tsv".format(sales_date, store)
        with path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=HEADERS, delimiter="\t", lineterminator="\n")
            writer.writeheader()
            writer.writerows(rows)


def write_input_workbook(grouped: Dict[Tuple[str, str], List[Dict[str, object]]]) -> Path:
    """集計前の入力用ブックを作り、店舗・日別に9シートへ分ける。"""
    workbook = Workbook()
    workbook.remove(workbook.active)
    for (sales_date, store), rows in sorted(grouped.items()):
        sheet = workbook.create_sheet("{}_{}".format(sales_date[5:], store))
        sheet.append(HEADERS)
        for row in rows:
            sheet.append([row[header] for header in HEADERS])
        sheet.freeze_panes = "A2"
    input_path = OUTPUT / "売上データ.xlsx"
    workbook.save(input_path)
    return input_path


def main() -> None:
    """架空明細を作り、TSVと9シートのExcelへ保存する。"""
    OUTPUT.mkdir(parents=True, exist_ok=True)
    TSV_DIR.mkdir(parents=True, exist_ok=True)
    rows = generate_sales_data()
    grouped = group_sales(rows)
    write_tsv_files(grouped)
    input_path = write_input_workbook(grouped)
    print("サンプル明細: {}行 / TSV: {}件 / Excelシート: {}".format(
        len(rows), len(grouped), len(grouped)
    ))
    print("入力Excel: {}".format(input_path))


if __name__ == "__main__":
    main()