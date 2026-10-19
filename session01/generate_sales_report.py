#!/usr/bin/env python3
"""Read the generated sales workbook and create Excel, PDF, and daily summaries."""

from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont


# スクリプトの場所を基準にすることで、どの作業ディレクトリからでも同じ場所へ出力する。
ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
IMAGE_DIR = OUTPUT / "images"
STORES = ["東京駅前店", "横浜店", "大宮店"]
HEADERS = ["日付", "店舗", "顧客ID", "品目", "単価", "数量", "売上金額"]
FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Medium.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
    "/System/Library/Fonts/ヒラギノ角ゴ ProN W3.otf",
    "C:/Windows/Fonts/YuGothM.ttc",
    "C:/Windows/Fonts/meiryo.ttc",
]


def group_sales(rows: List[Dict[str, object]]) -> Dict[Tuple[str, str], List[Dict[str, object]]]:
    """日付と店舗の組み合わせをキーにして明細をまとめる。"""
    grouped = defaultdict(list)
    for row in rows:
        grouped[(str(row["日付"]), str(row["店舗"]))].append(row)
    return dict(grouped)


def top_products(rows: List[Dict[str, object]], limit: int = 3) -> List[Tuple[str, int]]:
    """品目ごとの販売数量を合算し、同数の場合は品目名順で返す。"""
    quantities = Counter()
    for row in rows:
        quantities[str(row["品目"])] += int(row["数量"])
    return sorted(quantities.items(), key=lambda item: (-item[1], item[0]))[:limit]


def read_sales_workbook(input_path: Path) -> List[Dict[str, object]]:
    """入力ブックのシート構成と見出しを検証し、明細を辞書の一覧にする。"""
    # 読み取り専用モードで開き、元ファイルを変更せずに処理する。
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    if len(workbook.sheetnames) != 9:
        raise ValueError("入力Excelには店舗・日別の9シートが必要です")

    rows = []
    for sheet in workbook.worksheets:
        # 列の順番が変わっても誤集計しないよう、見出しを完全一致で確認する。
        header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        if list(header) != HEADERS:
            raise ValueError("{} の1行目が想定した見出しと一致しません".format(sheet.title))
        for values in sheet.iter_rows(min_row=2, values_only=True):
            # Excel末尾に混じる空行は集計対象から除外する。
            if not any(value is not None for value in values):
                continue
            rows.append(dict(zip(HEADERS, values)))
    workbook.close()
    if not rows:
        raise ValueError("入力Excelに売上明細がありません")
    return rows


def add_sheet_summary(sheet, rows: List[Dict[str, object]], start_row: int) -> None:
    """明細シートの下に売上金額・顧客数・売れ筋3品目を追記する。"""
    total_amount = sum(int(row["売上金額"]) for row in rows)
    customer_count = len({str(row["顧客ID"]) for row in rows})
    sheet.cell(start_row, 1, "集計")
    sheet.cell(start_row + 1, 1, "売上金額")
    sheet.cell(start_row + 1, 2, total_amount)
    sheet.cell(start_row + 2, 1, "顧客数")
    sheet.cell(start_row + 2, 2, customer_count)
    sheet.cell(start_row + 4, 1, "品目")
    sheet.cell(start_row + 4, 2, "販売個数")
    for offset, (product, quantity) in enumerate(top_products(rows), start=1):
        sheet.cell(start_row + 4 + offset, 1, product)
        sheet.cell(start_row + 4 + offset, 2, quantity)


def write_workbook(grouped: Dict[Tuple[str, str], List[Dict[str, object]]]) -> None:
    """集計結果を全体サマリーと店舗・日別シートに書き出す。"""
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "全体サマリー"
    summary_sheet.append(["日付", "店舗", "売上金額", "顧客数", "上位品目", "販売個数"])

    for (sales_date, store), rows in sorted(grouped.items()):
        # 入力ブックと同じ店舗・日付単位のシートを作り、明細と集計を並べる。
        sheet = workbook.create_sheet("{}_{}".format(sales_date[5:], store))
        sheet.append(HEADERS)
        for row in rows:
            sheet.append([row[header] for header in HEADERS])
        summary_row = len(rows) + 3
        add_sheet_summary(sheet, rows, summary_row)
        format_sheet(sheet, len(HEADERS), summary_row + 7)

        total_amount = sum(int(row["売上金額"]) for row in rows)
        customer_count = len({str(row["顧客ID"]) for row in rows})
        for product, quantity in top_products(rows):
            # サマリーは1店舗につき上位3品目を3行で記録する。
            summary_sheet.append([sales_date, store, total_amount, customer_count, product, quantity])

    format_sheet(summary_sheet, 6, summary_sheet.max_row)
    workbook.save(OUTPUT / "売上報告.xlsx")


def format_sheet(sheet, column_count: int, last_row: int) -> None:
    """見出し・列幅・金額表示を整え、明細部分にフィルターを付ける。"""
    header_fill = PatternFill("solid", fgColor="174A45")
    summary_fill = PatternFill("solid", fgColor="DCEFEA")
    bottom_border = Border(bottom=Side(style="thin", color="C8D7D3"))
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(name="Yu Gothic", color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in sheet.iter_rows(min_row=2, max_row=last_row, max_col=column_count):
        for cell in row:
            cell.border = bottom_border
            if cell.column in (5, 7) and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0"円"' if cell.column == 7 else '#,##0"円"'
    # 明細のオートフィルターが集計行まで含まないよう、集計開始行を探す。
    for row in sheet.iter_rows():
        if row and str(row[0].value) == "集計":
            for cell in row:
                cell.fill = summary_fill
                cell.font = Font(name="Yu Gothic", bold=True, color="174A45")
    widths = [14, 18, 12, 20, 14, 10, 16]
    for column in range(1, column_count + 1):
        sheet.column_dimensions[get_column_letter(column)].width = widths[column - 1]
    sheet.freeze_panes = "A2"
    summary_start = next(
        (row[0].row for row in sheet.iter_rows(min_col=1, max_col=1) if row[0].value == "集計"),
        sheet.max_row + 1,
    )
    filter_end = min(summary_start - 1, sheet.max_row)
    if filter_end > 1:
        sheet.auto_filter.ref = "A1:{}{}".format(get_column_letter(column_count), filter_end)


def load_font(size: int) -> ImageFont.FreeTypeFont:
    """OSに入っている日本語フォントを候補から探す。"""
    for font_path in FONT_CANDIDATES:
        if Path(font_path).is_file():
            return ImageFont.truetype(font_path, size, index=0)
    raise FileNotFoundError("日本語対応フォントが見つかりません。Noto Sans CJKなどをインストールしてください")


def draw_table(
    title: str,
    headers: List[str],
    rows: List[List[str]],
    output_path: Path,
    widths: List[int],
) -> None:
    """記事に載せる表の見本をPNG画像として描画する。"""
    row_height = 58
    margin = 48
    image = Image.new("RGB", (sum(widths) + margin * 2, 166 + row_height * (len(rows) + 1)), "#F5F8F7")
    draw = ImageDraw.Draw(image)
    draw.text((margin, 36), title, font=load_font(32), fill="#173F3A")
    y = 104
    x = margin
    for index, header in enumerate(headers):
        draw.rectangle((x, y, x + widths[index], y + row_height), fill="#174A45")
        draw.text((x + 12, y + 10), header, font=load_font(19), fill="#FFFFFF")
        x += widths[index]
    y += row_height
    for row_index, values in enumerate(rows):
        x = margin
        fill = "#FFFFFF" if row_index % 2 == 0 else "#EAF1EF"
        for index, value in enumerate(values):
            draw.rectangle((x, y, x + widths[index], y + row_height), fill=fill, outline="#CFDAD7")
            draw.text((x + 12, y + 11), value, font=load_font(19), fill="#243B37")
            x += widths[index]
        y += row_height
    image.save(output_path)


def write_preview_images(grouped: Dict[Tuple[str, str], List[Dict[str, object]]]) -> None:
    """代表シートの明細先頭3行と集計結果を、それぞれ画像にする。"""
    (sales_date, store), rows = sorted(grouped.items())[0]
    sample_rows = [
        [
            str(row["顧客ID"]),
            str(row["品目"]),
            "{:,}円".format(int(row["単価"])),
            str(row["数量"]),
            "{:,}円".format(int(row["売上金額"])),
        ]
        for row in rows[:3]
    ]
    draw_table(
        "売上明細 | {} {}".format(sales_date, store),
        ["顧客ID", "品目", "単価", "数量", "売上金額"],
        sample_rows,
        IMAGE_DIR / "sheet_preview.png",
        [190, 250, 170, 120, 200],
    )

    total_amount = sum(int(row["売上金額"]) for row in rows)
    customer_count = len({str(row["顧客ID"]) for row in rows})
    summary_rows = [
        ["売上金額", "{:,}円".format(total_amount)],
        ["顧客数", "{}人".format(customer_count)],
    ] + [[product, "{}個".format(quantity)] for product, quantity in top_products(rows)]
    draw_table(
        "集計結果 | {} {}".format(sales_date, store),
        ["項目", "集計値"],
        summary_rows,
        IMAGE_DIR / "summary_preview.png",
        [420, 360],
    )


def draw_pdf_page(sales_date: str, grouped: Dict[Tuple[str, str], List[Dict[str, object]]]) -> Image.Image:
    """指定日の3店舗分の集計をA4比率の1ページ画像として作る。"""
    image = Image.new("RGB", (1240, 1754), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    draw.text((72, 54), "日次売上レポート", font=load_font(42), fill="#173F3A")
    draw.text((74, 118), sales_date, font=load_font(27), fill="#52716B")
    y = 190

    for store in STORES:
        rows = grouped[(sales_date, store)]
        total_amount = sum(int(row["売上金額"]) for row in rows)
        customer_count = len({str(row["顧客ID"]) for row in rows})
        draw.rounded_rectangle((64, y, 1176, y + 394), radius=10, fill="#F5F8F7", outline="#D5E0DD", width=2)
        draw.text((92, y + 22), store, font=load_font(30), fill="#174A45")
        draw.text((92, y + 82), "売上金額  {:,}円".format(total_amount), font=load_font(24), fill="#243B37")
        draw.text((650, y + 82), "顧客数  {}人".format(customer_count), font=load_font(24), fill="#243B37")
        draw.text((92, y + 148), "販売個数 上位3品目", font=load_font(22), fill="#52716B")
        top_rows = top_products(rows)
        for rank, (product, quantity) in enumerate(top_rows, start=1):
            item_y = y + 202 + (rank - 1) * 54
            draw.text((110, item_y), "{}. {}".format(rank, product), font=load_font(23), fill="#243B37")
            draw.text((850, item_y), "{}個".format(quantity), font=load_font(23), fill="#243B37")
        y += 430

    draw.text((72, 1510), "対象: 3店舗・各店10顧客のサンプルデータ", font=load_font(18), fill="#71827E")
    return image


def write_pdf_and_manager_summary(
    grouped: Dict[Tuple[str, str], List[Dict[str, object]]],
) -> None:
    """日付別PDFと、1日につき3行程度の上司向け要約を保存する。"""
    dates = sorted({sales_date for sales_date, _ in grouped})
    pages = [draw_pdf_page(sales_date, grouped) for sales_date in dates]
    pdf_path = OUTPUT / "日次売上レポート.pdf"
    pages[0].save(pdf_path, "PDF", resolution=150, save_all=True, append_images=pages[1:])

    summary_lines = []
    for sales_date in dates:
        if summary_lines:
            summary_lines.append("")
        day_rows = [row for (row_date, _), rows in grouped.items() if row_date == sales_date for row in rows]
        store_totals = []
        for store in STORES:
            amount = sum(int(row["売上金額"]) for row in grouped[(sales_date, store)])
            store_totals.append("{} {:,}円".format(store, amount))
        total_amount = sum(int(row["売上金額"]) for row in day_rows)
        customer_count = len({(str(row["店舗"]), str(row["顧客ID"])) for row in day_rows})
        top_items = "、".join("{} {}個".format(product, quantity) for product, quantity in top_products(day_rows))
        summary_lines.extend(
            [
                "{} 店舗別売上: {}".format(sales_date, " / ".join(store_totals)),
                "全店売上: {:,}円 / 顧客数: {}人".format(total_amount, customer_count),
                "販売上位: {}".format(top_items),
            ]
        )
    (OUTPUT / "上司向け日報.txt").write_text("\n".join(summary_lines) + "\n", encoding="utf-8-sig")


def main() -> None:
    """作成済みExcelを読み直し、集計結果とレポートを保存する。"""
    OUTPUT.mkdir(parents=True, exist_ok=True)
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    input_path = OUTPUT / "売上データ.xlsx"
    if not input_path.exists():
        raise FileNotFoundError("入力用Excelがありません。先に create_sales_data.py を実行してください")
    # Excelから再読込するため、ブックを編集した場合もその内容でレポートを作れる。
    rows = read_sales_workbook(input_path)
    grouped = group_sales(rows)
    write_workbook(grouped)
    write_preview_images(grouped)
    write_pdf_and_manager_summary(grouped)
    print("Excel読込: {}行 / 店舗・日別シート: {}".format(len(rows), len(grouped)))
    print("出力先: {}".format(OUTPUT))


if __name__ == "__main__":
    main()