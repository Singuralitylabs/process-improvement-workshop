#!/usr/bin/env python3
"""第2回で使用する、整形を体験できる架空の売上Excelを作る。"""

# 標準ライブラリ: --overwriteなどのコマンドライン引数を読み取る。
import argparse
# 標準ライブラリ: datetimeで日付を作り、timedeltaで日数を加算する。
from datetime import datetime, timedelta
# 標準ライブラリ: ファイル・フォルダのパスを組み立て、存在確認や作成を行う。
from pathlib import Path

# 追加ライブラリ: Workbookで新しいExcelブックを作成し、保存する。
from openpyxl import Workbook
# 追加ライブラリ: 配置・罫線・フォント・背景色・罫線の各辺の書式を設定する。
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


ROOT = Path(__file__).resolve().parent
INPUT_PATH = ROOT / "input" / "売上データ.xlsx"
STORES = ["東京駅前店", "横浜店", "大宮店"]
PRODUCTS = [
    ("コーヒー豆", 1200),
    ("緑茶", 800),
    ("クッキー", 450),
    ("マグカップ", 1500),
    ("ドリップバッグ", 600),
]


def create_sample(path: Path) -> None:
    """3店舗・3日・5品目の45明細と、途中の空白行を保存する。

    入力: path はサンプルExcelの保存先パス。
    出力: 戻り値はなし。売上明細シートを持つExcelを path に保存する。
    """
    # 新しいExcelブックを作る。
    workbook = Workbook()
    # 最初から用意されているシートを取得する。
    sheet = workbook.active
    # 集計コードが読み込むシート名を設定する。
    sheet.title = "売上明細"
    # 1行目に5列の見出しを書き込む。
    sheet.append(["日付", "店舗", "品目", "単価", "数量"])
    # 3日分の日番号を0から順に取り出す。
    for day_offset in range(3):
        # 基準日の11月1日に日番号を加え、明細の日付を決める。
        sales_date = datetime(2026, 11, 1) + timedelta(days=day_offset)
        # 各店舗の番号と名前を順に取り出す。
        for store_index, store in enumerate(STORES):
            # 各品目の番号・名前・単価を順に取り出す。
            for product_index, (product, price) in enumerate(PRODUCTS):
                # 日・店舗・品目の番号から、毎回同じ1〜4個の数量を決める。
                quantity = (day_offset + store_index + product_index) % 4 + 1
                # 整形の練習用に空白や文字列の数値を含む明細を1行追加する。
                sheet.append([
                    # 明細の日付を日付型のまま保存する。
                    sales_date,
                    # 最初の品目だけ店舗名の前後に半角空白を入れる。
                    " {} ".format(store) if product_index == 0 else store,
                    # 最初の店舗だけ品目名の前後に全角空白を入れる。
                    "　{}　".format(product) if store_index == 0 else product,
                    # 最初の品目の単価は、数値ではなく文字列として保存する。
                    str(price) if product_index == 0 else price,
                    # 最初の品目の数量も文字列として保存する。
                    str(quantity) if product_index == 0 else quantity,
                ])
        # 1日目の明細をすべて書き終えたか確認する。
        if day_offset == 0:
            # 空白行の除外を練習するため、5列とも空欄の行を入れる。
            sheet.append([None] * 5)
    # セルの罫線に使う薄い灰色の細線を定義する。
    edge = Side(style="thin", color="D9D9D9")
    # セルの四辺に同じ罫線を設定するための書式を作る。
    border = Border(left=edge, right=edge, top=edge, bottom=edge)
    # 見出しと空白行を含む各行を順に取り出す。
    for row in sheet.iter_rows():
        # 行の高さを22ポイントに統一する。
        sheet.row_dimensions[row[0].row].height = 22
        # 行に含まれるセルを1つずつ取り出す。
        for cell in row:
            # 同じ列の見出しを取得して、列に応じた書式を判断する。
            heading = sheet.cell(1, cell.column).value
            # フォントを游ゴシック11ポイントに統一し、見出しだけ太字にする。
            cell.font = Font(name="Yu Gothic", size=11, bold=cell.row == 1)
            # セルの四辺に罫線を適用する。
            cell.border = border
            # 見出しの背景を灰色、データ行の背景を白にする。
            cell.fill = PatternFill(fill_type="solid", fgColor="EDEDED" if cell.row == 1 else "FFFFFF")
            # 名前や見出しの基本配置を左寄せにする。
            horizontal = "left"
            # 基本の表示形式を標準にする。
            cell.number_format = "General"
            # 見出し以外のセルだけ、列別の表示形式に切り替える。
            if cell.row > 1:
                # 日付列かどうかを確認する。
                if heading == "日付":
                    # 日付を中央に配置する。
                    horizontal = "center"
                    # 日付の表示を年-月-日に統一する。
                    cell.number_format = "yyyy-mm-dd"
                # 単価または数量の列かどうかを確認する。
                elif heading in ("単価", "数量"):
                    # 文字列の数値も含めて右寄せに統一する。
                    horizontal = "right"
                    # 数値型の値を桁区切り付きの整数で表示する。
                    cell.number_format = "#,##0"
            # 横の配置を反映し、縦は中央、折り返しはなしに統一する。
            cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=False)
    # スクロールしても1行目の見出しが見えるように固定する。
    sheet.freeze_panes = "A2"
    # 使用範囲全体に絞り込み用のフィルターを設定する。
    sheet.auto_filter.ref = sheet.dimensions
    # 5列の列記号を順に取り出す。
    for column in ("A", "B", "C", "D", "E"):
        # 列幅を22に統一する。
        sheet.column_dimensions[column].width = 22
    # 保存先フォルダがなければ親フォルダも含めて作成する。
    path.parent.mkdir(parents=True, exist_ok=True)
    # 完成したExcelを指定パスに保存する。
    workbook.save(path)


def main() -> None:
    """コマンドライン引数を読み、サンプルExcelを作成する。

    入力: コマンドラインの --overwrite オプション（省略可能）。
    出力: 戻り値はなし。入力Excelを保存し、件数と保存先を標準出力に表示する。
    """
    # コマンドライン引数を受け取るための解析器を作る。
    parser = argparse.ArgumentParser(description="第2回の架空売上Excelを作成します。")
    # 入力Excelの上書きを許可するオプションを登録する。
    parser.add_argument("--overwrite", action="store_true", help="編集済みの入力Excelも初期化します")
    # 実行時に指定されたオプションを読み取る。
    args = parser.parse_args()
    # 入力Excelが存在し、上書きが許可されていないかを確認する。
    if INPUT_PATH.exists() and not args.overwrite:
        # 編集済みの入力Excelを守るため、理由を表示して終了する。
        raise SystemExit("入力Excelは既にあります。初期化する場合のみ --overwrite を指定してください。")
    # 保存時のエラーを捕まえられるように、作成処理を実行する。
    try:
        # 架空の売上明細を作り、指定した入力Excelに保存する。
        create_sample(INPUT_PATH)
    # ファイルの保存に失敗した場合のエラーを受け取る。
    except OSError as error:
        # Excelの使用状態と書込権限の確認を案内して終了する。
        raise SystemExit("保存できません。Excelを閉じ、書込権限を確認してください: {}".format(error))
    # 作成したサンプルの明細数と構成を表示する。
    print("サンプル明細: 45件 / 3店舗 / 3日 / 5品目")
    # 入力Excelの保存先を表示する。
    print("入力Excel: {}".format(INPUT_PATH))


# このファイルを直接実行した場合だけ、メイン処理を開始する。
if __name__ == "__main__":
    # コマンドライン引数の解析とサンプル作成を実行する。
    main()