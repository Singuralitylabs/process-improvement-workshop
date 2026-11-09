#!/usr/bin/env python3
"""保存済みExcelをpandasで読み取り、整形・集計して別のExcelへ出力する。"""

# 標準ライブラリ: 入出力ファイルのパスを扱い、保存先フォルダを作成する。
from pathlib import Path
# 標準ライブラリ: シート名と表の辞書であることを型ヒントで示す。
from typing import Dict
# 標準ライブラリ: ZIP形式のExcelファイルが壊れている場合の例外を捕まえる。
from zipfile import BadZipFile

# 追加ライブラリ: 表データの読込・整形・集計・Excel出力に使い、pdと略記する。
import pandas as pd
# 追加ライブラリ: 配置・罫線・フォント・背景色・罫線の各辺の書式を設定する。
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


ROOT = Path(__file__).resolve().parent
INPUT_PATH = ROOT / "input" / "売上データ.xlsx"
OUTPUT_PATH = ROOT / "output" / "売上集計.xlsx"
REQUIRED_COLUMNS = ["日付", "店舗", "品目", "単価", "数量"]


def load_sales(path: Path) -> pd.DataFrame:
    """売上明細シートを見出し名で読み、空白行と不要列を除く。

    入力: path は読み込むExcelのパス。
    出力: 必須5列だけを持ち、全列が空欄の行を除いたDataFrame。
    """
    # 保存済みExcelの売上明細シートを、見出し付きの表として読み込む。
    raw = pd.read_excel(path, sheet_name="売上明細", engine="openpyxl")
    # 必須列のうち、入力Excelの見出しに存在しない列名を集める。
    missing = [column for column in REQUIRED_COLUMNS if column not in raw.columns]
    # 不足している列があるか確認する。
    if missing:
        # 不足列の名前を示し、集計せずに呼び出し元へエラーを通知する。
        raise ValueError("売上明細シートに必要な列がありません: {}".format("、".join(missing)))
    # 全列が空欄の行を除き、必須5列だけを選んだ独立した表を作る。
    sales = raw.dropna(how="all").loc[:, REQUIRED_COLUMNS].copy()
    # 空白行を除いた後に明細が残っているか確認する。
    if sales.empty:
        # 明細がない場合はエラーで止める。
        raise ValueError("売上明細シートに明細がありません。")
    # 読み込んだ明細の表を呼び出し元に返す。
    return sales


def clean_sales(sales: pd.DataFrame) -> pd.DataFrame:
    """前後の空白・日付・数値を整え、不正な明細は集計前に止める。

    入力: 必須5列を持つ売上明細のDataFrame。
    出力: 値を整形し、売上金額列を追加して日付・店舗・品目順に並べたDataFrame。
    入力のDataFrameは変更せず、不正な値がある場合はValueErrorを送出する。
    """
    # 入力の表を変更しないように、整形用のコピーを作る。
    sales = sales.copy()
    # 名前を整える店舗列と品目列を順に処理する。
    for column in ("店舗", "品目"):
        # 文字列型に変換し、前後の半角・全角空白を取り除く。
        sales[column] = sales[column].astype("string").str.strip()
        # 欠損値または空白除去後の空文字が1件でもあるか確認する。
        if sales[column].isna().any() or sales[column].eq("").any():
            # 空欄の列名を示して処理を止める。
            raise ValueError("{}が空欄の明細があります。入力Excelを確認してください。".format(column))
    # 日付型に変換し、時刻を午前0時に揃える。不正な値は欠損値にする。
    sales["日付"] = pd.to_datetime(sales["日付"], errors="coerce").dt.normalize()
    # 変換できなかった日付や空欄があるか確認する。
    if sales["日付"].isna().any():
        # 日付の入力形式を案内して処理を止める。
        raise ValueError("日付が空欄または不正です。Excelの日付セルか YYYY-MM-DD で入力してください。")
    # 単価と数量を順に数値へ変換する。
    for column in ("単価", "数量"):
        # 文字列の数値を数値型に変換し、変換できない値は欠損値にする。
        values = pd.to_numeric(sales[column], errors="coerce")
        # 欠損・負数・上限超過・小数のいずれかに該当する行を判定する。
        invalid = values.isna() | (values < 0) | (values > 1000000) | (values % 1 != 0)
        # 不正な行が1件でもあるか確認する。
        if invalid.any():
            # 許可する数値の範囲を示して処理を止める。
            raise ValueError("{}は0〜1,000,000の整数で入力してください。".format(column))
        # 検証済みの値を64ビット整数として表に保存する。
        sales[column] = values.astype("int64")
    # 各行の単価と数量を掛けて売上金額列を追加する。
    sales["売上金額"] = sales["単価"] * sales["数量"]
    # 日付・店舗・品目順に並べ、行番号を0から振り直した表を返す。
    return sales.sort_values(["日付", "店舗", "品目"]).reset_index(drop=True)


def summarize_sales(sales: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """同じキーの明細をまとめ、数量と売上金額の合計を計算する。

    入力: clean_sales が返す、売上金額列を持つ整形済み明細のDataFrame。
    出力: シート名をキー、整形済み明細と3種類の集計DataFrameを値とする辞書。
    """
    # 同じ店舗の行をまとめ、数量と売上金額を合計する。
    store_summary = sales.groupby("店舗", as_index=False)[["数量", "売上金額"]].sum()
    # 日付と店舗の組み合わせごとに、数量と売上金額を合計する。
    daily_summary = sales.groupby(["日付", "店舗"], as_index=False)[["数量", "売上金額"]].sum()
    # 同じ品目の行をまとめ、数量と売上金額を合計する。
    product_summary = sales.groupby("品目", as_index=False)[["数量", "売上金額"]].sum()
    # 出力シート名と書き込む表を対応付けた辞書を返す。
    return {
        # 明細シートには整形後の全明細を渡す。
        "整形済み明細": sales,
        # 店舗別シートには店舗ごとの合計を渡す。
        "店舗別集計": store_summary,
        # 日別店舗別シートには日付と店舗ごとの合計を渡す。
        "日別店舗別集計": daily_summary,
        # 品目別シートには品目ごとの合計を渡す。
        "品目別集計": product_summary,
    }


def write_report(tables: Dict[str, pd.DataFrame], path: Path) -> None:
    """入力Excelとは別のファイルへ、4シートの集計結果を保存する。

    入力: tables はシート名とDataFrameの辞書、path は出力Excelの保存先パス。
    出力: 戻り値はなし。辞書の各表を別シートに書き出したExcelを path に保存する。
    """
    # 出力先フォルダがなければ親フォルダも含めて作成する。
    path.parent.mkdir(parents=True, exist_ok=True)
    # Excel出力用の書込処理を開き、withを抜けるときに保存して閉じる。
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        # シート名と書き込む表を辞書から順に取り出す。
        for sheet_name, table in tables.items():
            # 表を指定シートへ書き込み、DataFrameの行番号は出力しない。
            table.to_excel(writer, sheet_name=sheet_name, index=False)
            # 作成したシートを取得し、書式を設定できるようにする。
            sheet = writer.sheets[sheet_name]
            # スクロールしても1行目の見出しが見えるように固定する。
            sheet.freeze_panes = "A2"
            # 使用範囲全体に絞り込み用のフィルターを設定する。
            sheet.auto_filter.ref = sheet.dimensions
            # 罫線に使う薄い灰色の細線を定義する。
            edge = Side(style="thin", color="D9D9D9")
            # セルの四辺に同じ罫線を設定するための書式を作る。
            border = Border(left=edge, right=edge, top=edge, bottom=edge)
            # 見出しを含む各行を順に取り出す。
            for row in sheet.iter_rows():
                # 行の高さを22ポイントに統一する。
                sheet.row_dimensions[row[0].row].height = 22
                # 行に含まれるセルを1つずつ取り出す。
                for cell in row:
                    # 同じ列の見出しを取得して、列別の書式を判断する。
                    heading = sheet.cell(1, cell.column).value
                    # フォントを游ゴシック11ポイントにし、見出しだけ太字にする。
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
                        # 単価・数量・売上金額の列かどうかを確認する。
                        elif heading in ("単価", "数量", "売上金額"):
                            # 数値を右寄せにする。
                            horizontal = "right"
                            # 数値を桁区切り付きの整数で表示する。
                            cell.number_format = "#,##0"
                    # 横の配置を反映し、縦は中央、折り返しはなしに統一する。
                    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=False)
            # シートの各列を順に取り出す。
            for column in sheet.columns:
                # 列幅を22に統一する。
                sheet.column_dimensions[column[0].column_letter].width = 22


def main() -> None:
    """入力Excelの読込・整形・集計・出力を順に実行する。

    入力: INPUT_PATH が示す、売上明細シートを持つ保存済みExcel。
    出力: 戻り値はなし。OUTPUT_PATH に集計Excelを保存し、結果を標準出力に表示する。
    """
    # 読込から保存までのエラーを捕まえられるように、処理を実行する。
    try:
        # 保存済みの入力Excelを読み、空白行と不要列を除く。
        sales = load_sales(INPUT_PATH)
        # 店舗名・品目名・日付・数値を整え、売上金額を計算する。
        sales = clean_sales(sales)
        # 店舗別・日別店舗別・品目別の数量と売上金額を集計する。
        tables = summarize_sales(sales)
        # 整形済み明細と3種類の集計表を、別のExcelに保存する。
        write_report(tables, OUTPUT_PATH)
    # 入力Excelが見つからない場合のエラーを受け取る。
    except FileNotFoundError:
        # サンプル作成コマンドを案内して終了する。
        raise SystemExit("入力Excelがありません。先に python create_sales_data.py を実行してください。")
    # ファイルを読み書きする権限がない場合のエラーを受け取る。
    except PermissionError:
        # Excelの使用状態とファイルの権限の確認を案内して終了する。
        raise SystemExit("Excelを閉じ、入力の読取権限と出力先の書込権限を確認してください。")
    # 不正な値・シート・Excel形式や入出力のエラーを受け取る。
    except (ValueError, OSError, BadZipFile) as error:
        # 入力Excelと保存先の確認を案内し、エラーの詳細を表示して終了する。
        raise SystemExit("集計できません。入力Excel・シート・保存先を確認してください: {}".format(error))
    # 集計した明細数と全明細の売上合計を表示する。
    print("集計明細: {}件 / 売上合計: {:,}円".format(len(sales), int(sales["売上金額"].sum())))
    # 店舗別の数量と売上金額を、行番号なしで表示する。
    print(tables["店舗別集計"].to_string(index=False))
    # 集計Excelの保存先を表示する。
    print("出力Excel: {}".format(OUTPUT_PATH))


# このファイルを直接実行した場合だけ、メイン処理を開始する。
if __name__ == "__main__":
    # 読込・整形・集計・保存と結果表示を実行する。
    main()