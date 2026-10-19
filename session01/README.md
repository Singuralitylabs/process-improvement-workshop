# 第1回: Pythonで小売店の日次売上を集計

営業本部の集計担当者を想定し、架空の売上データをPythonで作成してExcelに保存し、そのExcelをあらためて読み込んで、店舗別レポートと上司向け要約を作成します。

## 内容

- `create_sales_data.py`: 再現可能な架空データを作成し、TSVと9シートの入力用Excelを出力します。
- `generate_sales_report.py`: 入力用Excelを読み直し、集計Excel・PDF・日次要約を作成します。
- `prompts/implementation.md`: AIエージェントへ実装を依頼するときのプロンプト例。
- `requirements.txt`: Pythonライブラリの依存一覧。
- `output/tsv/`: Excelに貼り付けられる店舗・日別のTSV 9ファイル。
- `output/売上データ.xlsx`: 集計前の売上明細。店舗・日別に9シートあります。
- `output/売上報告.xlsx`: 全体サマリーと店舗・日別シートを含む集計結果。
- `output/日次売上レポート.pdf`: Excelから読み直した内容をまとめた、1日1ページ・全3ページのレポート。
- `output/上司向け日報.txt`: 店舗別売上と全店集計をまとめた、1日3行程度の報告文。
- `output/images/`: 記事掲載用の明細・集計画像。

## 実行方法

Python 3.8以降を用意し、このフォルダで仮想環境を作成します。

```bash
python -m venv .venv
```

仮想環境を有効にします。

```bash
# Windows PowerShell
.\.venv\Scripts\Activate.ps1

# macOS / Linux
source .venv/bin/activate
```

依存パッケージをインストールします。

```bash
python -m pip install -r requirements.txt
```

次の順に実行します。

```bash
python create_sales_data.py
python generate_sales_report.py
```

1つ目でTSVと入力用Excelを `output/` に作成します。必要に応じてExcelを確認・編集した後、2つ目を実行すると、保存済みExcelの全シートを再読込して集計し、Excel・PDF・日次要約を作成します。Excelを編集して再集計する場合は、データ作成スクリプトを再実行せず、`generate_sales_report.py` だけ実行してください。

## データの前提

- 3店舗、5品目、3日分。各店舗・各日につき10人の顧客。
- 顧客1人あたり1〜3種類の商品を購入。品目数量は1〜4個。
- 乱数シードを固定しており、生成結果を再現できます。
- 顧客IDは架空IDです。実際の顧客情報をこのサンプルへ入れないでください。
- Pythonと必要パッケージは [requirements.txt](requirements.txt) を参照してください。