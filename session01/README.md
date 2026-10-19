# 第1回: Pythonで小売店の日次売上を集計

営業本部の集計担当者を想定し、3店舗・3日分の売上Excelを読み込んで、店舗別レポートと上司向け要約を作成します。サンプルはすべて架空のデータです。

## 内容

- `generate_sales_report.py`: サンプルデータ作成、Excel読込・集計、レポート出力。主要な処理に日本語コメントを付けています。
- `prompts/implementation.md`: AIエージェントへ実装を依頼するときのプロンプト例。
- `requirements.txt`: Pythonライブラリの依存一覧。
- `output/tsv/`: Excelに貼り付けられる店舗・日別のTSV 9ファイル。
- `output/売上データ.xlsx`: 集計前の売上明細。店舗・日別に9シートあります。
- `output/売上報告.xlsx`: 全体サマリーと店舗・日別シートを含む集計結果。
- `output/日次売上レポート.pdf`: 1日1ページ、全3ページの店舗別レポート。
- `output/上司向け日報.txt`: 1日1行、全3行の報告文。
- `output/images/`: 記事掲載用の明細・集計画像。

## 実行方法

Python 3.8以降を用意し、このフォルダで以下を実行します。

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

依存パッケージをインストールし、サンプルデータとレポートを生成します。

```bash
python -m pip install -r requirements.txt
python generate_sales_report.py --generate-sample
python generate_sales_report.py
```

1つ目の実行でTSVと入力用Excelを `output/` に作り、2つ目の実行でExcelを読み込んでレポートを作成します。Excelを編集して再集計する場合は、データ生成コマンドを再実行せず、2つ目のコマンドだけ実行してください。

## データの前提

- 3店舗、5品目、3日分。各店舗・各日につき10人の顧客。
- 顧客1人あたり1〜3種類の商品を購入。品目数量は1〜4個。
- 乱数シードを固定しており、生成結果を再現できます。
- 顧客IDは架空IDです。実際の顧客情報をこのサンプルへ入れないでください。
- Pythonと必要パッケージは [requirements.txt](requirements.txt) を参照してください。