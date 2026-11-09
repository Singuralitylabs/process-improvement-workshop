第2回: Excel読み取りと集計の自動化
================================

架空の小売店の売上ExcelをPythonで作成し、保存済みのExcelをpandasで
読み直して、列選択・データ整形・グループ集計・Excel出力を体験します。
第1回と異なり、今回はPDFと上司向け日報は作成しません。

初めての方は、次の順に進めてください。
Pythonの準備 → VS Codeの準備 → 練習用フォルダを開く → AIへ実装を依頼
→ ライブラリを導入 → Pythonコードを実行 → Excelの結果を確認

必要なもの: パソコン、インターネット接続、GitHubアカウント。
Excelの結果を見るには、Microsoft ExcelやLibreOffice Calcなどを用意します。
ExcelソフトはPythonコードの実行には必要ありません。
AIの利用には通信が発生します。実データや個人情報は送らないでください。

1. 配布教材の内容
--------------------
create_sales_data.py       架空の売上Excelを作成します。
aggregate_sales.py         保存済みExcelを読み、整形・集計します。
requirements.txt          pandasとopenpyxlの依存一覧です。
prompts/implementation.md  AIエージェントへの実装プロンプト例です。
input/売上データ.xlsx      実行時に作成する入力Excelです。
output/売上集計.xlsx       実行時に作成する集計Excelです。

2. Python 3.14をインストールする
--------------------------------
ここから手順6までは、参加者のパソコンで初めて準備する場合の手順です。
講師の現在のLinux環境で実行する場合は、手順7を使ってください。

Windows:
・公式サイト https://www.python.org/downloads/windows/ を開きます。
・Python 3.14のインストール案内に従ってインストールします。
・インストーラーに「Add Python to PATH」が表示される場合はチェックします。
・インストール後にPowerShellを新しく開き、次を入力してEnterを押します。

    py -3.14 --version

macOS:
・公式サイト https://www.python.org/downloads/macos/ を開きます。
・Python 3.14のmacOS用インストーラーを実行します。
・ターミナルを新しく開き、次を実行します。

    python3.14 --version

Linux:
・OS付属のpython3は置き換えず、ユーザー領域にPython 3.14を追加します。
・uvの公式手順 https://docs.astral.sh/uv/getting-started/installation/ に
  従ってuvをインストールし、ターミナルを開き直して次を実行します。

    uv python install 3.14
    python3.14 --version

いずれも「Python 3.14.x」と表示されれば準備完了です。
「コマンドが見つからない」場合は手順10を確認してください。

3. VS CodeとAIチャットを準備する
--------------------------------
・https://code.visualstudio.com/ から、OSに合ったVS Codeをインストールします。
・VS Codeを起動し、左の「拡張機能」で「Python」を検索します。
・Microsoft提供のPython拡張機能をインストールします。
・Copilotのセットアップ案内、または拡張機能の「GitHub Copilot」から
  GitHub提供のCopilotを有効にし、GitHubアカウントでサインインします。
・チャットを開き、エージェント（Agent）モードを選べることを確認します。

Copilotの利用枠・料金・利用可能なモデルは契約によって異なります。
無料枠には制限があります。利用できない場合は講師に相談してください。
画面の名称や配置はVS Codeのバージョンによって変わる場合があります。

4. 練習用フォルダとPython実行環境を用意する
------------------------------------------
AIで最初から作る方は、配布コードを上書きしないように、パソコン上に
空のフォルダ「session02-practice」を作ります。
VS Codeの「ファイル」→「フォルダーを開く」でそのフォルダを開きます。
フォルダの信頼を尋ねられたら、自分で作成したフォルダか確認してください。

「ターミナル」→「新しいターミナル」を開きます。
以降のコマンドはこのターミナルに1行ずつ入力し、Enterを押してください。
ターミナルがsession02-practiceを開いていることを確認します。

参加者の環境では、他の教材とライブラリが混ざらないように仮想環境を使います。
仮想環境は、このフォルダ専用のPython実行環境です。

Windows（PowerShell）:

    py -3.14 -m venv .venv

macOS:

    python3.14 -m venv .venv

Linux（uvで導入した場合）:

    uv venv --python 3.14 --seed .venv

VS Codeでコマンドパレットを開きます（Windows/LinuxはCtrl+Shift+P、
macOSはCmd+Shift+P）。「Python: Select Interpreter」を選び、
作成した.venvのPythonを指定します。見つからない場合はパスを指定します。

    Windows: .venv\Scripts\python.exe
    macOS/Linux: .venv/bin/python

以下ではactivateを使わず、仮想環境のPythonを直接指定します。
PowerShellのスクリプト実行ポリシーを変更する必要はありません。

5. プロンプトを送信してコードを作る
----------------------------------
配布教材のprompts/implementation.mdを開き、```textと最後の```の間にある
「Pythonとpandas、openpyxlを使い」から末尾までの本文をコピーします。
練習用フォルダを開いたVS Codeのチャットに貼り付け、次の補足を付けます。

    現在開いているsession02-practiceフォルダに実装してください。
    create_sales_data.py、aggregate_sales.py、requirements.txtを作成してください。
    このREADMEや配布教材のファイルは変更しないでください。
    Python 3.14の.venvを用意済みです。検証にはそのPythonを使い、
    プロンプト内のpython3.14コマンドもOSに合った.venvのPythonに置き換えてください。
    仮想環境の作り直しやシステム環境の変更はしないでください。
    依存導入とサンプル実行の前に、実行するコマンドを示してください。

エージェントモードで送信すると、AIがファイルの作成や実行を提案します。
変更対象とコマンドを確認し、必要な操作だけを許可してください。
外部へのデータ送信、不要なファイル削除、管理者権限の操作は許可しません。
AIが回答内にコードを表示しただけでは、ファイルが作成されたとは限りません。
左のエクスプローラーに上記3ファイルがあることを確認してください。
足りない場合は「回答だけでなく、このフォルダにファイルを作成してください」
と追加で依頼します。AIの出力は毎回同じとは限らないため、実行結果も確認します。

AIへの依頼を省略する場合は、配布教材のsession02フォルダを開き、
手順4の仮想環境をそこで作成して、用意済みのコードで手順6へ進めます。

6. ライブラリを導入し、コードを実行する（参加者向け）
----------------------------------------------------
Windows（PowerShell）では、次を順に実行します。

    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe -c "import pandas, openpyxl; print(pandas.__version__, openpyxl.__version__)"
    .\.venv\Scripts\python.exe create_sales_data.py
    .\.venv\Scripts\python.exe aggregate_sales.py

macOS/Linuxでは、次を順に実行します。

    .venv/bin/python -m pip install -r requirements.txt
    .venv/bin/python -c "import pandas, openpyxl; print(pandas.__version__, openpyxl.__version__)"
    .venv/bin/python create_sales_data.py
    .venv/bin/python aggregate_sales.py

AIが既にサンプルを生成している場合は、create_sales_data.pyを再実行せず、
aggregate_sales.pyだけを実行してください。結果の確認は手順8へ進みます。

7. 現在のLinux環境で実行する（講師向け・仮想環境なし）
------------------------------------------------------
この節はPython 3.14を導入済みの講師環境専用です。
参加者は手順6を使い、管理環境の保護解除コマンドを流用しないでください。
システムのpython3とは別のため、必ずpython3.14を指定します。
仮想環境を有効にしている場合は、deactivateで解除します。
講座リポジトリのルートから、次のコマンドで移動します。

    cd session02

依存を導入する場合は、次を実行します（導入済みなら不要です）。
uv管理の保護解除はこの環境に限り、--userでユーザー領域に限定します。
システムPythonやPython本体のライブラリ領域は変更しません。

    python3.14 --version
    python3.14 -m pip install --user --break-system-packages -r requirements.txt

次の順に実行します。

    python3.14 create_sales_data.py
    python3.14 aggregate_sales.py

1つ目でinput/売上データ.xlsxを作り、2つ目でそのExcelを読み直して
output/売上集計.xlsxを保存します。入出力先はスクリプトの場所が基準です。
集計処理は入力Excelを変更しません。

入力Excelを編集して再集計するときは、Excelを保存して閉じてから、
aggregate_sales.pyだけを実行してください。

    python3.14 aggregate_sales.py

入力Excelが既にある場合、データ作成スクリプトは上書きせずに終了します。
編集内容を捨ててサンプルを初期状態に戻す場合だけ、次を実行します。

    python3.14 create_sales_data.py --overwrite

集計Excelは再実行時に上書きされます。必要な結果は事前に別名で保存してください。

参加者が再集計・初期化する場合は、上記python3.14を手順6の
仮想環境のPython（WindowsまたはmacOS/Linuxのパス）に置き換えます。

8. サンプルデータと整形
----------------------
・3店舗: 東京駅前店、横浜店、大宮店
・3日: 2026年11月1日～3日
・5品目: コーヒー豆、緑茶、クッキー、マグカップ、ドリップバッグ
・各店舗・日につき5明細、全45明細です。
・入力シート名は「売上明細」です。
・見出しは「日付、店舗、品目、単価、数量」の5列です。備考列はありません。
・1日目の明細の後に空白行を1行入れています。
・一部の店舗名・品目名に前後の空白を、一部の単価・数量に文字列を使っています。
・生成結果は毎回同じです。実データや個人情報をサンプルへ入れないでください。

集計では空白行を除き、必須5列を選択し、店舗名・品目名の前後の空白を取り除きます。
日付を日単位に統一し、単価・数量を整数に変換して、売上金額を計算します。

入力・出力Excelの書式は、Yu Gothicの11ポイント、行高22、列幅22、
薄い灰色の罫線で統一します。見出しは太字・灰色背景・左寄せです。
明細は白背景・上下中央揃え・折り返しなしにします。
日付は中央揃えでyyyy-mm-dd、店舗・品目は左寄せ、数値列は右寄せで
桁区切り表示にします。練習用の文字列の数値と前後の空白はそのまま残します。
文字列の数値には、Excelが桁区切りを適用しない場合や緑の警告マークを
表示する場合があります。集計後は整数に変換されます。
指定フォントがないパソコンでは代替フォントになり、見え方が異なる場合があります。

    売上金額 = 単価 × 数量

9. 集計結果の確認
-----------------
出力Excelには次の4シートがあります（見出し行を除く行数）。
VS Codeのエクスプローラーでoutputフォルダを確認し、売上集計.xlsxを
ExcelやLibreOffice Calcで開きます。VS Codeの標準機能だけでは
Excelの表は表示できません。OSのファイル管理画面から開いても構いません。

    整形済み明細       45行
    店舗別集計          3行
    日別店舗別集計      9行
    品目別集計          5行

初期サンプルの実行結果は次のとおりです。

    全店:       数量113個、売上合計103,050円
    大宮店:     数量38個、売上36,200円
    東京駅前店: 数量36個、売上32,100円
    横浜店:     数量39個、売上34,750円

店舗別・日別店舗別・品目別の数量と売上金額の合計が、整形済み明細の
合計と一致することを確認してください。入力を編集すると結果も変わります。

10. エラーが出た場合
-------------------
・入力Excelなし: 先にcreate_sales_data.pyを実行してください。
・シートや列の不足: 「売上明細」と必須5列の見出しを確認してください。
・明細なし: 全列が空欄ではない売上明細を入力してください。
・店舗・品目の空欄: 空白だけの値も許可されません。
・日付不正: Excelの日付セル、またはYYYY-MM-DD形式で入力してください。
・単価・数量不正: 0～1,000,000の整数で入力してください。
・読取・書込エラー: Excelを閉じ、ファイルと保存先の権限を確認してください。
・コマンドが見つからない: Pythonの導入後にVS Codeとターミナルを開き直します。
    Windowsはpy -3.14、macOS/Linuxはpython3.14でバージョンを確認します。
    Linuxのuv導入時は、公式手順のPATH設定も確認してください。
・can't open file / requirements.txtなし: ターミナルで開いているフォルダを
    確認します。VS Codeの「ターミナル」→「新しいターミナル」を開き直し、
    作成した3ファイルがあるフォルダで実行してください。
・ModuleNotFoundError: 手順6（講師環境は手順7）の依存導入を実行してください。
・No module named pip: 仮想環境の作成が完了したか確認してください。
    Linuxでは手順4のuv venvに--seedが含まれていることを確認します。
・エディターのインポート警告: 「Python: Select Interpreter」で、依存を
    導入したPythonを選びます。講師環境ではpython3.14の実体を選びます。
・Copilotが使えない: サインイン状態と利用枠を確認し、配布コードで実行する
    方法に切り替えるか、講師に相談してください。
・AIが生成したコードのエラー: エラーメッセージをチャットで示し、修正を
    依頼します。個人情報や認証情報が含まれていないことを確認してください。

不正な明細は黙って除外せず、出力Excelを保存する前にエラーで終了します。
以前の集計Excelが残っている場合、それは今回の結果ではありません。

11. コードの読み方
-----------------
各関数の先頭に入力・出力の説明があります。
main関数には、各処理行の直前に日本語の説明コメントがあります。
aggregate_sales.pyは次の順に処理します。

    load_sales -> clean_sales -> summarize_sales -> write_report

12. 使用するライブラリ
----------------------
importは、ライブラリが提供する機能をコードで使えるようにする命令です。
各import行の直前に、そのライブラリを何に使うかコメントを付けています。

Python標準ライブラリ（Pythonに付属し、追加インストールは不要）:
・argparse: 実行時の--overwriteオプションを読み取ります。
・datetime: datetimeで売上の日付を作り、timedeltaで日数を加算します。
・pathlib: Pathで入出力のパスを組み立て、存在確認やフォルダ作成を行います。
・typing: Dictで「シート名をキー、表を値とする辞書」の型ヒントを記載します。
    型ヒントはコードの理解やエディターの検査を助けるもので、集計処理ではありません。
・zipfile: 壊れたZIPファイルを示すBadZipFile例外を捕まえます。
    .xlsxは内部的にZIP形式なので、破損時のエラー対応に使用します。

追加ライブラリ（requirements.txtからインストール）:
・pandas: Excelの表をDataFrame（行と列を持つ表データ）として読み込み、
    列選択、空白・日付・数値の整形、groupbyによる集計、Excel出力を行います。
    「import pandas as pd」のpdは、コード内でpandasを短く書くための別名です。
・openpyxl: .xlsx形式のExcelを読み書きするライブラリです。
    データ生成ではWorkbookでExcelブックを作成します。集計ではpandasの
    engine="openpyxl"として読込・出力を支え、セル書式も設定します。
    openpyxl.stylesから使う各機能は次のとおりです。
        Alignment: 左・中央・右寄せ、縦位置、折り返しの設定。
        Border: セルの四辺の罫線設定。
        Font: フォント名、文字サイズ、太字の設定。
        PatternFill: セルの背景色の設定。
        Side: 罫線の各辺に使う線の種類と色の設定。

依存導入は手順6（講師環境は手順7）を使ってください。
Excelソフト自体はこれらのライブラリに含まれません。

Pythonコードのライセンスは、リポジトリルートのLICENSE（GNU GPL v3）に従います。