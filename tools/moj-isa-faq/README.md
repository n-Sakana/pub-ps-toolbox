# 法務省 出入国在留管理庁 FAQ Excel Scraper

法務省 出入国在留管理庁のFAQインデックスページから、配下のFAQページを自動検出し、Q&AをExcelに出力するPython CLIです。

対象URLは `URL.txt` に置きます。通常はこのディレクトリでコマンドを実行してください。

## ファイル構成

```text
tools/moj-isa-faq/
  URL.txt          # FAQインデックスURL。通常は1行だけ
  qa_scraper.py    # スクレイピング本体
  requirements.txt # 依存ライブラリ
  test_e2e.py      # 実HPとExcel内容を突き合わせるE2Eテスト
  README.md        # この説明
```

## セットアップ

macOS / Linux / WSL の場合。

```bash
cd tools/moj-isa-faq
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell の場合。

```powershell
cd tools\moj-isa-faq
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

依存ライブラリは `requests`、`beautifulsoup4`、`pandas`、`openpyxl` です。

## 使い方

既定のファイル名で出力する場合。

```bash
python qa_scraper.py
```

この場合、同じディレクトリに `moj_isa_faq.xlsx` が作成されます。

出力先を指定する場合。

```bash
python qa_scraper.py --output output/moj_isa_faq.xlsx
```

別のURLファイルを使う場合。

```bash
python qa_scraper.py --url-file URL.txt --output moj_isa_faq.xlsx
```

URLをコマンドラインで直接指定する場合。

```bash
python qa_scraper.py --index-url "https://www.moj.go.jp/isa/applications/faq/qa_index.html"
```

## URL.txt

`URL.txt` は、FAQインデックスURLを1行だけ書く形式です。

```text
https://www.moj.go.jp/isa/applications/faq/qa_index.html
```

空行と `#` で始まるコメント行は無視します。有効なURLが0件、または2件以上ある場合はエラーにします。曖昧なまま走らせると、Excelが静かに濁ります。紅茶なら廃棄です。

## 出力されるExcel

Excelには3シートを出力します。

- `QA`: Q&A本文。質問番号、質問、回答、元ページURLなど
- `Pages`: インデックス配下で検出したFAQページと件数
- `Summary`: 生成日時、インデックスURL、ページ数、Q&A件数

`QA` シートの主な列は以下です。

- `page_no`
- `faq_page_title`
- `category`
- `section`
- `question_no`
- `question`
- `answer`
- `faq_page_url`
- `answer_page_url`

## 主なオプション

```bash
python qa_scraper.py --help
```

よく使うものだけ挙げます。

- `--output`: Excelの出力先。既定値は `moj_isa_faq.xlsx`
- `--url-file`: FAQインデックスURLを書いたテキストファイル
- `--index-url`: URLファイルを使わず、URLを直接指定
- `--timeout`: HTTPタイムアウト秒数。既定値は30
- `--sleep`: ページ取得間隔秒数。既定値は0.2
- `--expected-page-count`: 検出すべきFAQページ数。既定値は8
- `--min-total-qa`: 最低Q&A件数。下回るとエラー

## E2Eテスト

実際の法務省サイトへアクセスし、Excel生成後に中身を検証します。

```bash
python test_e2e.py
```

このテストは以下を確認します。

- `URL.txt` のインデックスURLを読む
- 実HP上でFAQページが8件検出される
- `qa_scraper.py` でExcelを生成できる
- Excelに `QA`、`Pages`、`Summary` シートがある
- `Pages` シートのページ一覧が、実HPのインデックス配下8ページと一致する
- `QA` シートの各行について、質問本文と回答本文が実HTML上に存在する

ネットワークに接続できない環境、または法務省サイトのHTML構造が変わった場合は失敗します。サイレントデグレードはしません。

## 実行例

```text
Wrote 574 Q&A rows from 8 FAQ pages to moj_isa_faq.xlsx
- 出入国審査・在留審査Q&A: 84 rows
- 在留管理制度Q&A: 149 rows
- 特別永住者制度Q&A: 61 rows
- 所属機関等に関する届出・所属機関による届出Q&A: 106 rows
- 退去強制手続と出国命令制度Q&A: 40 rows
- 監理措置制度Q&A: 45 rows
- 育成就労制度・特定技能制度Q&A: 72 rows
- 永住許可制度の適正化Q&A: 17 rows
```

件数は法務省サイト側の更新で変わる可能性があります。実データを相手にしているので、茶葉の量が昨日と同じとは限りません。
