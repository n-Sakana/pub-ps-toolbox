# MOJ ISA Site Crawler

出入国在留管理庁サイト `https://www.moj.go.jp/isa/` 配下をクロールし、HTMLページ構造とPDFリンク一覧をExcelに出力する試作ツールです。

最初の狙いは「PDFを全部拾う」ですが、後で全ページ本文の構造化に寄せられるように、取得、解析、出力を分けています。

## 依存関係

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

使っている依存は `requests`、`beautifulsoup4`、`pandas`、`openpyxl` だけです。PDF本文の抽出やOCRはまだ入れていません。

## 基本実行

```bash
python crawler.py
```

既定ではPDFファイル本体は保存せず、HTMLページ、リンク、PDF参照の一覧を `moj_isa_crawl.xlsx` に出します。CLIにはページ単位の進捗が出て、詳細ログは `logs/moj_isa_crawler.log`、エラーログは `logs/moj_isa_errors.txt` に保存されます。

いきなり全PDFを落とすと床に茶葉を撒くので、試作段階ではメタデータ優先です。PDF保存が必要な場合は、必ず `--download-pdfs` を付けます。

PDFも保存する場合は次です。

```bash
python crawler.py --download-pdfs
```

保存先は既定で `downloads/pdfs/` です。

## よく使うオプション

```bash
python crawler.py --max-pages 30
python crawler.py --max-depth 2
python crawler.py --sleep 0.5
python crawler.py --probe-pdfs
python crawler.py --download-pdfs --max-pdf-downloads 10
python crawler.py --output out/moj_isa_crawl.xlsx
python crawler.py --log-file out/crawl.log --error-log-file out/errors.txt
python crawler.py --progress-every-pages 5 --progress-every-pdfs 20
```

`--probe-pdfs` はPDFを保存せず、HEADで `Content-Type`、`Content-Length`、`Last-Modified` を確認します。サーバ側がHEADを拒否する場合はエラーになります。既定は strict なので、失敗を無視したい検証時だけ `--allow-errors` を付けます。

`--download-pdfs` 実行時は、PDFごとに `PDF_DOWNLOAD_START`、`PDF_DOWNLOAD_PROGRESS`、`PDF_DOWNLOAD_OK` がCLIとログファイルに出ます。保存できなかったPDFは `PDF_ERROR` としてCLIに出し、Excelの `Errors` シートと `logs/moj_isa_errors.txt` にも残します。

PDF保存が途中で止まる場合は、まず少数で確認します。

```bash
python crawler.py --download-pdfs --max-pages 15 --max-pdf-downloads 3
```

全件取得時に一部PDFで失敗しても最後まで一覧を作りたい場合だけ、次のようにします。

```bash
python crawler.py --download-pdfs --allow-errors
```

この場合も失敗は黙殺しません。CLI、`logs/moj_isa_crawler.log`、`logs/moj_isa_errors.txt`、Excelの `Errors` シートに残ります。紅茶で言えば、こぼした量まで記録する方式です。

## 入力ファイル

`URL.txt` に起点URLを置きます。既定は次です。

```text
https://www.moj.go.jp/isa/
```

クロール範囲は `config.json` の `allowed_prefixes` で制限しています。現在は `https://www.moj.go.jp/isa/` 配下のみです。外部リンクはExcelに記録しますが、たどりません。

## 出力Excel

`moj_isa_crawl.xlsx` には以下のシートを出します。

- `Pages`: ページURL、タイトル、パンくず、見出しJSON、本文テキスト、表JSON、リンク数、PDF数
- `PDFs`: PDF URL、元ページ、リンク文字列、保存パス、Content-Type、Content-Length、Last-Modified、sha256
- `Links`: 各ページから出ているリンク一覧と分類
- `Errors`: 取得やPDF処理のエラー
- `Summary`: 件数サマリ

Excelのセル上限があるため、長すぎる本文や表JSONはセル内では切り詰めます。全ページ本文を完全保存する段階ではJSONL出力を足すのが筋です。

## ログ

既定では以下を作ります。

- `logs/moj_isa_crawler.log`: CLIと同じ進捗・結果ログ
- `logs/moj_isa_errors.txt`: エラーだけを読みやすく並べたテキストレポート。エラーが0件でも `No errors.` として作成

進捗の粒度は次で変えられます。

```bash
python crawler.py --progress-every-pages 1 --progress-every-pdfs 1
```

`1` なら毎件出力、`10` なら10件ごと、`0` ならその種別の進捗ログを抑制します。

## E2Eテスト

```bash
python tests/test_e2e_live.py
```

公式サイトを実際に取得し、15ページをクロールし、PDFを1件だけ保存します。テストではExcelのページ本文プローブとPDFリンクが実HTML上に存在するか、PDFファイルが実際に保存されたか、ログとエラーログが出ているかまで確認します。

## 構成

```text
moj-isa-crawler/
  URL.txt
  config.json
  crawler.py
  requirements.txt
  README.md
  scraper/
    fetcher.py
    parser.py
    exporter.py
    models.py
  tests/
    test_e2e_live.py
```
