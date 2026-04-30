# MOJ ISA Site Crawler

出入国在留管理庁サイト `https://www.moj.go.jp/isa/` 配下をクロールし、HTMLページ構造とPDFリンク一覧をExcelに出力する試作ツールです。

最初の狙いは「PDFを全部拾う」ですが、後で全ページ本文の構造化に寄せられるように、取得、解析、出力を分けています。

## 依存関係

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

使っている依存は `requests`、`beautifulsoup4`、`pandas`、`openpyxl`、`networkx`、`matplotlib`、`seaborn`、`numpy` です。PDF本文の抽出やOCRはまだ入れていません。

Graphviz系の図解も使えます。DOTファイルは依存なしで必ず出し、Graphviz本体とPythonバインディングが入っていればGraphvizレンダリング版PNGも追加で出します。

```bash
python -m pip install -r requirements-graphviz.txt
```

`pygraphviz` は環境によってGraphviz本体のインストールが先に必要です。入っていない場合でも黙って薄い紅茶にはせず、`GRAPHVIZ_RENDER_SKIPPED` をログに出して、通常の `networkx` / `matplotlib` PNGとDOTファイルを残します。

## 基本実行

```bash
python crawler.py
```

既定ではPDFファイル本体は保存せず、HTMLページ、リンク、PDF参照の一覧を `moj_isa_crawl.xlsx` に出します。CLIにはページ単位の進捗が出て、詳細ログは `logs/moj_isa_crawler.log`、エラーログは `logs/moj_isa_errors.txt` に保存されます。

ページ取得やPDF処理で一部失敗しても、既定では最後まで走って一覧、統計、エラーログを作ります。止めたい場合だけ `--strict` を付けます。行政サイトの古いリンクで全体を止めると、166ページで「全部見た気」になるので、少々危険です。

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
python crawler.py --graph-dir out/graphs
python crawler.py --no-graphs
python crawler.py --strict
```

`--probe-pdfs` はPDFを保存せず、HEADで `Content-Type`、`Content-Length`、`Last-Modified` を確認します。サーバ側がHEADを拒否する場合はエラーとして記録します。

`--download-pdfs` 実行時は、PDFごとに `PDF_DOWNLOAD_START`、`PDF_RESPONSE`、`PDF_DOWNLOAD_PROGRESS`、`PDF_PART_VERIFY`、`PDF_FINAL_VERIFY`、`PDF_DOWNLOAD_OK` がCLIとログファイルに出ます。保存できなかったPDFは `PDF_ERROR` としてCLIに出し、Excelの `Errors` シートと `logs/moj_isa_errors.txt` にも残します。

PDF保存は `.part` に一時保存してからリネームします。保存直後にstatと先頭バイトの再読み込みも行うので、EDRや権限で保存が潰れた場合は、Pythonから見える範囲で `errno`、`winerror`、traceback、HTTPレスポンスヘッダ、保存後サイズをログに残します。紅茶で言えば、こぼれたかどうかだけでなく、カップが存在したかまで見る方式です。

PDF保存が途中で止まる場合は、まず少数で確認します。

```bash
python crawler.py --download-pdfs --max-pages 15 --max-pdf-downloads 3
```

全件取得時に一部PDFで失敗しても既定で最後まで一覧を作ります。途中で止めたい検証時だけ、次のようにします。

```bash
python crawler.py --download-pdfs --strict
```

失敗は黙殺しません。CLI、`logs/moj_isa_crawler.log`、`logs/moj_isa_errors.txt`、Excelの `Errors` シートに残ります。

## 入力ファイル

`URL.txt` に起点URLを置きます。既定は次です。

```text
https://www.moj.go.jp/isa/
```

クロール範囲は `config.json` の `allowed_prefixes` で制限しています。現在は `https://www.moj.go.jp/isa/` 配下のみです。外部リンクはExcelに記録しますが、たどりません。

## 出力Excel

`moj_isa_crawl.xlsx` には以下のシートを出します。

- `Pages`: ページURL、タイトル、パンくず、見出しJSON、本文テキスト、表JSON、リンク数、PDF数
- `PDFs`: PDF URL、元ページ、リンク文字列、保存パス、Content-Type、Content-Length、Last-Modified、sha256、HTTPヘッダ、保存後検証結果
- `Links`: 各ページから出ているリンク一覧と分類
- `Errors`: 取得やPDF処理のエラー
- `Summary`: 件数サマリ
- `Stats`: 全体統計
- `DepthStats`: クロール深度別統計
- `SectionStats`: `/isa/` 直下セクション別統計
- `PdfStats`: PDF参照のセクション別統計
- `ErrorStats`: エラー種別別統計
- `TopPages`: PDF数・リンク数が多いページ
- `Graphs`: 出力した図解PNGの一覧

Excelのセル上限があるため、長すぎる本文や表JSONはセル内では切り詰めます。全ページ本文を完全保存する段階ではJSONL出力を足すのが筋です。

Excelに入れられない制御文字は出力時に除去します。古いXML風ページや壊れた文字が混じっても、Excel生成で落ちないようにしています。

## 図解

既定で `graphs/` にPNGを出します。

- `depth_distribution.png`: クロール深度別ページ数
- `section_pages.png`: セクション別ページ数
- `pdfs_by_section.png`: セクション別PDF参照数
- `link_structure.png`: 内部リンク構造。大きすぎる場合は上位ノードに絞ります
- `errors_by_phase.png`: エラーがある場合のみ出力
- `link_structure.dot`: Graphviz用の内部リンク構造DOT
- `section_links.dot`: Graphviz用のセクション間リンクDOT
- `link_structure_graphviz.png`: Graphvizで描いた内部リンク構造。Graphviz系が使える場合のみ
- `section_links_graphviz.png`: Graphvizで描いたセクション間リンク構造。Graphviz系が使える場合のみ
- `graphviz_status.txt`: Graphviz系依存と実行ファイルの検出結果

不要な場合は `--no-graphs` を付けます。

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

公式サイトを実際に取得し、15ページをクロールし、PDFを1件だけ保存します。テストではExcelのページ本文プローブとPDFリンクが実HTML上に存在するか、PDFファイルが実際に保存されたか、ログ、エラーログ、統計シート、図解PNGが出ているかまで確認します。

## 構成

```text
moj-isa-crawler/
  URL.txt
  config.json
  crawler.py
  requirements.txt
  requirements-graphviz.txt
  README.md
  scraper/
    fetcher.py
    parser.py
    exporter.py
    analytics.py
    models.py
  tests/
    test_e2e_live.py
```
