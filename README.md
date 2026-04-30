# ps-toolbox

PowerShell toolbox host for Explorer context-menu tools.

## Current Tools

- `Print`: print PDF / Word / Excel with a timestamp header
- `Rename`: preview and rename selected files or folders
- `MOJ ISA FAQ Excel Scraper`: scrape ISA/MOJ FAQ pages and export Q&A to Excel
- `MOJ ISA Site Crawler`: crawl ISA/MOJ pages and export page/PDF inventory to Excel

## How It Works

- Launch `launch.vbs` or `launch.bat`
- While the GUI stays open, enabled tools are registered under the Explorer context menu
- Open **Settings** to enable/disable tools and edit per-tool defaults
- Close the GUI to remove the context-menu registration

## Structure

```text
ps-toolbox/
  launch.bat
  launch.vbs
  ps-toolbox.ps1
  config.json
  src/
    01_App.cs
    02_Config.cs
    03_ToolRegistry.cs
    04_ContextMenuManager.cs
    05_HostWindow.cs
    06_SettingsWindow.cs
  tools/
    print/
      tool.json
      run.ps1
    rename/
      tool.json
      run.ps1
    moj-isa-faq/
      URL.txt
      qa_scraper.py
      requirements.txt
      test_e2e.py
      README.md
    moj-isa-crawler/
      URL.txt
      config.json
      crawler.py
      requirements.txt
      scraper/
      tests/
      README.md
```

## MOJ ISA FAQ Excel Scraper

法務省 出入国在留管理庁 FAQ インデックス配下の8ページからQ&Aを取得し、Excelに出力するPython CLIです。

```bash
cd tools/moj-isa-faq
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python qa_scraper.py
python test_e2e.py
```

`URL.txt` にFAQインデックスURLを1行で置き、`qa_scraper.py` が配下8ページを自動検出します。既定の出力ファイルは `moj_isa_faq.xlsx` です。
詳しい使い方は `tools/moj-isa-faq/README.md` を参照してください。

## MOJ ISA Site Crawler

出入国在留管理庁サイト `https://www.moj.go.jp/isa/` 配下をクロールし、HTMLページ構造とPDFリンク一覧をExcelに出力するPython CLIです。

```bash
cd tools/moj-isa-crawler
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python crawler.py
python crawler.py --download-pdfs --max-pdf-downloads 10
python crawler.py --log-file out/crawl.log --error-log-file out/errors.txt
python tests/test_e2e_live.py
```

既定ではPDF本体は保存せず、`moj_isa_crawl.xlsx` に Pages / PDFs / Links / Errors / Summary を出します。進捗はCLIに出し、詳細ログは `logs/moj_isa_crawler.log`、エラーログは `logs/moj_isa_errors.txt` に保存します。詳しい使い方は `tools/moj-isa-crawler/README.md` を参照してください。
