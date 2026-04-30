# ps-toolbox

PowerShell toolbox host for Explorer context-menu tools.

## Current Tools

- `Print`: print PDF / Word / Excel with a timestamp header
- `Rename`: preview and rename selected files or folders

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
```

## MOJ ISA FAQ Excel Scraper

法務省 出入国在留管理庁 FAQ インデックス配下の8ページからQ&Aを取得し、Excelに出力するPython CLIです。

```bash
cd tools/moj-isa-faq
python -m pip install -r requirements.txt
python qa_scraper.py
python test_e2e.py
```

`URL.txt` にFAQインデックスURLを1行で置き、`qa_scraper.py` が配下8ページを自動検出します。既定の出力ファイルは `moj_isa_faq.xlsx` です。
