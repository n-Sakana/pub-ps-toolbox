param(
    [string]$ConfigPath,
    [string]$ToolId = 'print',
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:SupportedExtensions = @(
    '.pdf', '.doc', '.docx', '.docm', '.dot', '.dotx',
    '.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.xltm'
)
$script:LogPath = Join-Path $PSScriptRoot 'print.log'
$script:Config = @{}
$script:AcrobatSaveFull = 0x01
$script:AcrobatSaveCollectGarbage = 0x20

function Load-ConfigMap {
    param([string]$Path)

    $map = @{}
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
        try {
            $obj = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) {
                $map[$p.Name] = [string]$p.Value
            }
        } catch {
        }
    }
    return $map
}

function Get-ToolSetting {
    param(
        [string]$Name,
        [string]$Default
    )

    $key = "tool.$ToolId.$Name"
    if ($script:Config.ContainsKey($key)) {
        return $script:Config[$key]
    }
    return $Default
}

function Get-BoolSetting {
    param(
        [string]$Name,
        [bool]$Default
    )

    $raw = Get-ToolSetting -Name $Name -Default ($(if ($Default) { '1' } else { '0' }))
    return $raw -eq '1' -or $raw -eq 'true'
}

function Write-Log {
    param([string]$Message)

    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -eq $ComObject) { return }
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Test-SupportedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    return $script:SupportedExtensions -contains $ext
}

function Get-StampText {
    param([string]$Path)

    $format = Get-ToolSetting -Name 'header_format' -Default 'Printed {timestamp} | {filename}'
    $fileName = [IO.Path]::GetFileName($Path)
    $nameOnly = [IO.Path]::GetFileNameWithoutExtension($Path)
    return $format.Replace('{timestamp}', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')).Replace('{filename}', $fileName).Replace('{name}', $nameOnly)
}

function Get-ExcelHeaderText {
    param([string]$StampText)

    $safe = $StampText.Replace('&', '&&')
    if ($safe.Length -gt 220) {
        $safe = $safe.Substring(0, 220)
    }
    return '&8 ' + $safe
}

function Set-WordHeaders {
    param(
        [object]$Document,
        [string]$StampText
    )

    foreach ($section in @($Document.Sections)) {
        foreach ($headerType in 1, 2, 3) {
            try {
                $header = $section.Headers.Item($headerType)
                $range = $header.Range
                $range.Text = $StampText
                $range.ParagraphFormat.Alignment = 1
                $range.Font.Size = 8
                $range.Font.Name = 'Meiryo'
            } catch {
                Write-Log "Word header update skipped: $($_.Exception.Message)"
            }
        }
    }
}

function Invoke-WordPrint {
    param([string]$Path)

    $word = $null
    $doc = $null
    $stampText = Get-StampText -Path $Path

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        try { $word.Options.PrintBackground = $false } catch {}

        $doc = $word.Documents.Open($Path, $false, $true)
        Set-WordHeaders -Document $doc -StampText $stampText
        $doc.PrintOut($false)
        Write-Log "Printed via Word: $Path"
    } catch {
        throw "Word print failed: $Path`n$($_.Exception.Message)"
    } finally {
        if ($null -ne $doc) { try { $doc.Close(0) } catch {} }
        if ($null -ne $word) { try { $word.Quit() } catch {} }
        Release-ComObject -ComObject $doc
        Release-ComObject -ComObject $word
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-ExcelPrint {
    param([string]$Path)

    $excel = $null
    $wb = $null
    $headerText = Get-ExcelHeaderText -StampText (Get-StampText -Path $Path)
    $fitExcelWidth = Get-BoolSetting -Name 'fit_excel_width' -Default $true
    $visibleSheetsOnly = Get-BoolSetting -Name 'visible_sheets_only' -Default $true
    $autoLandscape = Get-BoolSetting -Name 'auto_landscape' -Default $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path, 0, $true)

        try { $excel.PrintCommunication = $false } catch {}
        foreach ($sheet in @($wb.Worksheets)) {
            try {
                if ($visibleSheetsOnly -and $sheet.Visible -ne -1) { continue }
                $pageSetup = $sheet.PageSetup
                $pageSetup.CenterHeader = $headerText
                if ($fitExcelWidth) {
                    $pageSetup.Zoom = $false
                    $pageSetup.FitToPagesWide = 1
                    $pageSetup.FitToPagesTall = $false
                }
                if ($autoLandscape) {
                    $usedRange = $sheet.UsedRange
                    if ($usedRange.Columns.Count -ge 8) {
                        $pageSetup.Orientation = 2
                    }
                }
                try { $pageSetup.HeaderMargin = $excel.InchesToPoints(0.2) } catch {}
            } catch {
                Write-Log "Excel setup skipped: $Path / $($_.Exception.Message)"
            }
        }
        try { $excel.PrintCommunication = $true } catch {}

        $printedSheetCount = 0
        foreach ($sheet in @($wb.Worksheets)) {
            try {
                if ($visibleSheetsOnly -and $sheet.Visible -ne -1) { continue }
                $sheet.PrintOut()
                $printedSheetCount++
            } catch {
                Write-Log "Excel sheet print skipped: $Path / $($_.Exception.Message)"
            }
        }

        if ($printedSheetCount -eq 0) {
            throw 'No printable worksheet found.'
        }

        Write-Log "Printed via Excel: $Path (sheets=$printedSheetCount)"
    } catch {
        throw "Excel print failed: $Path`n$($_.Exception.Message)"
    } finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObject -ComObject $wb
        Release-ComObject -ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-PdfPrint {
    param([string]$Path)

    $acroApp = $null
    $avDoc = $null
    $pdDoc = $null
    $jsDoc = $null
    $tempPath = Join-Path ([IO.Path]::GetTempPath()) ('pstoolbox_pdf_' + [guid]::NewGuid().ToString('N') + '.pdf')
    $stampText = Get-StampText -Path $Path

    try {
        Copy-Item -LiteralPath $Path -Destination $tempPath -Force

        try {
            $acroApp = New-Object -ComObject AcroExch.App
        } catch {
            if ($_.Exception.HResult -eq -2147221164) {
                throw 'Adobe Acrobat Pro automation is not available on this PC.'
            }
            throw
        }
        try { $acroApp.Hide() | Out-Null } catch {}

        try {
            $avDoc = New-Object -ComObject AcroExch.AVDoc
        } catch {
            if ($_.Exception.HResult -eq -2147221164) {
                throw 'Adobe Acrobat Pro automation is not available on this PC.'
            }
            throw
        }
        if (-not $avDoc.Open($tempPath, 'ps-toolbox')) {
            throw 'Acrobat could not open the temporary PDF.'
        }

        try { $acroApp.Hide() | Out-Null } catch {}

        $pdDoc = $avDoc.GetPDDoc()
        if ($null -eq $pdDoc) {
            throw 'Acrobat PDDoc was not available.'
        }

        $jsDoc = $pdDoc.GetJSObject()
        if ($null -eq $jsDoc) {
            throw 'Acrobat JavaScript object was not available.'
        }

        $pageCount = [int]$pdDoc.GetNumPages()
        if ($pageCount -le 0) {
            throw 'PDF has no printable pages.'
        }

        $colorBlack = @('RGB', 0, 0, 0)
        $null = $jsDoc.addWatermarkFromText(
            $stampText,
            2,
            'Helv',
            8,
            $colorBlack,
            0,
            ($pageCount - 1),
            $true,
            $false,
            $true,
            2,
            3,
            -36,
            -18,
            $false,
            1.0,
            $true,
            0,
            1.0
        )

        $saveFlags = $script:AcrobatSaveFull -bor $script:AcrobatSaveCollectGarbage
        $saveOk = $pdDoc.Save($saveFlags, $tempPath)
        if ($saveOk -ne -1) {
            throw 'Acrobat could not save the stamped temporary PDF.'
        }

        $printOk = $avDoc.PrintPagesSilent(0, ($pageCount - 1), 2, 0, 1)
        if ($printOk -ne -1) {
            throw 'Acrobat silent print failed.'
        }

        Write-Log "Printed via Acrobat: $Path"
    } catch {
        throw "PDF print failed: $Path`n$($_.Exception.Message)"
    } finally {
        if ($null -ne $avDoc) { try { $avDoc.Close(1) | Out-Null } catch {} }
        if ($null -ne $acroApp) {
            try { $acroApp.Hide() | Out-Null } catch {}
            try { $acroApp.Exit() | Out-Null } catch {}
        }
        Release-ComObject -ComObject $jsDoc
        Release-ComObject -ComObject $pdDoc
        Release-ComObject -ComObject $avDoc
        Release-ComObject -ComObject $acroApp
        if (Test-Path -LiteralPath $tempPath) {
            try { Remove-Item -LiteralPath $tempPath -Force } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-PrintFile {
    param([string]$Path)

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($ext) {
        '.pdf' { Invoke-PdfPrint -Path $Path; break }
        '.doc' { Invoke-WordPrint -Path $Path; break }
        '.docx' { Invoke-WordPrint -Path $Path; break }
        '.docm' { Invoke-WordPrint -Path $Path; break }
        '.dot' { Invoke-WordPrint -Path $Path; break }
        '.dotx' { Invoke-WordPrint -Path $Path; break }
        '.xls' { Invoke-ExcelPrint -Path $Path; break }
        '.xlsx' { Invoke-ExcelPrint -Path $Path; break }
        '.xlsm' { Invoke-ExcelPrint -Path $Path; break }
        '.xlsb' { Invoke-ExcelPrint -Path $Path; break }
        '.xlt' { Invoke-ExcelPrint -Path $Path; break }
        '.xltx' { Invoke-ExcelPrint -Path $Path; break }
        '.xltm' { Invoke-ExcelPrint -Path $Path; break }
        default { throw "Unsupported extension: $ext" }
    }
}

$script:Config = Load-ConfigMap -Path $ConfigPath
$targets = @($Paths | Where-Object { Test-SupportedPath -Path $_ } | Select-Object -Unique)
if ($targets.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        'No supported file was selected.',
        'Print',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
    exit 0
}

$printed = 0
$errors = New-Object System.Collections.Generic.List[string]
foreach ($path in $targets) {
    try {
        Invoke-PrintFile -Path $path
        $printed++
    } catch {
        $errors.Add($_.Exception.Message)
        Write-Log $_.Exception.Message
    }
}

if ($errors.Count -gt 0) {
    [System.Windows.Forms.MessageBox]::Show(
        ('Printed {0} file(s), {1} failed.`nLog: {2}' -f $printed, $errors.Count, $script:LogPath),
        'Print',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit 1
}

exit 0


