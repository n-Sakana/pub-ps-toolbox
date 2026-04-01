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
$script:OriginalDevMode = $null
$script:OriginalPrinterName = $null

Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

public static class DuplexHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct PRINTER_DEFAULTS
    {
        public IntPtr pDatatype;
        public IntPtr pDevMode;
        public int DesiredAccess;
    }

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, ref PRINTER_DEFAULTS pDefault);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [DllImport("winspool.drv", SetLastError = true)]
    static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern int DocumentProperties(
        IntPtr hWnd, IntPtr hPrinter, string pDeviceName,
        IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetPrinter(IntPtr hPrinter, int Level, IntPtr pPrinter, int Command);

    const int DM_OUT_BUFFER = 2;
    const int DM_IN_BUFFER  = 8;
    const int DM_FIELDS_OFFSET = 72;
    const int DM_DUPLEX_OFFSET = 94;
    const int DM_DUPLEX_FLAG   = 0x1000;
    const int PRINTER_ACCESS_ADMINISTER = 0x04;
    const int PRINTER_ACCESS_USE        = 0x08;

    static IntPtr OpenWithAccess(string printerName)
    {
        IntPtr hPrinter;
        var defaults = new PRINTER_DEFAULTS
        {
            pDatatype = IntPtr.Zero,
            pDevMode  = IntPtr.Zero,
            DesiredAccess = PRINTER_ACCESS_ADMINISTER
        };
        if (OpenPrinter(printerName, out hPrinter, ref defaults))
            return hPrinter;

        defaults.DesiredAccess = PRINTER_ACCESS_USE;
        if (OpenPrinter(printerName, out hPrinter, ref defaults))
            return hPrinter;

        throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    public static byte[] GetDevMode(string printerName)
    {
        IntPtr hPrinter;
        if (!OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        try
        {
            int size = DocumentProperties(IntPtr.Zero, hPrinter, printerName,
                           IntPtr.Zero, IntPtr.Zero, 0);
            if (size <= 0)
                throw new Exception("DocumentProperties size query failed.");
            IntPtr pDev = Marshal.AllocHGlobal(size);
            try
            {
                if (DocumentProperties(IntPtr.Zero, hPrinter, printerName,
                        pDev, IntPtr.Zero, DM_OUT_BUFFER) < 0)
                    throw new Exception("DocumentProperties get failed.");
                byte[] buf = new byte[size];
                Marshal.Copy(pDev, buf, 0, size);
                return buf;
            }
            finally { Marshal.FreeHGlobal(pDev); }
        }
        finally { ClosePrinter(hPrinter); }
    }

    public static short GetDuplex(byte[] devMode)
    {
        return BitConverter.ToInt16(devMode, DM_DUPLEX_OFFSET);
    }

    static void MergeAndSetDevMode(IntPtr hPrinter, string printerName, byte[] devMode)
    {
        int size = DocumentProperties(IntPtr.Zero, hPrinter, printerName,
                       IntPtr.Zero, IntPtr.Zero, 0);
        if (size <= 0)
            throw new Exception("DocumentProperties size query failed.");

        IntPtr pIn = Marshal.AllocHGlobal(devMode.Length);
        try
        {
            Marshal.Copy(devMode, 0, pIn, devMode.Length);
            IntPtr pOut = Marshal.AllocHGlobal(size);
            try
            {
                if (DocumentProperties(IntPtr.Zero, hPrinter, printerName,
                        pOut, pIn, DM_IN_BUFFER | DM_OUT_BUFFER) < 0)
                    throw new Exception("DocumentProperties merge failed.");

                IntPtr pInfo9 = Marshal.AllocHGlobal(IntPtr.Size);
                try
                {
                    Marshal.WriteIntPtr(pInfo9, pOut);
                    if (!SetPrinter(hPrinter, 9, pInfo9, 0))
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                finally { Marshal.FreeHGlobal(pInfo9); }
            }
            finally { Marshal.FreeHGlobal(pOut); }
        }
        finally { Marshal.FreeHGlobal(pIn); }
    }

    public static void ApplyDuplex(string printerName, byte[] devMode, short duplex)
    {
        byte[] modified = (byte[])devMode.Clone();
        int fields = BitConverter.ToInt32(modified, DM_FIELDS_OFFSET);
        fields |= DM_DUPLEX_FLAG;
        Array.Copy(BitConverter.GetBytes(fields), 0, modified, DM_FIELDS_OFFSET, 4);
        Array.Copy(BitConverter.GetBytes(duplex),  0, modified, DM_DUPLEX_OFFSET, 2);

        IntPtr hPrinter = OpenWithAccess(printerName);
        try
        {
            MergeAndSetDevMode(hPrinter, printerName, modified);
        }
        finally { ClosePrinter(hPrinter); }
    }

    public static void RestoreDevMode(string printerName, byte[] originalDevMode)
    {
        IntPtr hPrinter = OpenWithAccess(printerName);
        try
        {
            MergeAndSetDevMode(hPrinter, printerName, originalDevMode);
        }
        finally { ClosePrinter(hPrinter); }
    }

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "SetDefaultPrinterW")]
    static extern bool NativeSetDefaultPrinter(string printerName);

    public static void SetDefaultPrinter(string printerName)
    {
        if (!NativeSetDefaultPrinter(printerName))
            throw new Win32Exception(Marshal.GetLastWin32Error());
    }
}
'@

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
    return '&9 ' + $safe
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
                $range.ParagraphFormat.Alignment = 2
                $range.Font.Size = 9
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
                $pageSetup.RightHeader = $headerText
                if ($fitExcelWidth) {
                    $pageSetup.Zoom = $false
                    $pageSetup.FitToPagesWide = 1
                    $pageSetup.FitToPagesTall = 32767
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

        $wb.PrintOut()

        Write-Log "Printed via Excel: $Path"
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

function Add-PdfStamp {
    param([string]$PdfPath, [string]$StampText)

    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $raw = [IO.File]::ReadAllText($PdfPath, $latin1)
    $LF = "`n"

    # Parse startxref
    $sxMatch = [regex]::Match($raw, 'startxref\s+(\d+)\s+%%EOF\s*$')
    if (-not $sxMatch.Success) { return $false }
    $prevStartXref = [long]$sxMatch.Groups[1].Value

    # Parse trailer
    $trailerMatch = [regex]::Match($raw, '(?s)trailer\s*<<(.*?)>>\s*startxref\s+\d+\s+%%EOF\s*$')
    if (-not $trailerMatch.Success) { return $false }
    $trailer = $trailerMatch.Groups[1].Value

    $sizeMatch = [regex]::Match($trailer, '/Size\s+(\d+)')
    if (-not $sizeMatch.Success) { return $false }
    $nextObj = [int]$sizeMatch.Groups[1].Value

    $rootMatch = [regex]::Match($trailer, '/Root\s+(\d+\s+\d+\s+R)')
    if (-not $rootMatch.Success) { return $false }
    $rootObjRef = $rootMatch.Groups[1].Value

    $escaped = $StampText.Replace('\', '\\').Replace('(', '\(').Replace(')', '\)')

    # Find page objects using bracket-balanced parsing
    $pageInfos = @()
    foreach ($om in [regex]::Matches($raw, '(\d+)\s+0\s+obj\s*<<')) {
        $objNum = [int]$om.Groups[1].Value
        $dictStart = $om.Index + $om.Length - 2
        $depth = 0; $pos = $dictStart; $dictEnd = -1
        while ($pos -lt $raw.Length - 1) {
            if ($raw[$pos] -eq '<' -and $raw[$pos + 1] -eq '<') { $depth++; $pos += 2 }
            elseif ($raw[$pos] -eq '>' -and $raw[$pos + 1] -eq '>') {
                $depth--
                if ($depth -eq 0) { $dictEnd = $pos + 2; break }
                $pos += 2
            } else { $pos++ }
        }
        if ($dictEnd -eq -1) { continue }
        $fullDict = $raw.Substring($dictStart, $dictEnd - $dictStart)
        if ($fullDict -match '/Type\s*/Page\b' -and $fullDict -notmatch '/Type\s*/Pages\b') {
            $w = 595.28; $h = 841.89
            $mbMatch = [regex]::Match($fullDict, '/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)')
            if ($mbMatch.Success) {
                $w = [double]$mbMatch.Groups[3].Value
                $h = [double]$mbMatch.Groups[4].Value
            }
            # Find existing /Contents reference
            $contRef = ''
            $contMatch = [regex]::Match($fullDict, '/Contents\s+(\d+\s+\d+\s+R)')
            $contArrayMatch = [regex]::Match($fullDict, '/Contents\s*\[([^\]]+)\]')
            if ($contMatch.Success) { $contRef = $contMatch.Groups[1].Value }
            elseif ($contArrayMatch.Success) { $contRef = $contArrayMatch.Groups[1].Value.Trim() }
            $pageInfos += @{ ObjNum = $objNum; FullDict = $fullDict; Width = $w; Height = $h; ContRef = $contRef }
        }
    }
    if ($pageInfos.Count -eq 0) { return $false }

    # Build incremental update
    $ms = New-Object System.IO.MemoryStream
    $baseOffset = (Get-Item -LiteralPath $PdfPath).Length
    $writeStr = { param([string]$s) $b = $latin1.GetBytes($s); $ms.Write($b, 0, $b.Length) }
    $getOffset = { [long]$baseOffset + $ms.Position }
    $xrefKeys = @()
    $xrefOffsets = @{}

    & $writeStr $LF

    # Shared font object (Helvetica)
    $fontObj = $nextObj++
    $xrefKeys += [string]$fontObj
    $xrefOffsets[[string]$fontObj] = (& $getOffset)
    & $writeStr "$fontObj 0 obj$LF<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>$LF endobj$LF"

    foreach ($pi in $pageInfos) {
        $formObj = $nextObj++
        $overlayObj = $nextObj++

        # Text position: right-aligned near top of page
        $estWidth = $StampText.Length * 4.5
        $tx = [int][Math]::Max(0, $pi.Width - 18 - $estWidth)
        $ty = [int]($pi.Height - 20)

        # Form XObject with stamp text (carries its own font resource)
        $formStream = "BT /Helv 9 Tf 0 g $tx $ty Td ($escaped) Tj ET"
        $formLen = $latin1.GetByteCount($formStream)
        $xrefKeys += [string]$formObj
        $xrefOffsets[[string]$formObj] = (& $getOffset)
        & $writeStr "$formObj 0 obj$LF"
        & $writeStr "<< /Type /XObject /Subtype /Form /BBox [0 0 $([int]$pi.Width) $([int]$pi.Height)]$LF"
        & $writeStr "   /Resources << /Font << /Helv $fontObj 0 R >> >> /Length $formLen >>$LF"
        & $writeStr "stream$LF"
        & $writeStr $formStream
        & $writeStr "${LF}endstream${LF}endobj$LF"

        # Overlay content stream: invoke the form XObject
        $overlayStream = "q /PsStamp Do Q"
        $overlayLen = $latin1.GetByteCount($overlayStream)
        $xrefKeys += [string]$overlayObj
        $xrefOffsets[[string]$overlayObj] = (& $getOffset)
        & $writeStr "$overlayObj 0 obj$LF<< /Length $overlayLen >>$LF"
        & $writeStr "stream$LF"
        & $writeStr $overlayStream
        & $writeStr "${LF}endstream${LF}endobj$LF"

        # Rewrite page: add overlay to /Contents, add XObject to /Resources
        $newDict = $pi.FullDict

        # Update /Contents to array including overlay
        if ($pi.ContRef -ne '') {
            $newDict = [regex]::Replace($newDict, '/Contents\s+\d+\s+\d+\s+R', "/Contents [$($pi.ContRef) $overlayObj 0 R]")
            $newDict = [regex]::Replace($newDict, '/Contents\s*\[[^\]]+\]', "/Contents [$($pi.ContRef) $overlayObj 0 R]")
        }

        # Add /XObject to /Resources
        $xoEntry = "/XObject << /PsStamp $formObj 0 R >>"
        if ($newDict -match '/Resources\s*<<') {
            $resPos = [regex]::Match($newDict, '/Resources\s*<<')
            $insertAt = $resPos.Index + $resPos.Length
            $newDict = $newDict.Insert($insertAt, " $xoEntry")
        }

        $xrefKeys += [string]$pi.ObjNum
        $xrefOffsets[[string]$pi.ObjNum] = (& $getOffset)
        & $writeStr "$($pi.ObjNum) 0 obj$LF$newDict${LF}endobj$LF"
    }

    # Xref table
    $xrefOffset = (& $getOffset)
    & $writeStr "xref$LF"
    foreach ($k in ($xrefKeys | Sort-Object { [int]$_ })) {
        & $writeStr "$k 1$LF"
        & $writeStr ("{0:D10} 00000 n {1}" -f $xrefOffsets[$k], $LF)
    }
    & $writeStr "trailer$LF<< /Size $nextObj /Prev $prevStartXref /Root $rootObjRef >>$LF"
    & $writeStr "startxref${LF}${xrefOffset}${LF}%%EOF"

    $appendBytes = $ms.ToArray()
    $ms.Dispose()
    $fs = [IO.File]::Open($PdfPath, [IO.FileMode]::Append)
    try { $fs.Write($appendBytes, 0, $appendBytes.Length) } finally { $fs.Close() }
    return $true
}

function Invoke-PdfPrint {
    param([string]$Path)

    $acroApp = $null
    $avDoc = $null
    $pdDoc = $null
    $tempDir = Join-Path ([IO.Path]::GetTempPath()) ('pstoolbox_' + [guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($tempDir) | Out-Null
    $tempPath = Join-Path $tempDir ([IO.Path]::GetFileName($Path))
    $stampText = Get-StampText -Path $Path

    try {
        Copy-Item -LiteralPath $Path -Destination $tempPath -Force

        try {
            $stamped = Add-PdfStamp -PdfPath $tempPath -StampText $stampText
            if (-not $stamped) {
                Write-Log "PDF stamp skipped (unsupported structure): $Path"
            }
        } catch {
            Write-Log "PDF stamp failed: $($_.Exception.Message)"
        }

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

        # Save to flatten incremental update, then reopen for printing
        $pdDoc = $avDoc.GetPDDoc()
        $saveFlags = $script:AcrobatSaveFull -bor $script:AcrobatSaveCollectGarbage
        $pdDoc.Save($saveFlags, $tempPath) | Out-Null
        $avDoc.Close(1) | Out-Null
        Release-ComObject -ComObject $pdDoc
        Release-ComObject -ComObject $avDoc
        $pdDoc = $null
        $avDoc = $null

        $avDoc = New-Object -ComObject AcroExch.AVDoc
        if (-not $avDoc.Open($tempPath, 'ps-toolbox')) {
            throw 'Acrobat could not reopen the flattened PDF.'
        }
        try { $acroApp.Hide() | Out-Null } catch {}

        $pdDoc = $avDoc.GetPDDoc()
        $pageCount = [int]$pdDoc.GetNumPages()

        $defaultPrinter = (New-Object System.Drawing.Printing.PrinterSettings).PrinterName
        [DuplexHelper]::SetDefaultPrinter($defaultPrinter)
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
        Release-ComObject -ComObject $pdDoc
        Release-ComObject -ComObject $avDoc
        Release-ComObject -ComObject $acroApp
        if (Test-Path -LiteralPath $tempDir) {
            try { Remove-Item -LiteralPath $tempDir -Recurse -Force } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-DuplexValue {
    $setting = Get-ToolSetting -Name 'duplex' -Default 'default'
    switch ($setting) {
        'simplex'    { return [int16]1 }
        'long_edge'  { return [int16]2 }
        'short_edge' { return [int16]3 }
        default      { return [int16]0 }
    }
}

function Set-PrinterDuplex {
    $duplex = Get-DuplexValue
    if ($duplex -eq 0) { return }

    try {
        $printerName = (New-Object System.Drawing.Printing.PrinterSettings).PrinterName
        $script:OriginalDevMode = [DuplexHelper]::GetDevMode($printerName)
        $script:OriginalPrinterName = $printerName
        [DuplexHelper]::ApplyDuplex($printerName, $script:OriginalDevMode, $duplex)
        Write-Log "Duplex set to $duplex on $printerName"
    } catch {
        Write-Log "Duplex override failed: $($_.Exception.Message)"
        $script:OriginalDevMode = $null
        $script:OriginalPrinterName = $null
    }
}

function Restore-PrinterDuplex {
    if ($null -eq $script:OriginalDevMode) { return }

    try {
        [DuplexHelper]::RestoreDevMode($script:OriginalPrinterName, $script:OriginalDevMode)
        Write-Log "Duplex restored on $script:OriginalPrinterName"
    } catch {
        Write-Log "Duplex restore failed: $($_.Exception.Message)"
    }
    $script:OriginalDevMode = $null
    $script:OriginalPrinterName = $null
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

Set-PrinterDuplex
try {
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
} finally {
    Restore-PrinterDuplex
}

if ($errors.Count -gt 0) {
    [System.Windows.Forms.MessageBox]::Show(
        ("Printed {0} file(s), {1} failed.`nLog: {2}" -f $printed, $errors.Count, $script:LogPath),
        'Print',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    exit 1
}

exit 0


