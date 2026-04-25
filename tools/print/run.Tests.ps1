$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptContent = Get-Content -Path "$here/run.ps1" -Raw

# Extract function definitions only (skip the param block and main execution logic)
$funcPattern = '(?ms)(function\s+[\w-]+\s*\{.+?\n\})'
$funcMatches = [regex]::Matches($scriptContent, $funcPattern)
$funcBlock = ($funcMatches | ForEach-Object { $_.Value }) -join "`n`n"

# Set up script-scope variables needed by the functions
$script:Config = @{}
$script:ToolId = 'print'
$script:SupportedExtensions = @(
    '.pdf', '.doc', '.docx', '.docm', '.dot', '.dotx',
    '.xls', '.xlsx', '.xlsm', '.xlsb', '.xlt', '.xltx', '.xltm'
)

# Define the functions in this scope
Invoke-Expression $funcBlock

function New-TestPdf {
    param(
        [string]$Path,
        [double]$Width = 595.28,
        [double]$Height = 841.89
    )

    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $lf = "`n"
    $widthText = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:0.##}', $Width)
    $heightText = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:0.##}', $Height)
    $objects = @(
        "1 0 obj$lf<< /Type /Catalog /Pages 2 0 R >>$lf" + 'endobj' + $lf,
        "2 0 obj$lf<< /Type /Pages /Kids [3 0 R] /Count 1 >>$lf" + 'endobj' + $lf,
        "3 0 obj$lf<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $widthText $heightText] /Resources << >> /Contents 4 0 R >>$lf" + 'endobj' + $lf,
        "4 0 obj$lf<< /Length 0 >>$lf" + 'stream' + $lf + $lf + 'endstream' + $lf + 'endobj' + $lf
    )

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append("%PDF-1.4$lf")
    $offsets = New-Object System.Collections.Generic.List[int]
    foreach ($objectText in $objects) {
        $offsets.Add($latin1.GetByteCount($builder.ToString()))
        [void]$builder.Append($objectText)
    }

    $xrefOffset = $latin1.GetByteCount($builder.ToString())
    [void]$builder.Append("xref$lf")
    [void]$builder.Append("0 5$lf")
    [void]$builder.Append("0000000000 65535 f $lf")
    foreach ($offset in $offsets) {
        [void]$builder.Append(('{0:D10} 00000 n {1}' -f $offset, $lf))
    }
    [void]$builder.Append("trailer$lf<< /Size 5 /Root 1 0 R >>$lf")
    [void]$builder.Append("startxref$lf$xrefOffset$lf" + '%%EOF')

    [IO.File]::WriteAllText($Path, $builder.ToString(), $latin1)
}

function Read-Latin1Text {
    param([string]$Path)

    return [IO.File]::ReadAllText($Path, [System.Text.Encoding]::GetEncoding(28591))
}

function Get-StampedPdfResult {
    param(
        [string]$StampText,
        [double]$Width = 595.28,
        [double]$Height = 841.89
    )

    $pdfPath = Join-Path ([IO.Path]::GetTempPath()) ('ps-toolbox-test-' + [guid]::NewGuid().ToString('N') + '.pdf')
    try {
        New-TestPdf -Path $pdfPath -Width $Width -Height $Height
        return [pscustomobject]@{
            Result = Add-PdfStamp -PdfPath $pdfPath -StampText $StampText
            Raw = Read-Latin1Text -Path $pdfPath
            Width = $Width
            Height = $Height
        }
    } finally {
        Remove-Item -LiteralPath $pdfPath -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Get-StampText' {
    It 'replaces {timestamp} and {filename} tokens' {
        $result = Get-StampText -Path 'C:\docs\report.pdf'
        $result | Should Match '\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}'
        $result | Should Match 'report\.pdf'
    }

    It 'replaces {name} token (without extension)' {
        $script:Config = @{ 'tool.print.header_format' = '{name}' }
        $result = Get-StampText -Path 'C:\docs\report.pdf'
        $result | Should Be 'report'
        $script:Config = @{}
    }
}

Describe 'Get-ExcelHeaderText' {
    It 'prefixes with &9 for small font' {
        $result = Get-ExcelHeaderText -StampText 'Hello'
        $result | Should Be '&9 Hello'
    }

    It 'escapes ampersands' {
        $result = Get-ExcelHeaderText -StampText 'A&B'
        $result | Should Be '&9 A&&B'
    }

    It 'truncates long text to 220 chars' {
        $long = 'x' * 300
        $result = Get-ExcelHeaderText -StampText $long
        # &9 + space + 220 chars = 223
        $result.Length | Should Be 223
    }
}

Describe 'Header alignment - Word' {
    It 'uses right alignment (wdAlignParagraphRight = 2)' {
        $scriptContent | Should Match 'ParagraphFormat\.Alignment\s*=\s*2'
    }
}

Describe 'Header alignment - Excel' {
    It 'uses RightHeader' {
        $scriptContent | Should Match '\$pageSetup\.RightHeader\s*='
    }

    It 'does not use CenterHeader' {
        $scriptContent | Should Not Match '\$pageSetup\.CenterHeader\s*='
    }
}

Describe 'PDF watermark parameters' {
    It 'stamps a minimal PDF end-to-end' {
        $stamped = Get-StampedPdfResult -StampText 'STAMP'
        $stamped.Result | Should Be $true
        $stamped.Raw | Should Match '/PsStamp Do'
        $stamped.Raw | Should Match '/Resources\s*<<\s*/XObject\s*<<\s*/PsStamp\s+\d+\s+0\s+R\s*>>'
    }

    It 'uses the expected right-aligned X position' {
        $stampText = 'STAMP'
        $stamped = Get-StampedPdfResult -StampText $stampText
        $streamMatch = [regex]::Match($stamped.Raw, 'BT /Helv 9 Tf 0 g (\d+) (\d+) Td \(STAMP\) Tj ET')

        $streamMatch.Success | Should Be $true
        [int]$streamMatch.Groups[1].Value | Should Be ([int][Math]::Max(0, $stamped.Width - 18 - ($stampText.Length * 4.5)))
    }

    It 'uses the expected header Y position below the top edge' {
        $stamped = Get-StampedPdfResult -StampText 'STAMP'
        $streamMatch = [regex]::Match($stamped.Raw, 'BT /Helv 9 Tf 0 g (\d+) (\d+) Td \(STAMP\) Tj ET')

        $streamMatch.Success | Should Be $true
        [int]$streamMatch.Groups[2].Value | Should Be ([int]($stamped.Height - 20))
        [int]$streamMatch.Groups[2].Value | Should BeGreaterThan 0
    }

    It 'escapes PDF control characters in stamp text' {
        $stamped = Get-StampedPdfResult -StampText 'A\B(C)'

        $stamped.Result | Should Be $true
        $stamped.Raw.Contains('(A\\B\(C\))') | Should Be $true
    }
}

Describe 'Get-DuplexValue' {
    It 'returns 0 for default setting' {
        $script:Config = @{}
        Get-DuplexValue | Should Be 0
    }

    It 'returns 1 for simplex' {
        $script:Config = @{ 'tool.print.duplex' = 'simplex' }
        Get-DuplexValue | Should Be 1
        $script:Config = @{}
    }

    It 'returns 2 for long_edge' {
        $script:Config = @{ 'tool.print.duplex' = 'long_edge' }
        Get-DuplexValue | Should Be 2
        $script:Config = @{}
    }

    It 'returns 3 for short_edge' {
        $script:Config = @{ 'tool.print.duplex' = 'short_edge' }
        Get-DuplexValue | Should Be 3
        $script:Config = @{}
    }
}

Describe 'Duplex - code structure' {
    It 'calls Set-PrinterDuplex before print loop' {
        $scriptContent | Should Match 'Set-PrinterDuplex'
    }

    It 'calls Restore-PrinterDuplex in finally block' {
        $scriptContent | Should Match 'finally\s*\{\s*\r?\n\s*Restore-PrinterDuplex'
    }

    It 'contains DuplexHelper type with ApplyDuplex method' {
        $scriptContent | Should Match 'public static void ApplyDuplex'
    }

    It 'contains DuplexHelper type with RestoreDevMode method' {
        $scriptContent | Should Match 'public static void RestoreDevMode'
    }

    It 'uses SetPrinter level 9 (per-user defaults)' {
        $scriptContent | Should Match 'SetPrinter\(hPrinter,\s*9,'
    }

    It 'saves printer name for restore' {
        $scriptContent | Should Match '\$script:OriginalPrinterName\s*=\s*\$printerName'
    }

    It 'restores using saved printer name' {
        $scriptContent | Should Match 'RestoreDevMode\(\$script:OriginalPrinterName'
    }

    It 'opens printer with PRINTER_ACCESS_ADMINISTER' {
        $scriptContent | Should Match 'PRINTER_ACCESS_ADMINISTER\s*=\s*0x04'
    }

    It 'RestoreDevMode uses MergeAndSetDevMode for validation' {
        # Extract the RestoreDevMode method body and verify it calls MergeAndSetDevMode
        $restorePattern = '(?ms)public static void RestoreDevMode.+?\n    \}'
        $restoreMatch = [regex]::Match($scriptContent, $restorePattern)
        $restoreMatch.Success | Should Be $true
        $restoreMatch.Value | Should Match 'MergeAndSetDevMode'
    }
}

Describe 'Excel FitToPagesTall' {
    It 'does not use $false for FitToPagesTall' {
        $scriptContent | Should Not Match 'FitToPagesTall\s*=\s*\$false'
    }

    It 'uses a large number for FitToPagesTall' {
        $scriptContent | Should Match 'FitToPagesTall\s*=\s*32767'
    }
}

Describe 'MessageBox string uses double quotes for escape sequences' {
    It 'error MessageBox uses double-quoted string with backtick-n' {
        $scriptContent | Should Match '"Printed \{0\}.*`n.*"'
    }
}

Describe 'Duplex setting in tool.json' {
    $toolJson = Get-Content -Path "$here/tool.json" -Raw | ConvertFrom-Json
    $duplexSetting = $toolJson.settings | Where-Object { $_.key -eq 'duplex' }

    It 'exists in tool.json' {
        $duplexSetting | Should Not BeNullOrEmpty
    }

    It 'is a choice type' {
        $duplexSetting.type | Should Be 'choice'
    }

    It 'defaults to printer default' {
        $duplexSetting.default | Should Be 'default'
    }

    It 'has all four options' {
        $duplexSetting.options.Count | Should Be 4
        $duplexSetting.options | Should Be @('default', 'simplex', 'long_edge', 'short_edge')
    }
}

Describe 'Test-SupportedPath' {
    It 'returns false for empty path' {
        Test-SupportedPath -Path '' | Should Be $false
    }

    It 'returns false for non-existent file' {
        Test-SupportedPath -Path 'C:\nonexistent\file.pdf' | Should Be $false
    }

    It 'returns true for supported extension on existing file' {
        $tmp = [IO.Path]::GetTempFileName()
        $pdfTmp = [IO.Path]::ChangeExtension($tmp, '.pdf')
        Rename-Item -Path $tmp -NewName $pdfTmp
        try {
            Test-SupportedPath -Path $pdfTmp | Should Be $true
        } finally {
            Remove-Item -Path $pdfTmp -Force -ErrorAction SilentlyContinue
        }
    }
}
