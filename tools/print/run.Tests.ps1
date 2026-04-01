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
    It 'prefixes with &8 for small font' {
        $result = Get-ExcelHeaderText -StampText 'Hello'
        $result | Should Be '&8 Hello'
    }

    It 'escapes ampersands' {
        $result = Get-ExcelHeaderText -StampText 'A&B'
        $result | Should Be '&8 A&&B'
    }

    It 'truncates long text to 220 chars' {
        $long = 'x' * 300
        $result = Get-ExcelHeaderText -StampText $long
        # &8 + space + 220 chars = 223
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
    # Extract arguments from addWatermarkFromText call (greedy to span multiple lines)
    $wmPattern = '(?ms)addWatermarkFromText\((.+)\)'
    $wmMatch = [regex]::Match($scriptContent, $wmPattern)
    $wmArgs = $wmMatch.Groups[1].Value -split ','

    It 'nHorizAlign is 2 (right)' {
        $wmArgs[10].Trim() | Should Be '2'
    }

    It 'nVertAlign is valid (0, 1, or 2)' {
        [int]($wmArgs[12].Trim()) | Should BeLessThan 3
        [int]($wmArgs[12].Trim()) | Should Not BeLessThan 0
    }

    It 'nVertAlign is 0 (top) for header positioning' {
        [int]($wmArgs[12].Trim()) | Should Be 0
    }

    It 'nVertValue is positive (below top edge)' {
        [int]($wmArgs[13].Trim()) | Should BeGreaterThan 0
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
