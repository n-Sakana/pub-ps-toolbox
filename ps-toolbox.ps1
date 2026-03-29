param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Arguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-ToolScript {
    param(
        [string]$ToolId,
        [string[]]$ToolPaths
    )

    $toolScript = Join-Path $PSScriptRoot ("tools\{0}\run.ps1" -f $ToolId)
    if (-not (Test-Path -LiteralPath $toolScript -PathType Leaf)) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Tool script not found: $ToolId",
            'ps-toolbox',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }

    try {
        & $toolScript -ConfigPath (Join-Path $PSScriptRoot 'config.json') -ToolId $ToolId -Paths $ToolPaths
        if ($LASTEXITCODE -is [int]) {
            exit $LASTEXITCODE
        }
        exit 0
    } catch {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'ps-toolbox',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }
}

if (@($Arguments).Count -ge 2 -and $Arguments[0] -eq '--invoke') {
    $toolId = $Arguments[1]
    $toolPaths = if (@($Arguments).Count -gt 2) { $Arguments[2..(@($Arguments).Count - 1)] } else { @() }
    Invoke-ToolScript -ToolId $toolId -ToolPaths $toolPaths
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

$combined = (Get-ChildItem "$PSScriptRoot\src\*.cs" | Sort-Object Name |
    ForEach-Object { Get-Content $_ -Raw }) -join "`n"

$pat = '(?m)^\s*using\s+[\w][\w.]*\s*;'
$usings = [regex]::Matches($combined, $pat) |
    ForEach-Object { $_.Value.Trim() } | Sort-Object -Unique
$body = $combined -replace $pat, ''
$source = ($usings -join "`n") + "`n`n" + $body

$refs = @(
    [System.Windows.Window].Assembly.Location
    [System.Windows.UIElement].Assembly.Location
    [System.Windows.DependencyObject].Assembly.Location
    [System.Xaml.XamlReader].Assembly.Location
    'Microsoft.CSharp'
    'System.Web.Extensions'
)

Add-Type -TypeDefinition $source -ReferencedAssemblies $refs
[PsToolbox.App]::Run($PSScriptRoot)

