param(
    [string]$ConfigPath,
    [string]$ToolId = 'rename',
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

$script:Config = @{}
$script:Items = @()
$script:RenameTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

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

function Get-IntSetting {
    param(
        [string]$Name,
        [int]$Default
    )

    $value = 0
    if ([int]::TryParse((Get-ToolSetting -Name $Name -Default ([string]$Default)), [ref]$value)) {
        return $value
    }
    return $Default
}

function Get-SelectedItems {
    param([string[]]$InputPaths)

    $resolved = foreach ($path in @($InputPaths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $full = (Resolve-Path -LiteralPath $path).ProviderPath
        $isFile = Test-Path -LiteralPath $full -PathType Leaf
        [pscustomobject]@{
            OriginalPath = $full
            Parent       = [IO.Path]::GetDirectoryName($full)
            Name         = [IO.Path]::GetFileName($full)
            IsFile       = $isFile
            BaseName     = if ($isFile) { [IO.Path]::GetFileNameWithoutExtension($full) } else { [IO.Path]::GetFileName($full) }
            Extension    = if ($isFile) { [IO.Path]::GetExtension($full) } else { '' }
        }
    }

    return @($resolved | Sort-Object OriginalPath -Unique)
}

function Test-NestedSelection {
    param([object[]]$Items)

    $itemsArray = @($Items)
    for ($i = 0; $i -lt $itemsArray.Count; $i++) {
        for ($j = 0; $j -lt $itemsArray.Count; $j++) {
            if ($i -eq $j) { continue }
            if ($itemsArray[$j].OriginalPath.StartsWith($itemsArray[$i].OriginalPath + [IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }
    return $false
}

function Get-NumberString {
    param(
        [int]$Index,
        [int]$Start,
        [int]$Padding
    )

    $value = $Start + $Index
    if ($Padding -gt 0) {
        return $value.ToString(('D' + $Padding))
    }
    return $value.ToString()
}

function Expand-NameTokens {
    param(
        [string]$Value,
        [object]$Item
    )

    if ($null -eq $Value) {
        return ''
    }

    $parentFolder = ''
    if ($null -ne $Item -and -not [string]::IsNullOrWhiteSpace($Item.Parent)) {
        $parentFolder = Split-Path -Path $Item.Parent -Leaf
    }

    return $Value.Replace('{timestamp}', $script:RenameTimestamp).Replace('{parent_folder}', $parentFolder)
}

function Join-NameAffix {
    param(
        [string]$BaseName,
        [string]$Affix,
        [string]$Separator,
        [switch]$AsPrefix
    )

    if ([string]::IsNullOrWhiteSpace($Affix)) {
        return $BaseName
    }

    if ([string]::IsNullOrEmpty($Separator)) {
        if ($AsPrefix) {
            return $Affix + $BaseName
        }
        return $BaseName + $Affix
    }

    if ($AsPrefix) {
        return $Affix + $Separator + $BaseName
    }

    return $BaseName + $Separator + $Affix
}

function Get-RenameOptions {
    $affixSeparator = Get-ToolSetting -Name 'affix_separator' -Default '_'
    if ($affixSeparator -eq 'none') {
        $affixSeparator = ''
    }

    return [pscustomobject]@{
        BaseNameOverride  = Get-ToolSetting -Name 'base_name' -Default ''
        PrefixTemplate    = Get-ToolSetting -Name 'prefix' -Default ''
        SuffixTemplate    = Get-ToolSetting -Name 'suffix' -Default '{timestamp}'
        AffixSeparator    = $affixSeparator
        ReplaceFrom       = Get-ToolSetting -Name 'replace_from' -Default ''
        ReplaceTo         = Get-ToolSetting -Name 'replace_to' -Default ''
        NumberingMode     = Get-ToolSetting -Name 'numbering_mode' -Default 'none'
        NumberStart       = Get-IntSetting -Name 'number_start' -Default 1
        NumberPadding     = [Math]::Max((Get-IntSetting -Name 'number_padding' -Default 2), 0)
        NumberSeparator   = Get-ToolSetting -Name 'number_separator' -Default '_'
        PreserveExtension = Get-BoolSetting -Name 'preserve_extension' -Default $true
    }
}
function Get-RenamePlans {
    param([object[]]$Items)

    $itemsArray = @($Items)
    $options = Get-RenameOptions
    $plans = New-Object 'System.Collections.Generic.List[object]'
    $nested = Test-NestedSelection -Items $itemsArray
    $selectedPaths = @{}

    foreach ($item in $itemsArray) {
        $selectedPaths[$item.OriginalPath.ToLowerInvariant()] = $true
    }

    for ($i = 0; $i -lt $itemsArray.Count; $i++) {
        $item = $itemsArray[$i]
        $name = if ([string]::IsNullOrWhiteSpace($options.BaseNameOverride)) { $item.BaseName } else { $options.BaseNameOverride }

        if (-not [string]::IsNullOrEmpty($options.ReplaceFrom)) {
            $name = $name.Replace($options.ReplaceFrom, $options.ReplaceTo)
        }

        $prefixText = Expand-NameTokens -Value $options.PrefixTemplate -Item $item
        $suffixText = Expand-NameTokens -Value $options.SuffixTemplate -Item $item
        $name = Join-NameAffix -BaseName $name -Affix $prefixText -Separator $options.AffixSeparator -AsPrefix
        $name = Join-NameAffix -BaseName $name -Affix $suffixText -Separator $options.AffixSeparator

        if ($options.NumberingMode -ne 'none') {
            $num = Get-NumberString -Index $i -Start $options.NumberStart -Padding $options.NumberPadding
            if ($options.NumberingMode -eq 'prefix') {
                $name = $num + $options.NumberSeparator + $name
            } elseif ($options.NumberingMode -eq 'suffix') {
                $name = $name + $options.NumberSeparator + $num
            }
        }

        $newLeaf = if ($item.IsFile -and $options.PreserveExtension) { $name + $item.Extension } else { $name }
        $newPath = if ([string]::IsNullOrWhiteSpace($newLeaf)) { $item.Parent } else { Join-Path $item.Parent $newLeaf }
        $status = 'OK'

        if ($nested) {
            $status = 'Nested selection unsupported'
        } elseif ([string]::IsNullOrWhiteSpace($newLeaf)) {
            $status = 'Empty name'
        } elseif ($newLeaf.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
            $status = 'Invalid name'
        } elseif ($newLeaf -eq $item.Name) {
            $status = 'Unchanged'
        } elseif ((Test-Path -LiteralPath $newPath) -and -not $selectedPaths.ContainsKey($newPath.ToLowerInvariant())) {
            $status = 'Already exists'
        }

        $plans.Add([pscustomobject]@{
            Item    = $item
            NewLeaf = $newLeaf
            NewPath = $newPath
            Status  = $status
        }) | Out-Null
    }

    $duplicates = @(
        $plans |
            Where-Object { $_.Status -eq 'OK' } |
            Group-Object { $_.NewPath.ToLowerInvariant() } |
            Where-Object { $_.Count -gt 1 }
    )
    foreach ($dup in $duplicates) {
        foreach ($plan in @($dup.Group)) {
            $plan.Status = 'Duplicate target'
        }
    }

    return $plans.ToArray()
}

function Show-BlockedPlans {
    param([object[]]$Plans)

    $blocked = @($Plans | Where-Object { $_.Status -notin @('OK', 'Unchanged') })
    if ($blocked.Count -eq 0) {
        return
    }

    $lines = @()
    foreach ($plan in ($blocked | Select-Object -First 12)) {
        $lines += ('{0}: {1}' -f $plan.Item.Name, $plan.Status)
    }
    if ($blocked.Count -gt 12) {
        $lines += '...'
    }

    [System.Windows.Forms.MessageBox]::Show(
        ("Rename was not applied.`n`n" + ($lines -join "`n")),
        'Rename',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

function Apply-Rename {
    param([object[]]$Plans)

    $changes = @($Plans | Where-Object { $_.Status -eq 'OK' })
    if ($changes.Count -eq 0) {
        return $true
    }

    $tempMappings = New-Object 'System.Collections.Generic.List[object]'
    try {
        foreach ($plan in $changes) {
            $tempLeaf = '.__pstoolbox_' + [guid]::NewGuid().ToString('N')
            if ($plan.Item.IsFile -and $plan.Item.Extension) {
                $tempLeaf += $plan.Item.Extension
            }
            $tempPath = Join-Path $plan.Item.Parent $tempLeaf
            Move-Item -LiteralPath $plan.Item.OriginalPath -Destination $tempPath -ErrorAction Stop
            $tempMappings.Add([pscustomobject]@{
                TempPath     = $tempPath
                FinalPath    = $plan.NewPath
                OriginalPath = $plan.Item.OriginalPath
            }) | Out-Null
        }

        foreach ($map in $tempMappings) {
            Move-Item -LiteralPath $map.TempPath -Destination $map.FinalPath -ErrorAction Stop
        }

        return $true
    } catch {
        foreach ($map in $tempMappings) {
            try {
                if (Test-Path -LiteralPath $map.TempPath) {
                    Move-Item -LiteralPath $map.TempPath -Destination $map.OriginalPath -ErrorAction SilentlyContinue
                }
            } catch {
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'Rename',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

$script:Config = Load-ConfigMap -Path $ConfigPath
$script:Items = @(Get-SelectedItems -InputPaths $Paths)
if ($script:Items.Count -eq 0) {
    exit 0
}

$plans = @(Get-RenamePlans -Items $script:Items)
$blocked = @($plans | Where-Object { $_.Status -notin @('OK', 'Unchanged') })
if ($blocked.Count -gt 0) {
    Show-BlockedPlans -Plans $plans
    exit 1
}

if (Apply-Rename -Plans $plans) {
    exit 0
}

exit 1









