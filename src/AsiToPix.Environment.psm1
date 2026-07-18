$frameFoldersModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.FrameFolders.psm1"
Import-Module $frameFoldersModule -Force

function Test-AsiToPixPathHasCyrillicC {
    param(
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    return $Path.Contains([string][char]0x0421) -or $Path.Contains([string][char]0x0441)
}

function ConvertTo-AsiToPixLatinCPath {
    param(
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    return $Path.Replace([string][char]0x0421, "C").Replace([string][char]0x0441, "c")
}

function Get-AsiToPixCanonicalNumericText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $trimmedValue = $Value.Trim()
    if ($trimmedValue -match '^-?\d+\.0+$') {
        return ($trimmedValue -replace '\.0+$', '')
    }

    return $trimmedValue
}

function Get-AsiToPixPathConventionIssue {
    param(
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }

    $issues = @()
    if (Test-AsiToPixPathHasCyrillicC -Path $Path) {
        $issues += [PSCustomObject]@{
            Kind       = "CyrillicC"
            Message    = "Possible Cyrillic C typo."
            Segment    = $null
            Suggestion = ConvertTo-AsiToPixLatinCPath -Path $Path
        }
    }

    $normalizedPath = ConvertTo-AsiToPixLatinCPath -Path $Path
    $escapedSeparator = '[\\/]'
    $flatFolderPattern = Get-AsiToPixFrameFolderRegex -Kind Flat
    $flatCalibrationPattern = "(?i:Calibration)${escapedSeparator}[^\\/]+${escapedSeparator}(?i:Master|Source)${escapedSeparator}${flatFolderPattern}(?:${escapedSeparator}|$)"
    $isFlatCalibrationPath = $normalizedPath -match $flatCalibrationPattern

    foreach ($segment in ($Path -split '[\\/]')) {
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        $suggestedSegment = $null
        $message = $null
        if ($segment -match '^(?<value>-?\d+(?:\.\d+)?)\s+[Cc]$') {
            $suggestedSegment = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])C"
            $message = "Temperature folder should not contain a space before C."
        } elseif (-not $isFlatCalibrationPath -and
            $segment -match '^(?<value>\d+(?:\.\d+)?)\s*(?<unit>s|sec|secs|second|seconds)$') {
            $suggestedSegment = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])sec"
            $message = "Exposure folder should use the canonical sec suffix."
        } elseif ($segment -match '^(?<value>\d+(?:\.\d+)?)\s*(?<unit>deg|degree|degrees)$') {
            $suggestedSegment = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])deg"
            $message = "Angle token should use the canonical deg suffix without spaces."
        } elseif ($segment -match '(?<token>(?<value>-?\d+(?:\.\d+)?)\s+[Cc])\b') {
            $suggestedToken = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])C"
            $suggestedSegment = $segment.Replace($Matches["token"], $suggestedToken)
            $message = "Temperature token should not contain a space before C."
        } elseif (-not $isFlatCalibrationPath -and
            $segment -match '(?<token>(?<value>\d+(?:\.\d+)?)\s+(?:s|sec|secs|second|seconds))\b') {
            $suggestedToken = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])sec"
            $suggestedSegment = $segment.Replace($Matches["token"], $suggestedToken)
            $message = "Exposure token should use the canonical sec suffix."
        } elseif ($segment -match '(?<token>(?<value>\d+(?:\.\d+)?)\s+(?:deg|degree|degrees))\b') {
            $suggestedToken = "$(Get-AsiToPixCanonicalNumericText -Value $Matches['value'])deg"
            $suggestedSegment = $segment.Replace($Matches["token"], $suggestedToken)
            $message = "Angle token should use the canonical deg suffix without spaces."
        }

        if ($suggestedSegment -and $suggestedSegment -cne $segment) {
            $issues += [PSCustomObject]@{
                Kind       = "FolderToken"
                Message    = $message
                Segment    = $segment
                Suggestion = $Path.Replace($segment, $suggestedSegment)
            }
        }
    }

    $biasFolderPattern = Get-AsiToPixFrameFolderRegex -Kind Bias
    $darkFolderPattern = Get-AsiToPixFrameFolderRegex -Kind Dark
    $flatDarkFolderPattern = Get-AsiToPixFrameFolderRegex -Kind FlatDark
    $calibrationKindPattern = "(?:${biasFolderPattern}|${darkFolderPattern}|${flatDarkFolderPattern})"
    $calibrationPattern = "(?i:Calibration)${escapedSeparator}[^\\/]+${escapedSeparator}(?<mode>(?i:Master|Source))${escapedSeparator}(?<kind>${calibrationKindPattern})${escapedSeparator}(?<next>[^\\/]+)"
    if ($normalizedPath -match $calibrationPattern -and $Matches["next"] -notmatch '^gain\d+$') {
        $mode = $Matches["mode"]
        $kind = $Matches["kind"]
        $next = $Matches["next"]
        $gain = if ($normalizedPath -match '(?i)(?:^|[_\\/])gain(?<gain>\d+)(?:[_\\/]|$)') { $Matches["gain"] } else { "<gain>" }
        $legacyPart = "$mode\$kind\$next"
        $suggestedPart = "$mode\$kind\Gain$gain\$next"
        $issues += [PSCustomObject]@{
            Kind       = "LegacyCalibrationLayout"
            Message    = "Legacy calibration layout: missing gain folder after $mode\$kind."
            Segment    = $legacyPart
            Suggestion = $normalizedPath.Replace($legacyPart, $suggestedPart)
        }
    }

    return @($issues)
}

function Write-AsiToPixPathConventionWarning {
    param(
        [AllowEmptyString()]
        [string]$Path,

        [string]$Context = "Path"
    )

    $issues = @(Get-AsiToPixPathConventionIssue -Path $Path)
    if ($issues.Count -eq 0) {
        return
    }

    if (-not (Get-Variable -Name AsiToPixPathConventionWarnings -Scope Script -ErrorAction SilentlyContinue)) {
        $script:AsiToPixPathConventionWarnings = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    foreach ($issue in $issues) {
        $warningKey = "$Context|$($issue.Kind)|$Path|$($issue.Suggestion)"
        if (-not $script:AsiToPixPathConventionWarnings.Add($warningKey)) {
            continue
        }

        Write-Host "[!] $($issue.Message) ${Context}: $Path" -ForegroundColor Yellow
        if ($issue.Suggestion -and $issue.Suggestion -ne $Path) {
            Write-Host "    Suggested spelling/structure: $($issue.Suggestion)" -ForegroundColor DarkGray
        }
    }
}

function Write-AsiToPixCyrillicPathWarning {
    param(
        [AllowEmptyString()]
        [string]$Path,

        [string]$Context = "Path"
    )

    Write-AsiToPixPathConventionWarning -Path $Path -Context $Context
}

function Test-AsiToPixAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-AsiToPixExecutionPolicyHelp {
    $effectivePolicy = Get-ExecutionPolicy
    if ($effectivePolicy -notin @("Restricted", "AllSigned")) {
        return
    }

    Write-Host "`n[INFO] If PowerShell blocks .ps1 files on this machine, run one of these commands:" -ForegroundColor Cyan
    Write-Host "  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned" -ForegroundColor Yellow
    Write-Host "  powershell -NoProfile -ExecutionPolicy Bypass -File .\CreateProject.ps1" -ForegroundColor Yellow
    Write-Host "You can also run Enable-PowerShellScripts.cmd once from this folder." -ForegroundColor Gray
}

function Get-AsiToPixSymlinkEvaluation {
    $result = @{
        L2L = $false
        L2R = $false
        R2L = $false
        R2R = $false
    }

    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem"
    try {
        $registryValues = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        $registryMap = @{
            L2L = "SymlinkLocalToLocalEvaluation"
            L2R = "SymlinkLocalToRemoteEvaluation"
            R2L = "SymlinkRemoteToLocalEvaluation"
            R2R = "SymlinkRemoteToRemoteEvaluation"
        }

        $foundAllRegistryValues = $true
        foreach ($rule in $registryMap.Keys) {
            $propertyName = $registryMap[$rule]
            if ($null -eq $registryValues.PSObject.Properties[$propertyName]) {
                $foundAllRegistryValues = $false
                break
            }

            $result[$rule] = ([int]$registryValues.$propertyName -eq 1)
        }

        if ($foundAllRegistryValues) {
            return $result
        }
    } catch {
        Write-Verbose "Could not read SymlinkEvaluation registry values from '$registryPath'. Falling back to fsutil output parsing."
    }

    if (-not (Get-Command fsutil -ErrorAction SilentlyContinue)) {
        return $result
    }

    $output = & fsutil behavior query SymlinkEvaluation 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $output) {
        return $result
    }

    foreach ($line in $output) {
        if ($line -match 'Local-to-local.*ENABLED') { $result.L2L = $true }
        elseif ($line -match 'Local-to-remote.*ENABLED') { $result.L2R = $true }
        elseif ($line -match 'Remote-to-local.*ENABLED') { $result.R2L = $true }
        elseif ($line -match 'Remote-to-remote.*ENABLED') { $result.R2R = $true }
    }

    return $result
}

function Enable-AsiToPixSymlinkEvaluation {
    $rules = @("L2L", "L2R", "R2L", "R2R")

    if (-not (Get-Command fsutil -ErrorAction SilentlyContinue)) {
        Write-Host "[!] fsutil is not available; cannot configure symlink evaluation automatically." -ForegroundColor Yellow
        return
    }

    $current = Get-AsiToPixSymlinkEvaluation
    $missingRules = @($rules | Where-Object { -not $current[$_] })

    if ($missingRules.Count -eq 0) {
        Write-Host "[INFO] Symlink evaluation is already enabled for local and network paths." -ForegroundColor DarkGray
        return
    }

    if (-not (Test-AsiToPixAdministrator)) {
        Write-Host "`n[!] Not running as Administrator. Missing symlink evaluation rules were not changed." -ForegroundColor Yellow
        Write-Host "Run these commands in an elevated PowerShell if source or project folders are on a network share:" -ForegroundColor Gray
        foreach ($rule in $missingRules) {
            Write-Host "  fsutil behavior set SymlinkEvaluation $rule`:1" -ForegroundColor Yellow
        }
        return
    }

    foreach ($rule in $missingRules) {
        $output = & fsutil behavior set SymlinkEvaluation "$rule`:1" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[INFO] SymlinkEvaluation $rule enabled." -ForegroundColor DarkGray
        } else {
            Write-Host "[!] Failed to enable SymlinkEvaluation $rule." -ForegroundColor Yellow
            Write-Host "    $output" -ForegroundColor DarkGray
        }
    }
}

function Initialize-AsiToPixEnvironment {
    Write-AsiToPixExecutionPolicyHelp
    Enable-AsiToPixSymlinkEvaluation
}
