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
