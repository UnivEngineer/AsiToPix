$integrationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.TotalityIntegration.psm1"
Import-Module $integrationModulePath -Force

function Get-AsiToPixTotalityNormalizationFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$InputPath
    )

    if (-not (Test-Path -LiteralPath $InputPath -PathType Container)) {
        throw "Totality normalization input directory not found: '$InputPath'."
    }
    $resolvedInputPath = (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).ProviderPath
    $namePattern = [regex]::new(
        "^Block-(?<block>\d+)-HDR\.xisf$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $frames = [System.Collections.Generic.List[object]]::new()
    $frameByBlock = [System.Collections.Generic.Dictionary[int, object]]::new()

    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedInputPath -File -ErrorAction Stop)) {
        $match = $namePattern.Match($file.Name)
        if (-not $match.Success) {
            continue
        }

        $blockNumber = 0
        if (-not [int]::TryParse(
                $match.Groups["block"].Value,
                [System.Globalization.NumberStyles]::None,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$blockNumber)) {
            throw "Totality normalization block number is outside the supported Int32 range: '$($file.FullName)'."
        }
        if ($frameByBlock.ContainsKey($blockNumber)) {
            throw "Duplicate Totality normalization frame for block ${blockNumber}: '$($frameByBlock[$blockNumber].Path)' and '$($file.FullName)'."
        }

        $frame = [PSCustomObject]@{
            BlockNumber = $blockNumber
            BlockText   = $match.Groups["block"].Value
            Name        = $file.Name
            Path        = $file.FullName
        }
        $frameByBlock.Add($blockNumber, $frame)
        $frames.Add($frame)
    }

    if ($frames.Count -eq 0) {
        throw "No Block-<number>-HDR.xisf frames found in '$resolvedInputPath'."
    }

    return @($frames | Sort-Object -Property BlockNumber, Name)
}

function Get-AsiToPixTotalityNormalizationPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Frames,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory
    )

    if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
        $OutputDirectory = Join-Path -Path $PWD.Path -ChildPath $OutputDirectory
    }
    $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
    $plan = foreach ($frame in @($Frames | Sort-Object -Property BlockNumber, Name)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension([string]$frame.Name)
        [PSCustomObject]@{
            BlockNumber       = [int]$frame.BlockNumber
            InputPath         = [string]$frame.Path
            NormalizedPath    = Join-Path -Path $fullOutputDirectory -ChildPath "${baseName}_norm.xisf"
            ColorCorrectedPath = Join-Path -Path $fullOutputDirectory -ChildPath "${baseName}_norm_cc.xisf"
            CanNormalize      = $true
            SkipReason        = $null
        }
    }
    return @($plan)
}

function New-AsiToPixTotalityNormalizationManifest {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Frames,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$AllowedReferenceFiles,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MetricsPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ManifestDirectory
    )

    if (-not (Test-Path -LiteralPath $ManifestDirectory -PathType Container)) {
        throw "Normalization manifest directory not found: '$ManifestDirectory'."
    }
    if ((Test-Path -LiteralPath $MetricsPath) -and
        -not (Test-Path -LiteralPath $MetricsPath -PathType Leaf)) {
        throw "Normalization metrics path exists but is not a file: '$MetricsPath'."
    }

    $manifestFrames = [System.Collections.Generic.List[object]]::new()
    foreach ($frame in @($Frames | Sort-Object -Property BlockNumber)) {
        if (-not (Test-Path -LiteralPath $frame.InputPath -PathType Leaf)) {
            throw "Normalization input frame not found: '$($frame.InputPath)'."
        }
        foreach ($outputPath in @($frame.NormalizedPath, $frame.ColorCorrectedPath)) {
            if ((Test-Path -LiteralPath $outputPath) -and
                -not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                throw "Normalization output path exists but is not a file: '$outputPath'."
            }
        }

        $manifestFrames.Add([ordered]@{
            blockNumber        = [int]$frame.BlockNumber
            inputPath          = [string]$frame.InputPath
            normalizedPath     = [string]$frame.NormalizedPath
            colorCorrectedPath = [string]$frame.ColorCorrectedPath
        })
    }

    $statusPath = Join-Path -Path $ManifestDirectory -ChildPath "NormalizationStatus.json"
    $manifest = [ordered]@{
        schemaVersion         = 1
        statusPath            = $statusPath
        metricsPath           = [string]$MetricsPath
        allowedReferenceFiles = [string[]]@($AllowedReferenceFiles)
        frames                = @($manifestFrames)
    }
    $manifestPath = Join-Path -Path $ManifestDirectory -ChildPath "NormalizationPlan.json"
    $json = $manifest | ConvertTo-Json -Depth 6
    if ($PSCmdlet.ShouldProcess($manifestPath, "Write PixInsight Totality normalization manifest")) {
        [System.IO.File]::WriteAllText(
            $manifestPath,
            $json,
            [System.Text.UTF8Encoding]::new($false)
        )
    }

    return [PSCustomObject]@{
        ManifestPath = $manifestPath
        StatusPath   = $statusPath
    }
}

function Invoke-AsiToPixTotalityNormalizationPlan {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Plan,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MetricsPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightScriptPath
    )

    if (-not (Test-Path -LiteralPath $PixInsightPath -PathType Leaf)) {
        throw "PixInsight executable not found: '$PixInsightPath'."
    }
    if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
        throw "PixInsight Totality normalization script not found: '$PixInsightScriptPath'."
    }
    if ($Plan.Count -eq 0) {
        throw "Cannot execute an empty Totality normalization plan."
    }
    if ((Test-Path -LiteralPath $MetricsPath) -and
        -not (Test-Path -LiteralPath $MetricsPath -PathType Leaf)) {
        throw "Normalization metrics path exists but is not a file: '$MetricsPath'."
    }

    $approvedFrames = [System.Collections.Generic.List[object]]::new()
    $notApprovedCount = 0
    foreach ($frame in @($Plan | Sort-Object -Property BlockNumber)) {
        if (-not $frame.CanNormalize) {
            Write-Warning "Skipping normalization block $($frame.BlockNumber): $($frame.SkipReason)."
            continue
        }
        if (-not (Test-Path -LiteralPath $frame.InputPath -PathType Leaf)) {
            throw "Normalization input frame not found: '$($frame.InputPath)'."
        }
        foreach ($outputPath in @($frame.NormalizedPath, $frame.ColorCorrectedPath)) {
            if ((Test-Path -LiteralPath $outputPath) -and
                -not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                throw "Normalization output path exists but is not a file: '$outputPath'."
            }
        }

        $description = "Create scalar-normalized and color-corrected Totality frames with PixInsight"
        if ($PSCmdlet.ShouldProcess($frame.ColorCorrectedPath, $description)) {
            $approvedFrames.Add($frame)
        }
        else {
            $notApprovedCount++
        }
    }

    if ($approvedFrames.Count -eq 0) {
        return [PSCustomObject]@{
            NormalizedCount = 0
            NotApprovedCount = $notApprovedCount
        }
    }

    $runningProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
    if ($runningProcesses.Count -ne 1) {
        throw "Totality normalization requires exactly one running PixInsight instance so the selected reference preview is unambiguous; found $($runningProcesses.Count)."
    }
    $targetPixInsightProcessId = [int]$runningProcesses[0].Id

    $outputDirectories = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $null = $outputDirectories.Add((Split-Path -Path $MetricsPath -Parent))
    foreach ($frame in $approvedFrames) {
        $null = $outputDirectories.Add((Split-Path -Path $frame.NormalizedPath -Parent))
        $null = $outputDirectories.Add((Split-Path -Path $frame.ColorCorrectedPath -Parent))
    }
    foreach ($outputDirectory in $outputDirectories) {
        if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
            throw "Totality normalization output path has no parent directory."
        }
        if (Test-Path -LiteralPath $outputDirectory) {
            if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
                throw "Totality normalization output path is not a directory: '$outputDirectory'."
            }
        }
        else {
            New-Item -ItemType Directory -Path $outputDirectory -ErrorAction Stop | Out-Null
        }
    }

    $temporaryDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath (
        "AsiToPix-TotalityNormalization-{0}" -f [guid]::NewGuid().ToString("N")
    )
    $manifestInfo = $null
    $standardOutputPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stdout.txt"
    $standardErrorPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stderr.txt"
    $ipcScriptPath = Join-Path -Path $temporaryDirectory -ChildPath "NormalizeTotality-IPC.js"
    $retainDiagnostics = $false
    $normalizedCount = 0

    try {
        New-Item -ItemType Directory -Path $temporaryDirectory -ErrorAction Stop | Out-Null
        $manifestInfo = New-AsiToPixTotalityNormalizationManifest `
            -Frames $approvedFrames.ToArray() `
            -AllowedReferenceFiles @($Plan.InputPath) `
            -MetricsPath $MetricsPath `
            -ManifestDirectory $temporaryDirectory `
            -Confirm:$false

        $scriptText = [System.IO.File]::ReadAllText(
            $PixInsightScriptPath,
            [System.Text.Encoding]::UTF8
        )
        $ipcMarker = "// ASITOPIX_IPC_MANIFEST"
        if (-not $scriptText.Contains($ipcMarker)) {
            throw "PixInsight Totality normalization script has no IPC manifest marker: '$PixInsightScriptPath'."
        }
        $manifestPathLiteral = ConvertTo-Json -InputObject ([string]$manifestInfo.ManifestPath) -Compress
        $ipcScriptText = $scriptText.Replace(
            $ipcMarker,
            "var ASITOPIX_MANIFEST_PATH = $manifestPathLiteral;"
        )
        [System.IO.File]::WriteAllText(
            $ipcScriptPath,
            $ipcScriptText,
            [System.Text.UTF8Encoding]::new($false)
        )

        $arguments = Get-AsiToPixPixInsightArgumentList `
            -ScriptPath $PixInsightScriptPath `
            -ManifestPath $manifestInfo.ManifestPath `
            -Mode Reuse `
            -IpcScriptPath $ipcScriptPath
        Write-Host "Sending $($approvedFrames.Count) Totality normalization frame(s) to the running PixInsight instance; it will remain open." -ForegroundColor Cyan
        $process = Start-Process `
            -FilePath $PixInsightPath `
            -ArgumentList $arguments `
            -WindowStyle Hidden `
            -RedirectStandardOutput $standardOutputPath `
            -RedirectStandardError $standardErrorPath `
            -PassThru `
            -ErrorAction Stop

        $lastReportedBlockNumber = $null
        $lastStatus = $null
        $deadline = [DateTime]::UtcNow.AddHours(12)
        while ($true) {
            if (Test-Path -LiteralPath $manifestInfo.StatusPath -PathType Leaf) {
                try {
                    $progressStatus = Get-Content -LiteralPath $manifestInfo.StatusPath -Raw -Encoding UTF8 |
                        ConvertFrom-Json -ErrorAction Stop
                    $lastStatus = $progressStatus
                    if ($progressStatus.state -eq "running" -and
                        $null -ne $progressStatus.currentBlockNumber -and
                        $progressStatus.currentBlockNumber -ne $lastReportedBlockNumber) {
                        Write-Host "Normalizing Totality block $($progressStatus.currentBlockNumber)..." -ForegroundColor Cyan
                        $lastReportedBlockNumber = $progressStatus.currentBlockNumber
                    }
                }
                catch {
                    # The status file can be observed between truncate and rewrite.
                    Write-Debug "Normalization status is temporarily unreadable: $($_.Exception.Message)"
                }
            }

            if ($null -ne $lastStatus -and $lastStatus.state -in @("completed", "failed")) {
                break
            }
            $process.Refresh()
            if ($process.HasExited -and $process.ExitCode -ne 0 -and $null -eq $lastStatus) {
                throw "PixInsight IPC client returned exit code $($process.ExitCode) before the normalization script started."
            }
            if (@(Get-Process -Id $targetPixInsightProcessId -ErrorAction SilentlyContinue).Count -eq 0) {
                throw "The PixInsight instance selected for Totality normalization (PID $targetPixInsightProcessId) closed before processing completed."
            }
            if ([DateTime]::UtcNow -ge $deadline) {
                throw "Timed out after 12 hours while waiting for Totality normalization in PixInsight."
            }
            Start-Sleep -Milliseconds 500
        }

        if (-not $process.HasExited) {
            $process.WaitForExit()
        }
        if ($null -eq $lastStatus) {
            throw "PixInsight did not write a Totality normalization status file."
        }
        if ($lastStatus.state -ne "completed") {
            $blockText = if ($null -ne $lastStatus.currentBlockNumber) {
                " Block $($lastStatus.currentBlockNumber)."
            }
            else {
                ""
            }
            $errorText = if ([string]::IsNullOrWhiteSpace([string]$lastStatus.errorMessage)) {
                "PixInsight normalization ended in state '$($lastStatus.state)'."
            }
            else {
                [string]$lastStatus.errorMessage
            }
            throw "$errorText$blockText"
        }

        foreach ($frame in $approvedFrames) {
            foreach ($outputPath in @($frame.NormalizedPath, $frame.ColorCorrectedPath)) {
                if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                    throw "PixInsight reported completion but normalization output was not created: '$outputPath'."
                }
            }
            $normalizedCount++
        }
        if (-not (Test-Path -LiteralPath $MetricsPath -PathType Leaf)) {
            throw "PixInsight reported completion but normalization metrics CSV was not created: '$MetricsPath'."
        }
        if ($null -ne $process.ExitCode -and $process.ExitCode -ne 0) {
            Write-Warning "PixInsight IPC client returned exit code $($process.ExitCode) after completing Totality normalization."
        }
    }
    catch {
        $retainDiagnostics = $true
        throw "PixInsight Totality normalization failed: $($_.Exception.Message) Diagnostics retained at '$temporaryDirectory'."
    }
    finally {
        if (-not $retainDiagnostics) {
            $temporaryFiles = @(
                if ($null -ne $manifestInfo) { $manifestInfo.ManifestPath; $manifestInfo.StatusPath }
                $ipcScriptPath
                $standardOutputPath
                $standardErrorPath
            )
            foreach ($temporaryFile in $temporaryFiles) {
                if (-not [string]::IsNullOrWhiteSpace([string]$temporaryFile) -and
                    (Test-Path -LiteralPath $temporaryFile -PathType Leaf)) {
                    Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
                }
            }
            if (Test-Path -LiteralPath $temporaryDirectory -PathType Container) {
                $remainingItems = @(Get-ChildItem -LiteralPath $temporaryDirectory -Force -ErrorAction SilentlyContinue)
                if ($remainingItems.Count -eq 0) {
                    Remove-Item -LiteralPath $temporaryDirectory -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    return [PSCustomObject]@{
        NormalizedCount = $normalizedCount
        NotApprovedCount = $notApprovedCount
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixTotalityNormalizationFrame, `
    Get-AsiToPixTotalityNormalizationPlan, `
    New-AsiToPixTotalityNormalizationManifest, `
    Invoke-AsiToPixTotalityNormalizationPlan
