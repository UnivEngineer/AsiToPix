$hdrModulePath = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.TotalityHDR.psm1"
$integrationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.TotalityIntegration.psm1"
Import-Module $hdrModulePath -Force
Import-Module $integrationModulePath -Force

function Test-AsiToPixTotalityAlignmentSourcePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceDirectory
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $fullSourceDirectory = [System.IO.Path]::GetFullPath($SourceDirectory).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    if ($fullPath.Equals($fullSourceDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    $sourcePrefix = $fullSourceDirectory + [System.IO.Path]::DirectorySeparatorChar
    return $fullPath.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-AsiToPixTotalityAlignmentFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TotalityPath
    )

    if (-not (Test-Path -LiteralPath $TotalityPath -PathType Container)) {
        throw "Totality source directory not found: '$TotalityPath'."
    }
    $resolvedTotalityPath = (Resolve-Path -LiteralPath $TotalityPath -ErrorAction Stop).ProviderPath
    $directoryPattern = [regex]::new(
        "^Total-(?<exposure>\d+(?:\.\d+)?(?:us|ms|s))$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $filePattern = [regex]::new(
        "^Block_(?<block>\d+)_(?<exposure>\d+(?:\.\d+)?(?:us|ms|s))_(?<phase>.+)_(?<first>\d+)_(?<last>\d+)\.xisf$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $frames = [System.Collections.Generic.List[object]]::new()
    $frameByIdentity = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($exposureDirectory in @(Get-ChildItem -LiteralPath $resolvedTotalityPath -Directory -ErrorAction Stop)) {
        $directoryMatch = $directoryPattern.Match($exposureDirectory.Name)
        if (-not $directoryMatch.Success) {
            continue
        }

        $folderExposure = ConvertTo-AsiToPixTotalityHdrExposure `
            -Text $directoryMatch.Groups["exposure"].Value
        $integratedPath = Join-Path -Path $exposureDirectory.FullName -ChildPath "Frames\integrated"
        if (-not (Test-Path -LiteralPath $integratedPath -PathType Container)) {
            continue
        }
        $resolvedIntegratedPath = (Resolve-Path -LiteralPath $integratedPath -ErrorAction Stop).ProviderPath

        foreach ($file in @(Get-ChildItem -LiteralPath $resolvedIntegratedPath -File -ErrorAction Stop)) {
            $fileMatch = $filePattern.Match($file.Name)
            if (-not $fileMatch.Success) {
                continue
            }

            $fileExposure = ConvertTo-AsiToPixTotalityHdrExposure `
                -Text $fileMatch.Groups["exposure"].Value
            if ($fileExposure.Key -ne $folderExposure.Key) {
                throw "Totality exposure mismatch: folder '$($exposureDirectory.FullName)' is $($folderExposure.Label), but file '$($file.FullName)' is $($fileExposure.Label)."
            }

            $blockNumber = 0
            if (-not [int]::TryParse(
                    $fileMatch.Groups["block"].Value,
                    [System.Globalization.NumberStyles]::None,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [ref]$blockNumber)) {
                throw "Totality block number is outside the supported Int32 range: '$($file.FullName)'."
            }

            $identity = "{0}|{1}" -f $blockNumber, $fileExposure.Key
            if ($frameByIdentity.ContainsKey($identity)) {
                $previous = $frameByIdentity[$identity]
                throw "Duplicate integrated Totality frame for block $blockNumber and exposure $($fileExposure.Label): '$($previous.Path)' and '$($file.FullName)'."
            }

            $frame = [PSCustomObject]@{
                BlockNumber          = $blockNumber
                BlockText            = $fileMatch.Groups["block"].Value
                ExposureLabel        = $fileExposure.Label
                ExposureKey          = $fileExposure.Key
                ExposureMilliseconds = $fileExposure.Milliseconds
                Phase                = $fileMatch.Groups["phase"].Value
                FirstIndex           = $fileMatch.Groups["first"].Value
                LastIndex            = $fileMatch.Groups["last"].Value
                Name                 = $file.Name
                Path                 = $file.FullName
                IntegratedPath       = $resolvedIntegratedPath
            }
            $frameByIdentity.Add($identity, $frame)
            $frames.Add($frame)
        }
    }

    if ($frames.Count -eq 0) {
        throw "No Block_<number>_<exposure>_<phase>_<first>_<last>.xisf files were found under Total-<exposure>\\Frames\\integrated in '$resolvedTotalityPath'."
    }

    return @($frames | Sort-Object -Property BlockNumber, ExposureMilliseconds, Name)
}

function Get-AsiToPixTotalityAlignmentPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Frames,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory,

        [ValidateSet(0, 90, 180, 270)]
        [int]$Rotation = 0,

        [ValidateSet("None", "Horizontal", "Vertical")]
        [string]$Flip = "None"
    )

    if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
        $OutputDirectory = Join-Path -Path $PWD.Path -ChildPath $OutputDirectory
    }
    $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)

    $plan = foreach ($frame in @($Frames | Sort-Object -Property BlockNumber, ExposureMilliseconds, Name)) {
        $suffix = "_ca"
        if ($Rotation -ne 0) {
            $suffix += "_r$Rotation"
        }
        if ($Flip -eq "Horizontal") {
            $suffix += "_fh"
        }
        elseif ($Flip -eq "Vertical") {
            $suffix += "_fv"
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension([string]$frame.Name)
        [PSCustomObject]@{
            BlockNumber          = [int]$frame.BlockNumber
            ExposureLabel        = [string]$frame.ExposureLabel
            ExposureMilliseconds = [decimal]$frame.ExposureMilliseconds
            Phase                = [string]$frame.Phase
            InputPath            = [string]$frame.Path
            OutputPath           = Join-Path -Path $fullOutputDirectory -ChildPath "${baseName}${suffix}.xisf"
            Rotation             = $Rotation
            Flip                 = $Flip
            CanAlign             = $true
            SkipReason           = $null
        }
    }

    return @($plan)
}

function Get-AsiToPixTotalityAlignmentReferenceFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Frames
    )

    $exposures = @($Frames |
        Group-Object -Property ExposureKey |
        ForEach-Object { @($_.Group | Sort-Object -Property ExposureMilliseconds, Name)[0] } |
        Sort-Object -Property ExposureMilliseconds, ExposureLabel)
    $middleExposureIndex = [int][Math]::Floor(($exposures.Count - 1) / 2)
    $middleExposure = $exposures[$middleExposureIndex]
    $candidates = @($Frames |
        Where-Object { $_.ExposureKey -eq $middleExposure.ExposureKey } |
        Sort-Object -Property BlockNumber, Name)
    $middleFrameIndex = [int][Math]::Floor(($candidates.Count - 1) / 2)
    return $candidates[$middleFrameIndex]
}

function Read-AsiToPixTotalityAlignmentRotation {
    [CmdletBinding()]
    param(
        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    Write-Host "Rotation after ChannelMatch:" -ForegroundColor Cyan
    Write-Host "  [0] None (default)"
    Write-Host "  [1] 90 degrees clockwise"
    Write-Host "  [2] 180 degrees"
    Write-Host "  [3] 270 degrees clockwise (90 degrees counter-clockwise)"
    while ($true) {
        $answer = ([string](& $InputReader "Rotation [0]")).Trim()
        switch ($answer.ToLowerInvariant()) {
            "" { return 0 }
            "0" { return 0 }
            "none" { return 0 }
            "1" { return 90 }
            "90" { return 90 }
            "2" { return 180 }
            "180" { return 180 }
            "3" { return 270 }
            "270" { return 270 }
        }
        Write-Host "Enter 0/None, 1/90, 2/180, or 3/270." -ForegroundColor Yellow
    }
}

function Read-AsiToPixTotalityAlignmentFlip {
    [CmdletBinding()]
    param(
        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    Write-Host "Mirror after rotation:" -ForegroundColor Cyan
    Write-Host "  [0] None (default)"
    Write-Host "  [1] Horizontal mirror"
    Write-Host "  [2] Vertical mirror"
    while ($true) {
        $answer = ([string](& $InputReader "Mirror [0]")).Trim()
        switch ($answer.ToLowerInvariant()) {
            "" { return "None" }
            "0" { return "None" }
            "none" { return "None" }
            "1" { return "Horizontal" }
            "horizontal" { return "Horizontal" }
            "h" { return "Horizontal" }
            "2" { return "Vertical" }
            "vertical" { return "Vertical" }
            "v" { return "Vertical" }
        }
        Write-Host "Enter 0/None, 1/Horizontal, or 2/Vertical." -ForegroundColor Yellow
    }
}

function Read-AsiToPixTotalityChannelMatchSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Totality ChannelMatch settings file not found: '$Path'."
    }
    try {
        $settings = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Cannot parse Totality ChannelMatch settings '$Path': $($_.Exception.Message)"
    }

    if ($null -eq $settings -or $settings -isnot [PSCustomObject] -or
        $settings.schemaVersion -ne 1) {
        throw "Unsupported or invalid Totality ChannelMatch settings file: '$Path'."
    }
    $channels = @($settings.channels)
    if ($channels.Count -ne 3) {
        throw "Totality ChannelMatch settings must contain exactly three RGB channels: '$Path'."
    }

    $validatedChannels = [System.Collections.Generic.List[object]]::new()
    $expectedNames = @("R", "G", "B")
    for ($index = 0; $index -lt 3; $index++) {
        $channel = $channels[$index]
        if ($null -eq $channel -or $channel -isnot [PSCustomObject]) {
            throw "Invalid $($expectedNames[$index]) channel in Totality ChannelMatch settings: '$Path'."
        }
        $propertyNames = @($channel.PSObject.Properties.Name)
        foreach ($requiredProperty in @("name", "enabled", "dx", "dy")) {
            if ($propertyNames -notcontains $requiredProperty) {
                throw "Missing '$requiredProperty' in the $($expectedNames[$index]) channel of Totality ChannelMatch settings: '$Path'."
            }
        }
        if ([string]$channel.name -cne $expectedNames[$index]) {
            throw "Unexpected RGB channel order in Totality ChannelMatch settings: expected $($expectedNames[$index]), got '$($channel.name)' in '$Path'."
        }
        if ($channel.enabled -isnot [bool]) {
            throw "The $($expectedNames[$index]) enabled flag must be Boolean in Totality ChannelMatch settings: '$Path'."
        }
        $dx = [double]$channel.dx
        $dy = [double]$channel.dy
        if ([double]::IsNaN($dx) -or [double]::IsInfinity($dx) -or
            [double]::IsNaN($dy) -or [double]::IsInfinity($dy)) {
            throw "Non-finite $($expectedNames[$index]) offset in Totality ChannelMatch settings: '$Path'."
        }
        $validatedChannels.Add([PSCustomObject]@{
            Name    = $expectedNames[$index]
            Enabled = [bool]$channel.enabled
            Dx      = $dx
            Dy      = $dy
        })
    }

    return [PSCustomObject]@{
        SchemaVersion = 1
        SourceIconId  = [string]$settings.sourceIconId
        CapturedAtUtc = [string]$settings.capturedAtUtc
        Channels      = $validatedChannels.ToArray()
        Path          = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    }
}

function Invoke-AsiToPixTotalityAlignmentIpc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Manifest,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightScriptPath,

        [ValidateRange(1, 720)]
        [int]$TimeoutMinutes = 10
    )

    $runningProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
    if ($runningProcesses.Count -ne 1) {
        throw "Totality alignment requires exactly one running PixInsight instance; found $($runningProcesses.Count)."
    }
    $targetPixInsightProcessId = [int]$runningProcesses[0].Id
    $temporaryDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath (
        "AsiToPix-TotalityAlignment-{0}" -f [guid]::NewGuid().ToString("N")
    )
    $manifestPath = Join-Path -Path $temporaryDirectory -ChildPath "AlignmentPlan.json"
    $statusPath = Join-Path -Path $temporaryDirectory -ChildPath "AlignmentStatus.json"
    $ipcScriptPath = Join-Path -Path $temporaryDirectory -ChildPath "AlignTotality-IPC.js"
    $standardOutputPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stdout.txt"
    $standardErrorPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stderr.txt"
    $retainDiagnostics = $false
    $process = $null

    try {
        New-Item -ItemType Directory -Path $temporaryDirectory -ErrorAction Stop | Out-Null
        $manifest.statusPath = $statusPath
        $json = $Manifest | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText(
            $manifestPath,
            $json,
            [System.Text.UTF8Encoding]::new($false)
        )

        $scriptText = [System.IO.File]::ReadAllText(
            $PixInsightScriptPath,
            [System.Text.Encoding]::UTF8
        )
        $ipcMarker = "// ASITOPIX_IPC_MANIFEST"
        if (-not $scriptText.Contains($ipcMarker)) {
            throw "PixInsight Totality alignment script has no IPC manifest marker: '$PixInsightScriptPath'."
        }
        $manifestPathLiteral = ConvertTo-Json -InputObject ([string]$manifestPath) -Compress
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
            -ManifestPath $manifestPath `
            -Mode Reuse `
            -IpcScriptPath $ipcScriptPath
        $process = Start-Process `
            -FilePath $PixInsightPath `
            -ArgumentList $arguments `
            -WindowStyle Hidden `
            -RedirectStandardOutput $standardOutputPath `
            -RedirectStandardError $standardErrorPath `
            -PassThru `
            -ErrorAction Stop

        $lastStatus = $null
        $lastReportedInputPath = ""
        $deadline = [DateTime]::UtcNow.AddMinutes($TimeoutMinutes)
        while ($true) {
            if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
                try {
                    $progressStatus = Get-Content -LiteralPath $statusPath -Raw -Encoding UTF8 |
                        ConvertFrom-Json -ErrorAction Stop
                    $lastStatus = $progressStatus
                    if ($progressStatus.state -eq "running" -and
                        -not [string]::IsNullOrWhiteSpace([string]$progressStatus.currentInputPath) -and
                        $progressStatus.currentInputPath -ne $lastReportedInputPath) {
                        Write-Host "Aligning: $($progressStatus.currentInputPath)" -ForegroundColor Cyan
                        $lastReportedInputPath = [string]$progressStatus.currentInputPath
                    }
                }
                catch {
                    # The status file can be observed between truncate and rewrite.
                    Write-Debug "Alignment status is temporarily unreadable: $($_.Exception.Message)"
                }
            }

            if ($null -ne $lastStatus -and $lastStatus.state -in @("completed", "failed")) {
                break
            }
            if (@(Get-Process -Id $targetPixInsightProcessId -ErrorAction SilentlyContinue).Count -eq 0) {
                throw "The PixInsight instance selected for Totality alignment (PID $targetPixInsightProcessId) closed before the operation completed."
            }
            $process.Refresh()
            if ($process.HasExited -and $process.ExitCode -ne 0 -and $null -eq $lastStatus) {
                throw "PixInsight IPC client returned exit code $($process.ExitCode) before the Totality alignment script started."
            }
            if ([DateTime]::UtcNow -ge $deadline) {
                throw "Timed out after $TimeoutMinutes minutes while waiting for Totality alignment in PixInsight."
            }
            Start-Sleep -Milliseconds 500
        }

        if (-not $process.HasExited) {
            $process.WaitForExit()
        }
        if ($lastStatus.state -ne "completed") {
            $errorText = if ([string]::IsNullOrWhiteSpace([string]$lastStatus.errorMessage)) {
                "PixInsight Totality alignment ended in state '$($lastStatus.state)'."
            }
            else {
                [string]$lastStatus.errorMessage
            }
            throw $errorText
        }
        return $lastStatus
    }
    catch {
        $retainDiagnostics = $true
        throw "PixInsight Totality alignment failed: $($_.Exception.Message) Diagnostics retained at '$temporaryDirectory'."
    }
    finally {
        if (-not $retainDiagnostics) {
            foreach ($temporaryFile in @(
                    $manifestPath,
                    $statusPath,
                    $ipcScriptPath,
                    $standardOutputPath,
                    $standardErrorPath)) {
                if (Test-Path -LiteralPath $temporaryFile -PathType Leaf) {
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
}

function Get-AsiToPixChannelMatchIconSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightScriptPath
    )

    $manifest = [PSCustomObject][ordered]@{
        schemaVersion = 1
        operation     = "inspect"
        statusPath    = ""
    }
    $status = Invoke-AsiToPixTotalityAlignmentIpc `
        -Manifest $manifest `
        -PixInsightPath $PixInsightPath `
        -PixInsightScriptPath $PixInsightScriptPath `
        -TimeoutMinutes 2
    return [string[]]@($status.channelMatchIcons)
}

function Invoke-AsiToPixTotalityAlignmentPlan {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Plan,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PixInsightScriptPath,

        [Parameter(Mandatory)]
        [ValidateSet("Saved", "Capture")]
        [string]$ChannelMatchMode,

        [AllowEmptyCollection()]
        [object[]]$Channels = @(),

        [AllowEmptyCollection()]
        [string[]]$BaselineIconIds = @(),

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SettingsPath
    )

    if (-not (Test-Path -LiteralPath $PixInsightPath -PathType Leaf)) {
        throw "PixInsight executable not found: '$PixInsightPath'."
    }
    if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
        throw "PixInsight Totality alignment script not found: '$PixInsightScriptPath'."
    }

    $approvedFrames = [System.Collections.Generic.List[object]]::new()
    $existingCount = 0
    $notApprovedCount = 0
    foreach ($frame in @($Plan | Sort-Object -Property BlockNumber, ExposureMilliseconds, InputPath)) {
        if (-not $frame.CanAlign) {
            Write-Warning "Skipping Totality alignment input '$($frame.InputPath)': $($frame.SkipReason)."
            continue
        }
        if (-not (Test-Path -LiteralPath $frame.InputPath -PathType Leaf)) {
            throw "Totality alignment input frame not found: '$($frame.InputPath)'."
        }
        if (Test-Path -LiteralPath $frame.OutputPath) {
            if (-not (Test-Path -LiteralPath $frame.OutputPath -PathType Leaf)) {
                throw "Totality alignment output path exists but is not a file: '$($frame.OutputPath)'."
            }
            Write-Warning "Totality alignment output already exists and will not be overwritten: '$($frame.OutputPath)'."
            $existingCount++
            continue
        }
        if ($PSCmdlet.ShouldProcess(
                $frame.OutputPath,
                "Apply ChannelMatch and the selected FastRotation transforms")) {
            $approvedFrames.Add($frame)
        }
        else {
            $notApprovedCount++
        }
    }

    if ($approvedFrames.Count -eq 0) {
        return [PSCustomObject]@{
            AlignedCount     = 0
            ExistingCount    = $existingCount
            NotApprovedCount = $notApprovedCount
            ChannelMatch     = $null
        }
    }

    if ($ChannelMatchMode -eq "Saved" -and @($Channels).Count -ne 3) {
        throw "Exactly three saved RGB ChannelMatch channels are required for Totality alignment."
    }
    if ($ChannelMatchMode -eq "Capture" -and [string]::IsNullOrWhiteSpace($SettingsPath)) {
        throw "SettingsPath is required when capturing ChannelMatch offsets."
    }
    $sourceDirectories = @($Plan |
        ForEach-Object { Split-Path -Path $_.InputPath -Parent } |
        Sort-Object -Unique)
    foreach ($sourceDirectory in $sourceDirectories) {
        foreach ($frame in $approvedFrames) {
            if (Test-AsiToPixTotalityAlignmentSourcePath `
                    -Path $frame.OutputPath `
                    -SourceDirectory $sourceDirectory) {
                throw "Totality alignment output must not be written into a read-only source directory: '$($frame.OutputPath)'. Source: '$sourceDirectory'."
            }
        }
        if ($ChannelMatchMode -eq "Capture" -and
            (Test-AsiToPixTotalityAlignmentSourcePath `
                -Path $SettingsPath `
                -SourceDirectory $sourceDirectory)) {
            throw "Totality ChannelMatch settings must not be written into a read-only source directory: '$SettingsPath'. Source: '$sourceDirectory'."
        }
    }

    $runningProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
    if ($runningProcesses.Count -ne 1) {
        throw "Totality alignment requires exactly one running PixInsight instance; found $($runningProcesses.Count)."
    }

    $outputDirectories = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($frame in $approvedFrames) {
        $null = $outputDirectories.Add((Split-Path -Path $frame.OutputPath -Parent))
    }
    if ($ChannelMatchMode -eq "Capture") {
        $settingsParent = Split-Path -Path $SettingsPath -Parent
        if ([string]::IsNullOrWhiteSpace($settingsParent)) {
            throw "Totality ChannelMatch settings path has no parent directory: '$SettingsPath'."
        }
        $null = $outputDirectories.Add($settingsParent)
    }
    foreach ($outputDirectory in $outputDirectories) {
        if (Test-Path -LiteralPath $outputDirectory) {
            if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
                throw "Totality alignment output path is not a directory: '$outputDirectory'."
            }
        }
        else {
            New-Item -ItemType Directory -Path $outputDirectory -Force -ErrorAction Stop | Out-Null
        }
    }

    $manifestChannels = @($Channels | ForEach-Object {
        [ordered]@{
            name    = [string]$_.Name
            enabled = [bool]$_.Enabled
            dx      = [double]$_.Dx
            dy      = [double]$_.Dy
        }
    })
    $manifestFrames = @($approvedFrames | ForEach-Object {
        [ordered]@{
            blockNumber  = [int]$_.BlockNumber
            exposureLabel = [string]$_.ExposureLabel
            inputPath    = [string]$_.InputPath
            outputPath   = [string]$_.OutputPath
        }
    })
    $firstFrame = $approvedFrames[0]
    $manifest = [PSCustomObject][ordered]@{
        schemaVersion      = 1
        operation          = "align"
        statusPath         = ""
        channelMatchSource = $ChannelMatchMode.ToLowerInvariant()
        channels           = $manifestChannels
        baselineIconIds    = [string[]]@($BaselineIconIds)
        settingsPath       = if ($ChannelMatchMode -eq "Capture") {
            [System.IO.Path]::GetFullPath($SettingsPath)
        }
        else {
            ""
        }
        transform          = [ordered]@{
            rotation = [int]$firstFrame.Rotation
            flip     = [string]$firstFrame.Flip
        }
        frames             = $manifestFrames
    }

    $status = Invoke-AsiToPixTotalityAlignmentIpc `
        -Manifest $manifest `
        -PixInsightPath $PixInsightPath `
        -PixInsightScriptPath $PixInsightScriptPath `
        -TimeoutMinutes 720
    foreach ($frame in $approvedFrames) {
        if (-not (Test-Path -LiteralPath $frame.OutputPath -PathType Leaf)) {
            throw "PixInsight reported completed Totality alignment but output was not created: '$($frame.OutputPath)'."
        }
    }

    return [PSCustomObject]@{
        AlignedCount     = $approvedFrames.Count
        ExistingCount    = $existingCount
        NotApprovedCount = $notApprovedCount
        ChannelMatch     = $status.channelMatch
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixTotalityAlignmentFrame, `
    Get-AsiToPixTotalityAlignmentPlan, `
    Get-AsiToPixTotalityAlignmentReferenceFrame, `
    Read-AsiToPixTotalityAlignmentRotation, `
    Read-AsiToPixTotalityAlignmentFlip, `
    Read-AsiToPixTotalityChannelMatchSetting, `
    Get-AsiToPixChannelMatchIconSnapshot, `
    Invoke-AsiToPixTotalityAlignmentPlan
