$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.TotalityAlignment.psm1"
$entryPointPath = Join-Path -Path $PSScriptRoot -ChildPath "..\AlignTotality.ps1"
$pixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\AlignTotality.js"
Import-Module $modulePath -Force

function Add-TestTotalityIntegratedFrame {
    param(
        [Parameter(Mandatory)]
        [string]$TotalityPath,

        [Parameter(Mandatory)]
        [string]$Exposure,

        [Parameter(Mandatory)]
        [int]$Block,

        [string]$Phase = "Totality",

        [string]$FirstIndex = "00102",

        [string]$LastIndex = "00125"
    )

    $integratedPath = Join-Path -Path $TotalityPath -ChildPath (
        "Total-{0}\Frames\integrated" -f $Exposure
    )
    New-Item -ItemType Directory -Path $integratedPath -Force | Out-Null
    $path = Join-Path -Path $integratedPath -ChildPath (
        "Block_{0:D3}_{1}_{2}_{3}_{4}.xisf" -f
        $Block, $Exposure, $Phase, $FirstIndex, $LastIndex
    )
    Set-Content -LiteralPath $path -Value "test"
    return $path
}

Describe "Totality alignment frame discovery" {
    It "discovers every Total-exposure integrated folder and parses frame metadata" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-discovery"
        New-Item -ItemType Directory -Path $totalityPath -Force | Out-Null
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1 `
            -Phase "BeadsC2" -FirstIndex "00002" -LastIndex "00025" | Out-Null
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "100ms" -Block 5 | Out-Null
        New-Item -ItemType Directory -Path (
            Join-Path $totalityPath "Total-1ms\Frames\registered"
        ) -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $totalityPath "notes.txt") -Value "ignore"

        $frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)

        $frames.Count | Should Be 2
        $frames[0].BlockNumber | Should Be 1
        $frames[0].ExposureLabel | Should Be "10ms"
        $frames[0].Phase | Should Be "BeadsC2"
        $frames[0].FirstIndex | Should Be "00002"
        $frames[1].BlockNumber | Should Be 5
        $frames[1].ExposureMilliseconds | Should Be ([decimal]100)
    }

    It "rejects a filename exposure that disagrees with its Total-exposure folder" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-mismatch"
        $integratedPath = Join-Path -Path $totalityPath -ChildPath "Total-10ms\Frames\integrated"
        New-Item -ItemType Directory -Path $integratedPath -Force | Out-Null
        Set-Content -LiteralPath (
            Join-Path $integratedPath "Block_001_100ms_Totality_00001_00025.xisf"
        ) -Value "test"

        $caughtError = $null
        try {
            $null = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should Not BeNullOrEmpty
        $caughtError.Exception.Message | Should Match "exposure mismatch"
    }

    It "rejects duplicate numeric block and exposure identities" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-duplicates"
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1 -Phase "BeadsC2" | Out-Null
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1 -Phase "Totality" | Out-Null

        $caughtError = $null
        try {
            $null = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should Not BeNullOrEmpty
        $caughtError.Exception.Message | Should Match "Duplicate integrated Totality frame"
    }
}

Describe "Totality alignment plan" {
    It "records ChannelMatch, rotation, and mirror operations in output names" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-plan-input"
        $outputPath = Join-Path -Path $TestDrive -ChildPath "alignment-plan-output"
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "100ms" -Block 5 `
            -FirstIndex "00123" -LastIndex "00150" | Out-Null
        $frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)

        $plan = @(Get-AsiToPixTotalityAlignmentPlan `
            -Frames $frames `
            -OutputDirectory $outputPath `
            -Rotation 180 `
            -Flip Horizontal)

        $plan.Count | Should Be 1
        $plan[0].OutputPath | Should Be (
            Join-Path $outputPath "Block_005_100ms_Totality_00123_00150_ca_r180_fh.xisf"
        )
        $plan[0].Rotation | Should Be 180
        $plan[0].Flip | Should Be "Horizontal"
    }

    It "uses only the ca suffix when no FastRotation transform is selected" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-plan-ca"
        $outputPath = Join-Path -Path $TestDrive -ChildPath "alignment-plan-ca-output"
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1 | Out-Null
        $frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)

        $plan = @(Get-AsiToPixTotalityAlignmentPlan `
            -Frames $frames `
            -OutputDirectory $outputPath)

        $plan[0].OutputPath | Should Be (
            Join-Path $outputPath "Block_001_10ms_Totality_00102_00125_ca.xisf"
        )
    }

    It "recommends a central block from the middle discovered exposure" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-reference"
        foreach ($exposure in @("1ms", "10ms", "100ms")) {
            foreach ($block in @(1, 5, 9)) {
                Add-TestTotalityIntegratedFrame `
                    -TotalityPath $totalityPath -Exposure $exposure -Block $block | Out-Null
            }
        }
        $frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)

        $reference = Get-AsiToPixTotalityAlignmentReferenceFrame -Frames $frames

        $reference.ExposureLabel | Should Be "10ms"
        $reference.BlockNumber | Should Be 5
    }
}

Describe "Totality alignment source safety" {
    It "refuses an alignment output directory inside an integrated source folder" {
        $totalityPath = Join-Path -Path $TestDrive -ChildPath "alignment-source-safety"
        $inputPath = Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1
        $sourceDirectory = Split-Path -Path $inputPath -Parent
        $frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $totalityPath)
        $plan = @(Get-AsiToPixTotalityAlignmentPlan `
            -Frames $frames `
            -OutputDirectory $sourceDirectory)
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "source-safety-PixInsight.exe"
        Set-Content -LiteralPath $pixInsightPath -Value "test executable"

        $caughtError = $null
        try {
            $null = Invoke-AsiToPixTotalityAlignmentPlan `
                -Plan $plan `
                -PixInsightPath $pixInsightPath `
                -PixInsightScriptPath $pixInsightScriptPath `
                -ChannelMatchMode Saved `
                -Channels @(
                    [PSCustomObject]@{ Name = "R"; Enabled = $true; Dx = 0; Dy = 0 },
                    [PSCustomObject]@{ Name = "G"; Enabled = $true; Dx = 0; Dy = 0 },
                    [PSCustomObject]@{ Name = "B"; Enabled = $true; Dx = 0; Dy = 0 }
                ) `
                -SettingsPath (Join-Path $TestDrive "unused.json") `
                -Confirm:$false
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should Not BeNullOrEmpty
        $caughtError.Exception.Message | Should Match "read-only source directory"
        Test-Path -LiteralPath $plan[0].OutputPath | Should Be $false
    }
}

Describe "Totality ChannelMatch settings" {
    It "loads three saved RGB offsets" {
        $settingsPath = Join-Path -Path $TestDrive -ChildPath "channel_match_offsets.json"
        $json = @'
{
  "schemaVersion": 1,
  "sourceIconId": "ChannelMatch01",
  "capturedAtUtc": "2026-08-22T12:00:00.000Z",
  "channels": [
    { "name": "R", "enabled": true, "dx": 1.25, "dy": -0.5 },
    { "name": "G", "enabled": true, "dx": 0, "dy": 0 },
    { "name": "B", "enabled": false, "dx": -1.75, "dy": 0.25 }
  ]
}
'@
        Set-Content -LiteralPath $settingsPath -Value $json -Encoding UTF8

        $settings = Read-AsiToPixTotalityChannelMatchSetting -Path $settingsPath

        $settings.SourceIconId | Should Be "ChannelMatch01"
        $settings.Channels.Count | Should Be 3
        $settings.Channels[0].Name | Should Be "R"
        $settings.Channels[0].Dx | Should Be 1.25
        $settings.Channels[2].Enabled | Should Be $false
    }

    It "rejects settings without exactly three channels" {
        $settingsPath = Join-Path -Path $TestDrive -ChildPath "bad_channel_match_offsets.json"
        Set-Content -LiteralPath $settingsPath -Value (
            '{"schemaVersion":1,"channels":[{"enabled":true,"dx":0,"dy":0}]}'
        ) -Encoding UTF8

        $caughtError = $null
        try {
            $null = Read-AsiToPixTotalityChannelMatchSetting -Path $settingsPath
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should Not BeNullOrEmpty
        $caughtError.Exception.Message | Should Match "exactly three RGB channels"
    }
}

Describe "PixInsight Totality alignment script" {
    It "captures ChannelMatch icon offsets and applies ChannelMatch before FastRotation" {
        $scriptText = Get-Content -LiteralPath $pixInsightScriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'ProcessInstance\.iconsByProcessId\( "ChannelMatch" \)'
        $scriptText | Should Match 'ProcessInstance\.fromIcon\( sourceIconId \)'
        $scriptText | Should Match 'process\.channels = channelRows\( channels \);'
        $scriptText | Should Match 'rows\.push\( \[ channels\[i\]\.enabled, channels\[i\]\.dx, channels\[i\]\.dy, 1\.0 \] \);'
        $scriptText | Should Match 'applyChannelMatch\( window\.mainView, channels \);\s+applyTransforms'
        $scriptText | Should Match 'FastRotation\.prototype\.Rotate90CW'
        $scriptText | Should Match 'FastRotation\.prototype\.Rotate180'
        $scriptText | Should Match 'FastRotation\.prototype\.Rotate90CCW'
        $scriptText | Should Match 'FastRotation\.prototype\.HorizontalMirror'
        $scriptText | Should Match 'FastRotation\.prototype\.VerticalMirror'
        $scriptText | Should Match 'window\.saveAs\( frame\.outputPath, false, false, false, false \);'
        $scriptText | Should Match '// ASITOPIX_IPC_MANIFEST'
    }
}

Describe "Totality alignment entry point" {
    It "builds a WhatIf plan without creating the Aligned output directory" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "AstroPhoto"
        $totalityPath = Join-Path -Path $astroPhotoRoot -ChildPath "SharpCap\Totality-2026"
        $outputPath = Join-Path -Path $totalityPath -ChildPath "HDR\Aligned"
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "PixInsight.exe"
        Add-TestTotalityIntegratedFrame `
            -TotalityPath $totalityPath -Exposure "10ms" -Block 1 | Out-Null
        Set-Content -LiteralPath $pixInsightPath -Value "test executable"

        $null = & $entryPointPath `
            -AstroPhotoRoot $astroPhotoRoot `
            -Rotation 180 `
            -Flip None `
            -ChannelMatchMode Capture `
            -PixInsightPath $pixInsightPath `
            -WhatIf `
            -Confirm:$false

        Test-Path -LiteralPath $outputPath | Should Be $false
    }

    It "uses shared root aliases, the expected folders, and a reusable JSON settings file" {
        $scriptText = Get-Content -LiteralPath $entryPointPath -Raw -Encoding UTF8

        $scriptText | Should Match 'Resolve-AstroPhotoRoot'
        $scriptText | Should Match 'SharpCap\\Totality-2026'
        $scriptText | Should Match 'HDR\\Aligned'
        $scriptText | Should Match 'channel_match_offsets\.json'
        $scriptText | Should Match 'Use saved ChannelMatch offsets'
        $scriptText | Should Match 'New Instance triangle'
    }
}

Describe "Totality HDR entry point naming" {
    It "uses HDRTotality.ps1 and no longer keeps TotalityHDR.ps1" {
        Test-Path -LiteralPath (
            Join-Path -Path $PSScriptRoot -ChildPath "..\HDRTotality.ps1"
        ) -PathType Leaf | Should Be $true
        Test-Path -LiteralPath (
            Join-Path -Path $PSScriptRoot -ChildPath "..\TotalityHDR.ps1"
        ) | Should Be $false
    }
}
